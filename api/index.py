from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, HTMLResponse, RedirectResponse
from starlette.middleware.sessions import SessionMiddleware
from pydantic import BaseModel
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import mimetypes
import re
import traceback
import io
from dotenv import load_dotenv
from api.core.processor import convert_docx_to_html, inject_variables, resolve_gmail_font
from api.core.gmail_svc import (
    create_draft,
    exchange_code_for_token,
    get_auth_url,
    load_user_credentials,
    revoke_user_credentials,
)
import os
load_dotenv()
port = int(os.environ.get("PORT", 6311))
SESSION_SECRET = os.environ.get("SESSION_SECRET")
if not SESSION_SECRET:
    print("[CRITICAL] Missing environment variable: SESSION_SECRET")
    # 給予一個暫時的預設值，讓 API 至少能啟動以便記錄日誌
    SESSION_SECRET = "fallback_secret_please_set_this_in_vercel"

print(f"[BOOT] Environment: {'Vercel' if os.getenv('VERCEL') == '1' else 'Local'}")
print(f"[BOOT] CORS Origins: {os.environ.get('CORS_ORIGINS', 'Default')}")

app = FastAPI()
cors_origins = [
    origin.strip()
    for origin in os.environ.get("CORS_ORIGINS", "http://localhost:6406").split(",")
    if origin.strip()
]
app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    same_site="lax",
    https_only=os.environ.get("SESSION_HTTPS_ONLY", "false").lower() == "true",
)

@app.get("/")
def read_root():
    return {"status": "ok"}

@app.get("/api/health")
def health_check():
    return {"status": "ok", "message": "Gmail Replicator API is running"}

@app.get("/")
def root_login_page():
    html = """<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>DraftCopier Login</title>
  </head>
  <body>
    <h1>DraftCopier</h1>
    <p>Click below to start Google OAuth.</p>
    <button id="loginBtn">Login with Google</button>
    <script>
      const btn = document.getElementById("loginBtn");
      btn.addEventListener("click", async () => {
        await fetch("/api/dev/login", { method: "POST" });
        window.location.href = "/api/auth/google";
      });
    </script>
  </body>
</html>
"""
    return HTMLResponse(content=html)

@app.get("/api/dev/login")
def dev_login_get(request: Request):
    request.session["user_key"] = "dev_user"
    return {"ok": True, "user_key": "dev_user", "method": "GET"}

@app.post("/api/dev/login")
def dev_login(request: Request):
    request.session["user_key"] = "dev_user"
    return {"ok": True, "user_key": "dev_user"}

@app.post("/api/dev/logout")
def dev_logout(request: Request):
    request.session.pop("user_key", None)
    request.session.pop("oauth_state", None)
    request.session.pop("oauth_user_key", None)
    return {"ok": True}

class DraftRequest(BaseModel):
    to: str
    subject: str
    body_html: str

EMAIL_FIELD_CANDIDATES = {
    "email",
    "e-mail",
    "mail",
    "email address",
    "e-mail address",
    "電子郵件",
    "信箱",
    "收件人",
    "收件人信箱",
}

SUBJECT_FIELD_CANDIDATES = {
    "subject",
    "email subject",
    "mail subject",
    "title",
    "subject line",
    "主旨",
    "標題",
    "信件主旨",
}

ATTACHMENT_HEADER_PREFIX = "附件"
ATTACHMENTS_DIR = Path(os.environ.get("ATTACHMENTS_DIR", "attachments")).resolve()
ALLOW_ABSOLUTE_ATTACHMENTS = os.environ.get("ATTACHMENTS_ALLOW_ABSOLUTE", "true").lower() == "true"
ATTACHMENTS_ROOTS = [
    Path(p).resolve()
    for p in os.environ.get("ATTACHMENTS_ROOTS", str(ATTACHMENTS_DIR)).split(",")
    if p.strip()
]


def _normalize_header(value: object) -> str:
    return str(value).strip()


def _find_header(headers: list[object], candidates: set[str]) -> str | None:
    normalized = [_normalize_header(h) for h in headers]
    lower_map = {h.lower(): h for h in normalized}
    for candidate in candidates:
        if candidate in lower_map:
            return lower_map[candidate]
    for header in normalized:
        lower = header.lower()
        if any(candidate in lower for candidate in candidates):
            return header
    return None


def _find_attachment_headers(headers: list[object]) -> list[object]:
    matched = []
    for header in headers:
        name = _normalize_header(header)
        if name.startswith(ATTACHMENT_HEADER_PREFIX):
            matched.append(header)
    return matched


def _split_attachment_names(value: object) -> list[str]:
    if value is None:
        return []
    if isinstance(value, float) and pd.isna(value):
        return []
    text = str(value).strip()
    if not text:
        return []
    # Allow multiple filenames separated by comma/semicolon/newline
    parts = [p.strip() for p in re.split(r"[;,\n]+", text)]
    return [p for p in parts if p and str(p).lower() not in ('nan', 'none', 'null')]


def _is_within_roots(path: Path, roots: list[Path]) -> bool:
    for root in roots:
        if root in path.parents or path == root:
            return True
    return False


def _resolve_attachment_from_disk(
    name: str,
    base_dir: Path,
    allow_absolute: bool,
    roots: list[Path],
) -> dict[str, str | bytes] | None:
    if not name:
        return None
    candidate_path = Path(name)
    if candidate_path.is_absolute():
        if not allow_absolute:
            return None
        candidate = candidate_path.resolve()
    else:
        candidate = (base_dir / name).resolve()
        if not _is_within_roots(candidate, roots):
            return None
    if not candidate.exists() or not candidate.is_file():
        return None
    content = candidate.read_bytes()
    mime_type = mimetypes.guess_type(candidate.name)[0] or "application/octet-stream"
    return {"filename": candidate.name, "content": content, "mime_type": mime_type}

def _require_session_user_key(request: Request) -> str:
    user_key = request.session.get("user_key")
    if not user_key:
        raise HTTPException(status_code=401, detail="Missing session user_key")
    return user_key

@app.get("/api/auth/google")
def google_auth(request: Request):
    try:
        user_key = _require_session_user_key(request)
        auth_url, state = get_auth_url()
        request.session["oauth_state"] = state
        request.session["oauth_user_key"] = user_key
        return {"auth_url": auth_url}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/auth/callback/google")
def google_auth_callback(request: Request, code: str, state: str):
    try:
        print("callback state(query) =", state)
        print("callback session =", dict(request.session))
        expected_state = request.session.get("oauth_state")
        if not expected_state or state != expected_state:
            raise HTTPException(status_code=400, detail="Invalid OAuth state")
        user_key = request.session.get("oauth_user_key")
        if not user_key:
            raise HTTPException(status_code=401, detail="Missing session user_key")
        creds = exchange_code_for_token(code=code, state=state, user_key=user_key)
        request.session.pop("oauth_state", None)
        request.session.pop("oauth_user_key", None)
        front_base = os.environ.get("FRONTEND_BASE_URL", "http://localhost:6406")
        redirect_url = f"{front_base}/?auth=success"
        return RedirectResponse(url=redirect_url, status_code=302)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/auth/revoke")
def google_auth_revoke(request: Request):
    try:
        user_key = _require_session_user_key(request)
        revoke_user_credentials(user_key)
        return {"status": "ok", "revoked": True}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/drafts")
def create_draft_route(request: Request, payload: DraftRequest):
    try:
        user_key = _require_session_user_key(request)
        creds = load_user_credentials(user_key)
        draft = create_draft(
            creds=creds,
            to=payload.to,
            subject=payload.subject,
            body_html=payload.body_html,
        )
        return {"status": "ok", "draft": draft}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/drafts/batch")
async def create_drafts_batch(
    request: Request,
    docx_file: UploadFile = File(...),
    xlsx_file: UploadFile = File(...),
    sheet: str | None = None,
    font: str | None = None,
    attachments_dir: str | None = None,
):
    try:
        user_key = _require_session_user_key(request)
        creds = load_user_credentials(user_key)

        docx_content = await docx_file.read()
        font_family = resolve_gmail_font(font)
        html_template = convert_docx_to_html(docx_content, base_font_family=font_family)

        xlsx_content = await xlsx_file.read()
        sheet_name: int | str = 0
        if sheet is not None:
            sheet = sheet.strip()
            if sheet:
                sheet_name = int(sheet) if sheet.isdigit() else sheet
        df = pd.read_excel(io.BytesIO(xlsx_content), sheet_name=sheet_name)

        email_header = _find_header(list(df.columns), EMAIL_FIELD_CANDIDATES)
        subject_header = _find_header(list(df.columns), SUBJECT_FIELD_CANDIDATES)
        # attachment_headers = [str(h) for h in _find_attachment_headers(list(df.columns))]
        missing_headers = []
        if not email_header:
            missing_headers.append("email")
        if not subject_header:
            missing_headers.append("subject")
        if missing_headers:
            raise HTTPException(
                status_code=400,
                detail={
                    "error": "missing_required_headers",
                    "missing": missing_headers,
                },
            )

        rows = df.to_dict(orient="records")
        errors = []
        attachment_headers_raw = _find_attachment_headers(list(df.columns))
        base_dir = Path(attachments_dir).expanduser().resolve() if attachments_dir else ATTACHMENTS_DIR
        
        print(f"[DEBUG] Excel 欄位列表: {list(df.columns)}")
        print(f"[DEBUG] 偵測到的附件欄位 (attachment_headers_raw): {attachment_headers_raw}")
        print(f"[DEBUG] 附件根目錄 (base_dir): {base_dir} (存在: {base_dir.exists()})")
        
        row_attachments: list[list[dict[str, str | bytes]]] = []
        for idx, row in enumerate(rows, start=1):
            to_value = row.get(email_header)
            subject_value = row.get(subject_header)
            if to_value is None or (isinstance(to_value, float) and pd.isna(to_value)) or str(to_value).strip() == "":
                errors.append({"row": idx, "field": "email"})
            if subject_value is None or (isinstance(subject_value, float) and pd.isna(subject_value)) or str(subject_value).strip() == "":
                errors.append({"row": idx, "field": "subject"})

            attachments_for_row: list[dict[str, str | bytes]] = []
            for header in attachment_headers_raw:
                value = row.get(header)
                if idx <= 3:
                    print(f"[attachments] row {idx} header={header!r} type={type(value)} value={value!r}")
                for name in _split_attachment_names(value):
                    if not Path(name).is_absolute() and not base_dir.exists():
                        print(f"[ERROR] 找不到附件根目錄: {base_dir} (嘗試讀取相對路徑附件 '{name}')")
                        raise HTTPException(
                            status_code=400,
                            detail={
                                "error": "missing_attachments_dir",
                                "message": f"表格中指定了相對路徑附件 '{name}'，但附件資料夾 {base_dir} 不存在。請建立該資料夾並放入檔案。",
                            },
                        )
                    resolved = _resolve_attachment_from_disk(
                        name=name,
                        base_dir=base_dir,
                        allow_absolute=ALLOW_ABSOLUTE_ATTACHMENTS,
                        roots=ATTACHMENTS_ROOTS,
                    )
                    if not resolved:
                        print(f"[ERROR] 找不到附件檔案: {name} (位於第 {idx} 列)")
                        errors.append({"row": idx, "field": "attachment", "name": name})
                    else:
                        attachments_for_row.append(resolved)
            row_attachments.append(attachments_for_row)

        if errors:
            print(f"[ERROR] 批次建檔失敗，以下資料有缺失：{errors}")
            raise HTTPException(
                status_code=400,
                detail={
                    "error": "missing_required_values",
                    "missing": errors,
                },
            )

        drafts = []
        for idx, row in enumerate(rows, start=1):
            to_value = row.get(email_header)
            subject_value = row.get(subject_header)
            body_html = inject_variables(html_template, row)
            draft = create_draft(
                creds=creds,
                to=str(to_value).strip(),
                subject=str(subject_value).strip(),
                body_html=body_html,
                attachments=row_attachments[idx - 1],
            )
            drafts.append(draft.get("id"))

        return {"status": "ok", "draft_count": len(drafts), "draft_ids": drafts}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/process")
async def process_files(
    docx_file: UploadFile = File(...),
    xlsx_file: UploadFile = File(...),
    sheet: str | None = None,
    font: str | None = None,
):
    try:
        # 1. 讀取 Word 模板並轉為 HTML
        docx_content = await docx_file.read()
        font_family = resolve_gmail_font(font)
        html_template = convert_docx_to_html(docx_content, base_font_family=font_family)

        # 2. 使用 Pandas 讀取 Excel 數據
        xlsx_content = await xlsx_file.read()
        sheet_name: int | str = 0
        if sheet is not None:
            sheet = sheet.strip()
            if sheet:
                sheet_name = int(sheet) if sheet.isdigit() else sheet
        df = pd.read_excel(io.BytesIO(xlsx_content), sheet_name=sheet_name)
        
        # 將 DataFrame 轉為字典列表，方便處理自定義欄位
        rows = df.to_dict(orient="records")
        attachment_headers = [str(h) for h in _find_attachment_headers(list(df.columns))]
        
        # 3. 產生第一筆預覽 (作為測試)
        preview = ""
        first_row: dict[str, str] = {}
        if rows:
            preview = inject_variables(html_template, rows[0])
            for key, value in rows[0].items():
                # Normalize to string-safe values for JSON output
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    first_row[str(key)] = ""
                elif isinstance(value, (datetime, date, pd.Timestamp)):
                    first_row[str(key)] = value.isoformat()
                else:
                    first_row[str(key)] = str(value)

        email_header = _find_header(list(df.columns), EMAIL_FIELD_CANDIDATES)
        subject_header = _find_header(list(df.columns), SUBJECT_FIELD_CANDIDATES)

        return {
            "total_records": len(rows),
            "headers": list(df.columns),
            "preview_first_row": preview,
            "first_row": first_row,
            "detected_fields": {
                "email": email_header,
                "subject": subject_header,
                "attachments": attachment_headers,
            },
        }
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"error": str(e)})

# 為了讓 Vercel 以外的環境也能執行
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=port)
