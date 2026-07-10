from fastapi import FastAPI, UploadFile, File, Form, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, HTMLResponse, RedirectResponse, Response
from starlette.middleware.sessions import SessionMiddleware
from pydantic import BaseModel
import pandas as pd
from datetime import date, datetime
from pathlib import Path
import mimetypes
import json
import re
import traceback
import io
from dotenv import load_dotenv
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from api.core.processor import (
    export_html_to_docx_bytes,
    inject_variables,
    resolve_gmail_font,
    resolve_template_html,
)
from api.core.gmail_svc import (
    create_draft,
    exchange_code_for_token,
    get_auth_url,
    load_user_credentials,
    revoke_user_credentials,
)
import os
from uuid import uuid4
load_dotenv()
port = int(os.environ.get("PORT", 6311))
IS_PRODUCTION = os.getenv("VERCEL") == "1"

SESSION_SECRET = os.environ.get("SESSION_SECRET")
if not SESSION_SECRET:
    raise RuntimeError(
        "[CRITICAL] Missing required environment variable: SESSION_SECRET. "
        "Set it before starting the server."
    )

# 開發環境預設開啟 dev login，正式環境須明確設定才開啟
DEV_LOGIN_ENABLED = os.environ.get(
    "DEV_LOGIN_ENABLED", "false" if IS_PRODUCTION else "true"
).lower() == "true"

print(f"[BOOT] Environment: {'Production (Vercel)' if IS_PRODUCTION else 'Local'}")
print(f"[BOOT] DEV_LOGIN_ENABLED: {DEV_LOGIN_ENABLED}")
print(f"[BOOT] CORS Origins: {os.environ.get('CORS_ORIGINS', 'Default')}")

app = FastAPI()

# 正式環境：CORS_ORIGINS 必須明確設定，不提供預設值（未設定則封鎖所有跨域請求）
# 開發環境：預設允許常見的本機端口
_cors_default = "" if IS_PRODUCTION else "http://localhost:3000,http://localhost:6406"
cors_origins = [
    o.strip()
    for o in os.environ.get("CORS_ORIGINS", _cors_default).split(",")
    if o.strip()
]
if IS_PRODUCTION and not cors_origins:
    print("[WARN] CORS_ORIGINS is not set. All cross-origin requests will be blocked.")

# 正式環境收緊 methods/headers；開發環境保持寬鬆
cors_methods = ["GET", "POST", "OPTIONS"] if IS_PRODUCTION else ["*"]
cors_headers = ["Content-Type", "Authorization"] if IS_PRODUCTION else ["*"]

app.add_middleware(
    CORSMiddleware,
    allow_origins=cors_origins,
    allow_credentials=True,
    allow_methods=cors_methods,
    allow_headers=cors_headers,
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
        const res = await fetch("/api/auth/google");
        const data = await res.json();
        if (data.auth_url) window.location.href = data.auth_url;
      });
    </script>
  </body>
</html>
"""
    return HTMLResponse(content=html)

@app.get("/api/dev/login")
def dev_login_get(request: Request):
    if not DEV_LOGIN_ENABLED:
        raise HTTPException(status_code=404, detail="Not found")
    request.session["user_key"] = "dev_user"
    return {"ok": True, "user_key": "dev_user", "method": "GET"}

@app.post("/api/dev/login")
def dev_login(request: Request):
    if not DEV_LOGIN_ENABLED:
        raise HTTPException(status_code=404, detail="Not found")
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
    cc: str | None = None
    bcc: str | None = None

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

CC_FIELD_CANDIDATES = {
    "cc",
    "copy",
    "carbon copy",
    "抄送",
    "副本",
    "副本收件人",
}

BCC_FIELD_CANDIDATES = {
    "bcc",
    "blind carbon copy",
    "密件副本",
    "秘密副本",
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
UPLOAD_CACHE_DIR = Path(os.environ.get("UPLOAD_CACHE_DIR", "/tmp/draftcopier_uploads")).resolve()
GOOGLE_DOC_MIME = "application/vnd.google-apps.document"
GOOGLE_SHEET_MIME = "application/vnd.google-apps.spreadsheet"
GOOGLE_SLIDE_MIME = "application/vnd.google-apps.presentation"
GOOGLE_FOLDER_MIME = "application/vnd.google-apps.folder"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
XLS_MIME = "application/vnd.ms-excel"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
DRIVE_FILE_KINDS = {
    "docx": {
        "mime_types": {DOCX_MIME, GOOGLE_DOC_MIME},
        "export_mime": DOCX_MIME,
        "extension": ".docx",
    },
    "xlsx": {
        "mime_types": {XLSX_MIME, XLS_MIME, GOOGLE_SHEET_MIME},
        "export_mime": XLSX_MIME,
        "extension": ".xlsx",
    },
    "attachment": {
        "blocked_mime_types": {GOOGLE_FOLDER_MIME},
        "workspace_export": {
            GOOGLE_DOC_MIME: {"mime_type": DOCX_MIME, "extension": ".docx"},
            GOOGLE_SHEET_MIME: {"mime_type": XLSX_MIME, "extension": ".xlsx"},
            GOOGLE_SLIDE_MIME: {"mime_type": PPTX_MIME, "extension": ".pptx"},
        },
    },
    "folder": {
        "mime_types": {GOOGLE_FOLDER_MIME},
    },
}


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


def _normalize_email_list(value: object) -> str | None:
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    parts = [p.strip() for p in re.split(r"[;,\n]+", text)]
    normalized = [p for p in parts if p]
    return ", ".join(normalized) if normalized else None


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


def _safe_upload_name(name: str, fallback: str) -> str:
    base = Path(name or "").name.strip()
    return base or fallback


def _get_upload_cache_namespace(request: Request) -> str:
    namespace = request.session.get("upload_cache_ns")
    if isinstance(namespace, str) and namespace:
        return namespace
    namespace = uuid4().hex
    request.session["upload_cache_ns"] = namespace
    return namespace


def _cache_root_for_namespace(namespace: str) -> Path:
    safe_name = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in namespace)
    return UPLOAD_CACHE_DIR / safe_name


def _write_upload_cache(
    *,
    namespace: str,
    docx_name: str,
    docx_content: bytes,
    xlsx_name: str,
    xlsx_content: bytes,
) -> dict[str, str]:
    cache_id = uuid4().hex
    cache_root = _cache_root_for_namespace(namespace)
    cache_dir = cache_root / cache_id
    cache_dir.mkdir(parents=True, exist_ok=True)

    safe_docx_name = _safe_upload_name(docx_name, "template.docx")
    safe_xlsx_name = _safe_upload_name(xlsx_name, "list.xlsx")
    (cache_dir / "template.docx").write_bytes(docx_content)
    (cache_dir / "recipients.xlsx").write_bytes(xlsx_content)
    (cache_dir / "meta.txt").write_text(f"{safe_docx_name}\n{safe_xlsx_name}\n", encoding="utf-8")

    return {"cache_id": cache_id, "docx_name": safe_docx_name, "xlsx_name": safe_xlsx_name}


def _read_upload_cache(*, namespace: str, cache_id: str) -> dict[str, bytes | str]:
    if not cache_id or not re.fullmatch(r"[a-f0-9]{32}", cache_id):
        raise HTTPException(status_code=400, detail={"error": "invalid_cache_id"})
    cache_dir = _cache_root_for_namespace(namespace) / cache_id
    docx_path = cache_dir / "template.docx"
    xlsx_path = cache_dir / "recipients.xlsx"
    if not docx_path.exists() or not xlsx_path.exists():
        raise HTTPException(status_code=400, detail={"error": "cache_not_found", "cache_id": cache_id})

    docx_name = "template.docx"
    xlsx_name = "list.xlsx"
    meta_path = cache_dir / "meta.txt"
    if meta_path.exists():
        lines = meta_path.read_text(encoding="utf-8").splitlines()
        if len(lines) >= 2:
            docx_name = lines[0] or docx_name
            xlsx_name = lines[1] or xlsx_name

    return {
        "docx_name": docx_name,
        "xlsx_name": xlsx_name,
        "docx_content": docx_path.read_bytes(),
        "xlsx_content": xlsx_path.read_bytes(),
    }


def _build_preview_payload(
    *,
    docx_content: bytes | None,
    template_html: str | None,
    xlsx_content: bytes,
    sheet: str | None,
    font: str | None,
    cache_info: dict[str, str] | None,
    preview_page: int,
    preview_page_size: int,
) -> dict:
    font_family = resolve_gmail_font(font)
    html_template = resolve_template_html(
        docx_content=docx_content,
        template_html=template_html,
        base_font_family=font_family,
    )

    df, sheet_names, selected_sheet = _read_recipient_sheet(xlsx_content, sheet)

    def serialize_row(row: dict) -> dict[str, str]:
        serialized: dict[str, str] = {}
        for key, value in row.items():
            if value is None or (isinstance(value, float) and pd.isna(value)):
                serialized[str(key)] = ""
            elif isinstance(value, (datetime, date, pd.Timestamp)):
                serialized[str(key)] = value.isoformat()
            else:
                serialized[str(key)] = str(value)
        return serialized

    rows = df.to_dict(orient="records")
    attachment_headers = [str(h) for h in _find_attachment_headers(list(df.columns))]
    total_records = len(rows)
    page_size = max(1, min(preview_page_size, 200))
    total_pages = max(1, (total_records + page_size - 1) // page_size) if total_records else 1
    page = max(1, min(preview_page, total_pages))
    page_start = (page - 1) * page_size
    page_end = page_start + page_size

    preview = ""
    first_row: dict[str, str] = {}
    if rows:
        preview = inject_variables(html_template, rows[0])
        first_row = serialize_row(rows[0])

    email_header = _find_header(list(df.columns), EMAIL_FIELD_CANDIDATES)
    cc_header = _find_header(list(df.columns), CC_FIELD_CANDIDATES)
    bcc_header = _find_header(list(df.columns), BCC_FIELD_CANDIDATES)
    subject_header = _find_header(list(df.columns), SUBJECT_FIELD_CANDIDATES)
    payload = {
        "total_records": total_records,
        "headers": list(df.columns),
        "preview_first_row": preview,
        "template_html": html_template,
        "first_row": first_row,
        "preview_rows": [serialize_row(row) for row in rows[page_start:page_end]],
        "preview_pagination": {
            "page": page,
            "page_size": page_size,
            "total_pages": total_pages,
        },
        "sheet_names": sheet_names,
        "selected_sheet": selected_sheet,
        "detected_fields": {
            "email": email_header,
            "cc": cc_header,
            "bcc": bcc_header,
            "subject": subject_header,
            "attachments": attachment_headers,
        },
    }
    if cache_info:
        payload["cache_id"] = cache_info["cache_id"]
        payload["cached_files"] = {
            "docx_name": cache_info["docx_name"],
            "xlsx_name": cache_info["xlsx_name"],
        }
    return payload


def get_attachment_content(
    file_path: str,
    credentials: Credentials,
    *,
    drive_service=None,
    parent_folder_id: str | None = None,
) -> dict[str, str | bytes]:
    """
    Google Drive fallback：依檔名從 Drive 查詢並下載。
    本地檔案解析須透過 _resolve_attachment_from_disk（含根目錄驗證），不在此處理。
    """
    filename = os.path.basename(file_path)
    if not filename:
        raise FileNotFoundError(f"Invalid attachment path: {file_path!r}")

    if drive_service is None:
        drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
    query_name = filename.replace("'", "\\'")
    clauses = [f"name='{query_name}'", "trashed=false"]
    if parent_folder_id:
        safe_parent = parent_folder_id.replace("'", "\\'")
        clauses.append(f"'{safe_parent}' in parents")
    query = " and ".join(clauses)
    results = drive_service.files().list(
        q=query,
        fields="files(id,name,mimeType)",
        pageSize=1,
    ).execute()
    items = results.get("files", [])
    if not items and parent_folder_id:
        return get_attachment_content(
            file_path,
            credentials,
            drive_service=drive_service,
            parent_folder_id=None,
        )
    if not items:
        raise FileNotFoundError(f"Attachment not found on disk or Google Drive: {filename}")

    item = items[0]
    request = drive_service.files().get_media(fileId=item["id"])
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    mime_type = item.get("mimeType") or mimetypes.guess_type(filename)[0] or "application/octet-stream"
    return {"filename": item.get("name", filename), "content": fh.getvalue(), "mime_type": mime_type}


def _read_recipient_sheet(xlsx_content: bytes, sheet: str | None) -> tuple[pd.DataFrame, list[str], str]:
    workbook = pd.ExcelFile(io.BytesIO(xlsx_content))
    sheet_names = list(workbook.sheet_names)
    if not sheet_names:
        raise HTTPException(status_code=400, detail={"error": "empty_workbook"})

    selected_sheet = sheet_names[0]
    requested_sheet = (sheet or "").strip()
    if requested_sheet:
        if requested_sheet.isdigit():
            index = int(requested_sheet)
            if index < 0 or index >= len(sheet_names):
                raise HTTPException(
                    status_code=400,
                    detail={"error": "invalid_sheet", "message": f"Unknown sheet index: {index}"},
                )
            selected_sheet = sheet_names[index]
        elif requested_sheet in sheet_names:
            selected_sheet = requested_sheet
        else:
            raise HTTPException(
                status_code=400,
                detail={"error": "invalid_sheet", "message": f"Unknown sheet: {requested_sheet}"},
            )

    df = workbook.parse(sheet_name=selected_sheet)
    return df, sheet_names, selected_sheet


def _get_drive_kind_config(kind: str) -> dict[str, object]:
    config = DRIVE_FILE_KINDS.get(kind)
    if not config:
        raise HTTPException(status_code=400, detail={"error": "invalid_drive_kind", "kind": kind})
    return config


def _kind_allows_mime_type(kind: str, mime_type: str) -> bool:
    config = _get_drive_kind_config(kind)
    if "mime_types" in config:
        return mime_type in config["mime_types"]  # type: ignore[operator]
    blocked = config.get("blocked_mime_types", set())
    if mime_type in blocked:  # type: ignore[operator]
        return False
    if mime_type.startswith("application/vnd.google-apps."):
        workspace_export = config.get("workspace_export", {})
        return mime_type in workspace_export  # type: ignore[operator]
    return True


def _derive_google_app_id(client_id: str | None) -> str | None:
    if not client_id:
        return None
    prefix = client_id.split("-", 1)[0].strip()
    return prefix if prefix.isdigit() else None


def _drive_query_for_kind(kind: str, query: str | None) -> str:
    config = _get_drive_kind_config(kind)
    if "mime_types" in config:
        mime_filters = " or ".join(
            f"mimeType='{mime}'" for mime in sorted(config["mime_types"])  # type: ignore[index]
        )
        clauses = [f"({mime_filters})", "trashed=false"]
    else:
        blocked = sorted(config.get("blocked_mime_types", set()))  # type: ignore[arg-type]
        clauses = [f"mimeType!='{mime}'" for mime in blocked]
        clauses.append("trashed=false")
    if query:
        safe_query = query.replace("\\", "\\\\").replace("'", "\\'")
        clauses.append(f"name contains '{safe_query}'")
    return " and ".join(clauses)


def _download_drive_file(
    drive_service,
    *,
    file_id: str,
    kind: str,
) -> dict[str, str | bytes]:
    config = _get_drive_kind_config(kind)
    metadata = drive_service.files().get(
        fileId=file_id,
        fields="id,name,mimeType,modifiedTime",
    ).execute()
    mime_type = str(metadata.get("mimeType") or "")
    if not _kind_allows_mime_type(kind, mime_type):
        raise HTTPException(
            status_code=400,
            detail={"error": "unsupported_drive_file", "file_id": file_id, "mime_type": mime_type},
        )

    filename = str(metadata.get("name") or f"{kind}{config.get('extension', '')}")
    if kind == "attachment" and mime_type.startswith("application/vnd.google-apps."):
        workspace_export = config.get("workspace_export", {})
        export_config = workspace_export.get(mime_type)  # type: ignore[assignment]
        if not export_config:
            raise HTTPException(
                status_code=400,
                detail={"error": "unsupported_drive_file", "file_id": file_id, "mime_type": mime_type},
            )
        extension = str(export_config["extension"])
        if not filename.lower().endswith(extension):
            filename = f"{filename}{extension}"
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType=str(export_config["mime_type"]),
        )
        output_mime = str(export_config["mime_type"])
    elif mime_type.startswith("application/vnd.google-apps."):
        if not filename.lower().endswith(str(config["extension"])):
            filename = f"{filename}{config['extension']}"
        request = drive_service.files().export_media(
            fileId=file_id,
            mimeType=str(config["export_mime"]),
        )
        output_mime = str(config["export_mime"])
    else:
        request = drive_service.files().get_media(fileId=file_id)
        output_mime = mime_type

    buffer = io.BytesIO()
    downloader = MediaIoBaseDownload(buffer, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()

    return {
        "file_id": str(metadata.get("id") or file_id),
        "name": filename,
        "mime_type": output_mime,
        "content": buffer.getvalue(),
    }


def _parse_attachment_drive_file_ids(raw: str | None) -> list[str]:
    if raw is None or not raw.strip():
        return []
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise HTTPException(
            status_code=400,
            detail={"error": "invalid_attachment_drive_file_ids", "message": str(exc)},
        ) from exc
    if not isinstance(parsed, list) or not all(isinstance(item, str) and item.strip() for item in parsed):
        raise HTTPException(
            status_code=400,
            detail={"error": "invalid_attachment_drive_file_ids", "message": "Expected a JSON array of file ids."},
        )
    return [item.strip() for item in parsed]


def _load_selected_drive_attachments(
    drive_service,
    file_ids: list[str],
) -> dict[str, dict[str, str | bytes]]:
    selected: dict[str, dict[str, str | bytes]] = {}
    for file_id in file_ids:
        downloaded = _download_drive_file(drive_service, file_id=file_id, kind="attachment")
        filename = str(downloaded["name"])
        attachment = {
            "filename": filename,
            "content": downloaded["content"],
            "mime_type": str(downloaded["mime_type"]),
        }
        selected[filename] = attachment
        selected[Path(filename).name] = attachment
    return selected


async def _load_uploaded_local_attachments(
    upload_files: list[UploadFile] | None,
) -> dict[str, dict[str, str | bytes]]:
    selected: dict[str, dict[str, str | bytes]] = {}
    for upload_file in upload_files or []:
        filename = upload_file.filename or "attachment"
        content = await upload_file.read()
        attachment = {
            "filename": filename,
            "content": content,
            "mime_type": upload_file.content_type or mimetypes.guess_type(filename)[0] or "application/octet-stream",
        }
        selected[filename] = attachment
        selected[Path(filename).name] = attachment
    return selected


async def _read_input_source(
    *,
    kind: str,
    upload_file: UploadFile | None,
    drive_file_id: str | None,
    drive_service,
) -> tuple[str, bytes]:
    if upload_file:
        return upload_file.filename or f"{kind}.{kind}", await upload_file.read()
    if drive_file_id:
        downloaded = _download_drive_file(drive_service, file_id=drive_file_id, kind=kind)
        name = downloaded["name"]
        content = downloaded["content"]
        if not isinstance(name, str) or not isinstance(content, bytes):
            raise HTTPException(status_code=500, detail={"error": "invalid_drive_download"})
        return name, content
    raise HTTPException(
        status_code=400,
        detail={"error": "missing_upload_source", "message": f"Need {kind} file or drive file id."},
    )


def _require_session_user_key(request: Request) -> str:
    user_key = request.session.get("user_key")
    if not user_key:
        raise HTTPException(status_code=401, detail="Missing session user_key")
    return user_key


@app.get("/api/session/info")
def session_info(request: Request):
    user_key = request.session.get("user_key")
    email = request.session.get("user_email")
    name = request.session.get("user_name")
    return {"user_key": user_key, "email": email, "name": name}


def _fetch_gmail_email(creds: Credentials) -> str | None:
    try:
        service = build("gmail", "v1", credentials=creds, cache_discovery=False)
        profile = service.users().getProfile(userId="me").execute()
        return profile.get("emailAddress")
    except Exception:
        return None


def _fetch_google_profile(creds: Credentials) -> dict[str, str | None]:
    profile: dict[str, str | None] = {"email": None, "name": None}
    try:
        service = build("oauth2", "v2", credentials=creds, cache_discovery=False)
        userinfo = service.userinfo().get().execute()
        email = userinfo.get("email")
        name = userinfo.get("name")
        profile["email"] = str(email) if email else None
        profile["name"] = str(name) if name else None
    except Exception:
        profile["email"] = _fetch_gmail_email(creds)
    return profile


class RestoreSessionRequest(BaseModel):
    user_key: str


@app.post("/api/session/restore")
def session_restore(request: Request, body: RestoreSessionRequest):
    if not re.fullmatch(r"[a-f0-9]{32}", body.user_key):
        raise HTTPException(status_code=400, detail="Invalid user_key")
    try:
        creds = load_user_credentials(body.user_key)
        request.session["user_key"] = body.user_key
        email = request.session.get("user_email")
        name = request.session.get("user_name")
        if not email or not name:
            profile = _fetch_google_profile(creds)
            email = email or profile.get("email")
            name = name or profile.get("name")
            if email:
                request.session["user_email"] = email
            if name:
                request.session["user_name"] = name
        return {"ok": True, "restored": True, "email": email, "name": name}
    except Exception:
        return {"ok": True, "restored": False, "email": None, "name": None}


@app.get("/api/auth/google")
def google_auth(request: Request):
    try:
        user_key = request.session.get("user_key")
        if not user_key:
            user_key = uuid4().hex
            request.session["user_key"] = user_key
        print(f"[auth/google] user_key={user_key!r}")
        print(f"[auth/google] GOOGLE_CLIENT_ID set: {bool(os.environ.get('GOOGLE_CLIENT_ID'))}")
        print(f"[auth/google] GOOGLE_CLIENT_SECRET set: {bool(os.environ.get('GOOGLE_CLIENT_SECRET'))}")
        print(f"[auth/google] GOOGLE_REDIRECT_URI={os.environ.get('GOOGLE_REDIRECT_URI')!r}")
        auth_url, state, code_verifier = get_auth_url()
        request.session["oauth_state"] = state
        request.session["oauth_user_key"] = user_key
        if code_verifier:
            request.session["oauth_code_verifier"] = code_verifier
        return {"auth_url": auth_url}
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
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
            
        code_verifier = request.session.get("oauth_code_verifier")
        
        creds = exchange_code_for_token(code=code, state=state, user_key=user_key, code_verifier=code_verifier)
        profile = _fetch_google_profile(creds)
        email = profile.get("email")
        name = profile.get("name")
        if email:
            request.session["user_email"] = email
        if name:
            request.session["user_name"] = name

        request.session.pop("oauth_state", None)
        request.session.pop("oauth_user_key", None)
        request.session.pop("oauth_code_verifier", None)
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


@app.get("/api/drive/files")
def list_drive_files(
    request: Request,
    kind: str,
    q: str | None = None,
):
    try:
        user_key = _require_session_user_key(request)
        creds = load_user_credentials(user_key)
        drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
        response = drive_service.files().list(
            q=_drive_query_for_kind(kind, q),
            fields="files(id,name,mimeType,modifiedTime)",
            orderBy="modifiedTime desc",
            pageSize=20,
        ).execute()
        config = _get_drive_kind_config(kind)
        items = []
        for file in response.get("files", []):
            mime_type = str(file.get("mimeType") or "")
            if not _kind_allows_mime_type(kind, mime_type):
                continue
            items.append(
                {
                    "id": file.get("id"),
                    "name": file.get("name"),
                    "mime_type": mime_type,
                    "modified_time": file.get("modifiedTime"),
                    "kind": kind,
                    "is_google_workspace": mime_type.startswith("application/vnd.google-apps."),
                }
            )
        return {"files": items}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/google/picker/config")
def google_picker_config(request: Request):
    client_id = os.environ.get("GOOGLE_PICKER_CLIENT_ID") or os.environ.get("GOOGLE_CLIENT_ID")
    api_key = os.environ.get("GOOGLE_PICKER_API_KEY", "").strip()
    app_id = os.environ.get("GOOGLE_PICKER_APP_ID") or _derive_google_app_id(client_id)
    enabled = bool(client_id and api_key and app_id)
    return {
        "enabled": enabled,
        "client_id": client_id,
        "api_key": api_key,
        "app_id": app_id,
        "scope": "https://www.googleapis.com/auth/drive.file",
        "login_hint": request.session.get("user_email"),
    }


class DriveFolderCreateRequest(BaseModel):
    parent_folder_id: str | None = None
    folder_name: str | None = None


@app.post("/api/drive/folders/create")
def create_drive_folder(request: Request, payload: DriveFolderCreateRequest):
    try:
        user_key = _require_session_user_key(request)
        creds = load_user_credentials(user_key)
        drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
        folder_name = (payload.folder_name or "").strip() or f"DraftCopier Attachments {datetime.now().strftime('%Y%m%d-%H%M%S')}"
        body: dict[str, object] = {
            "name": folder_name,
            "mimeType": GOOGLE_FOLDER_MIME,
        }
        if payload.parent_folder_id:
            body["parents"] = [payload.parent_folder_id]
        folder = drive_service.files().create(
            body=body,
            fields="id,name,mimeType,modifiedTime",
        ).execute()
        return {
            "folder": {
                "id": folder.get("id"),
                "name": folder.get("name"),
                "mime_type": folder.get("mimeType"),
                "modified_time": folder.get("modifiedTime"),
                "kind": "folder",
                "is_google_workspace": True,
            }
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/drive/folders/upload")
async def upload_files_to_drive_folder(
    request: Request,
    folder_id: str = Form(...),
    files: list[UploadFile] | None = File(None),
):
    try:
        if not files:
            raise HTTPException(status_code=400, detail={"error": "missing_files"})
        user_key = _require_session_user_key(request)
        creds = load_user_credentials(user_key)
        drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
        uploaded_items: list[dict[str, object]] = []
        for upload_file in files:
            filename = upload_file.filename or "attachment"
            content = await upload_file.read()
            media = MediaIoBaseUpload(
                io.BytesIO(content),
                mimetype=upload_file.content_type or mimetypes.guess_type(filename)[0] or "application/octet-stream",
                resumable=False,
            )
            file = drive_service.files().create(
                body={"name": filename, "parents": [folder_id]},
                media_body=media,
                fields="id,name,mimeType,modifiedTime",
            ).execute()
            uploaded_items.append(
                {
                    "id": file.get("id"),
                    "name": file.get("name"),
                    "mime_type": file.get("mimeType"),
                    "modified_time": file.get("modifiedTime"),
                    "kind": "attachment",
                    "is_google_workspace": False,
                }
            )
        return {"files": uploaded_items}
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
            cc=payload.cc,
            bcc=payload.bcc,
        )
        return {"status": "ok", "draft": draft}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/drafts/batch")
async def create_drafts_batch(
    request: Request,
    docx_file: UploadFile | None = File(None),
    xlsx_file: UploadFile | None = File(None),
    cache_id: str | None = Form(None),
    docx_drive_file_id: str | None = Form(None),
    xlsx_drive_file_id: str | None = Form(None),
    template_html: str | None = Form(None),
    attachment_drive_file_ids_json: str | None = Form(None),
    attachment_drive_folder_id: str | None = Form(None),
    attachment_local_files: list[UploadFile] | None = File(None),
    sheet: str | None = None,
    font: str | None = None,
    attachments_dir: str | None = None,
):
    try:
        user_key = _require_session_user_key(request)
        creds = load_user_credentials(user_key)
        cache_namespace = _get_upload_cache_namespace(request)
        attachment_drive_file_ids = _parse_attachment_drive_file_ids(attachment_drive_file_ids_json)
        selected_drive_attachments: dict[str, dict[str, str | bytes]] = {}
        selected_local_attachments = await _load_uploaded_local_attachments(attachment_local_files)

        docx_content: bytes | None = None
        drive_service_for_inputs = None
        if docx_drive_file_id or xlsx_drive_file_id:
            drive_service_for_inputs = build("drive", "v3", credentials=creds, cache_discovery=False)

        if docx_file and xlsx_file:
            docx_content = await docx_file.read()
            xlsx_content = await xlsx_file.read()
        elif cache_id:
            cached = _read_upload_cache(namespace=cache_namespace, cache_id=cache_id)
            docx_content = cached["docx_content"]
            xlsx_content = cached["xlsx_content"]
            if not isinstance(docx_content, bytes) or not isinstance(xlsx_content, bytes):
                raise HTTPException(status_code=500, detail={"error": "invalid_cache_payload"})
        elif xlsx_file and template_html:
            xlsx_content = await xlsx_file.read()
        elif xlsx_drive_file_id and template_html:
            _, xlsx_content = await _read_input_source(
                kind="xlsx",
                upload_file=None,
                drive_file_id=xlsx_drive_file_id,
                drive_service=drive_service_for_inputs,
            )
            if docx_drive_file_id:
                _, docx_content = await _read_input_source(
                    kind="docx",
                    upload_file=None,
                    drive_file_id=docx_drive_file_id,
                    drive_service=drive_service_for_inputs,
                )
        else:
            raise HTTPException(
                status_code=400,
                detail={
                    "error": "missing_upload_source",
                    "message": "Need cache_id, docx/xlsx files, or xlsx source with template_html.",
                },
            )

        font_family = resolve_gmail_font(font)
        html_template = resolve_template_html(
            docx_content=docx_content,
            template_html=template_html,
            base_font_family=font_family,
        )
        df, _, _ = _read_recipient_sheet(xlsx_content, sheet)

        email_header = _find_header(list(df.columns), EMAIL_FIELD_CANDIDATES)
        cc_header = _find_header(list(df.columns), CC_FIELD_CANDIDATES)
        bcc_header = _find_header(list(df.columns), BCC_FIELD_CANDIDATES)
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
        failed_items: list[dict[str, object]] = []
        attachment_headers_raw = _find_attachment_headers(list(df.columns))
        base_dir = Path(attachments_dir).expanduser().resolve() if attachments_dir else ATTACHMENTS_DIR
        drive_service = None
        if attachment_drive_file_ids or attachment_drive_folder_id:
            drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)
        if attachment_drive_file_ids:
            selected_drive_attachments = _load_selected_drive_attachments(drive_service, attachment_drive_file_ids)
        
        print(f"[DEBUG] Excel 欄位列表: {list(df.columns)}")
        print(f"[DEBUG] 偵測到的附件欄位 (attachment_headers_raw): {attachment_headers_raw}")
        print(f"[DEBUG] 附件根目錄 (base_dir): {base_dir} (存在: {base_dir.exists()})")
        
        row_attachments: list[list[dict[str, str | bytes]]] = []
        row_can_send: list[bool] = []

        def _is_empty_cell(value: object) -> bool:
            return value is None or (isinstance(value, float) and pd.isna(value)) or str(value).strip() == ""

        for idx, row in enumerate(rows, start=1):
            to_value = row.get(email_header)
            cc_value = _normalize_email_list(row.get(cc_header)) if cc_header else None
            bcc_value = _normalize_email_list(row.get(bcc_header)) if bcc_header else None
            subject_value = row.get(subject_header)
            can_send = True
            if _is_empty_cell(to_value):
                failed_items.append({"row": idx, "field": "email", "message": "Missing recipient email"})
                can_send = False
            if _is_empty_cell(subject_value):
                failed_items.append({"row": idx, "field": "subject", "message": "Missing email subject"})
                can_send = False
            row_can_send.append(can_send)

            attachments_for_row: list[dict[str, str | bytes]] = []
            for header in attachment_headers_raw:
                value = row.get(header)
                if idx <= 3:
                    print(f"[attachments] row {idx} header={header!r} type={type(value)} value={value!r}")
                for name in _split_attachment_names(value):
                    resolved: dict[str, str | bytes] | None = None
                    selected_attachment = selected_drive_attachments.get(name) or selected_drive_attachments.get(Path(name).name)
                    if selected_attachment:
                        resolved = selected_attachment
                    if not resolved:
                        resolved = selected_local_attachments.get(name) or selected_local_attachments.get(Path(name).name)
                    name_path = Path(name)
                    should_try_local = name_path.is_absolute() or base_dir.exists()
                    if not resolved and should_try_local:
                        resolved = _resolve_attachment_from_disk(
                            name=name,
                            base_dir=base_dir,
                            allow_absolute=ALLOW_ABSOLUTE_ATTACHMENTS,
                            roots=ATTACHMENTS_ROOTS,
                        )
                    elif not resolved and idx <= 3:
                        print(
                            f"[attachments] base_dir 不存在，略過本地相對路徑，改用 Drive 查詢: row={idx} name={name!r}"
                        )
                    if not resolved:
                        print(f"[attachments] Drive fallback lookup: row={idx} name={name!r}")
                        try:
                            resolved = get_attachment_content(
                                name,
                                creds,
                                drive_service=drive_service,
                                parent_folder_id=attachment_drive_folder_id,
                            )
                            if idx <= 3:
                                print(
                                    f"[attachments] Drive fallback success: row={idx} name={name!r} resolved={resolved.get('filename')!r}"
                                )
                        except Exception as e:
                            print(f"[ERROR] 找不到附件檔案: {name} (位於第 {idx} 列) err={e}")
                            failed_items.append(
                                {
                                    "row": idx,
                                    "field": "attachment",
                                    "name": name,
                                    "message": str(e),
                                }
                            )
                            continue
                    attachments_for_row.append(resolved)
            row_attachments.append(attachments_for_row)

        drafts = []
        for idx, row in enumerate(rows, start=1):
            if not row_can_send[idx - 1]:
                continue
            to_value = row.get(email_header)
            subject_value = row.get(subject_header)
            body_html = inject_variables(html_template, row)
            try:
                draft = create_draft(
                    creds=creds,
                    to=_normalize_email_list(to_value) or str(to_value).strip(),
                    subject=str(subject_value).strip(),
                    body_html=body_html,
                    cc=cc_value,
                    bcc=bcc_value,
                    attachments=row_attachments[idx - 1],
                )
                drafts.append(draft.get("id"))
            except Exception as e:
                failed_items.append({"row": idx, "field": "draft", "message": str(e)})

        status = "ok" if not failed_items else ("partial" if drafts else "failed")
        print(
            f"[BATCH] status={status} drafts={len(drafts)} failed_items={len(failed_items)}"
        )
        return {
            "status": status,
            "draft_count": len(drafts),
            "draft_ids": drafts,
            "failed_items": failed_items,
        }
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/process/cache")
def process_cached_preview(
    request: Request,
    cache_id: str,
    sheet: str | None = None,
    font: str | None = None,
    preview_page: int = 1,
    preview_page_size: int = 20,
):
    try:
        cache_namespace = _get_upload_cache_namespace(request)
        cached = _read_upload_cache(namespace=cache_namespace, cache_id=cache_id)
        docx_content = cached["docx_content"]
        xlsx_content = cached["xlsx_content"]
        if not isinstance(docx_content, bytes) or not isinstance(xlsx_content, bytes):
            raise HTTPException(status_code=500, detail={"error": "invalid_cache_payload"})
        cache_info = {
            "cache_id": cache_id,
            "docx_name": str(cached["docx_name"]),
            "xlsx_name": str(cached["xlsx_name"]),
        }
        return _build_preview_payload(
            docx_content=docx_content,
            template_html=None,
            xlsx_content=xlsx_content,
            sheet=sheet,
            font=font,
            cache_info=cache_info,
            preview_page=preview_page,
            preview_page_size=preview_page_size,
        )
    except HTTPException:
        raise
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"error": str(e)})


class TemplateExportRequest(BaseModel):
    template_html: str
    filename: str | None = None


@app.post("/api/template/load")
async def load_template_from_docx(
    request: Request,
    docx_file: UploadFile | None = File(None),
    docx_drive_file_id: str | None = Form(None),
    font: str | None = None,
):
    try:
        drive_service = None
        if docx_drive_file_id:
            user_key = _require_session_user_key(request)
            creds = load_user_credentials(user_key)
            drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)

        docx_name, docx_content = await _read_input_source(
            kind="docx",
            upload_file=docx_file,
            drive_file_id=docx_drive_file_id,
            drive_service=drive_service,
        )
        font_family = resolve_gmail_font(font)
        template_html = resolve_template_html(
            docx_content=docx_content,
            template_html=None,
            base_font_family=font_family,
        )
        return {"docx_name": docx_name, "template_html": template_html}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/template/export")
def export_template_docx(payload: TemplateExportRequest):
    try:
        filename = (payload.filename or "template").strip() or "template"
        safe_name = re.sub(r"[^A-Za-z0-9._-]+", "_", filename).strip("._") or "template"
        if not safe_name.lower().endswith(".docx"):
            safe_name = f"{safe_name}.docx"
        content = export_html_to_docx_bytes(payload.template_html)
        return Response(
            content=content,
            media_type=DOCX_MIME,
            headers={"Content-Disposition": f'attachment; filename="{safe_name}"'},
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/process")
async def process_files(
    request: Request,
    docx_file: UploadFile | None = File(None),
    xlsx_file: UploadFile | None = File(None),
    docx_drive_file_id: str | None = Form(None),
    xlsx_drive_file_id: str | None = Form(None),
    template_html: str | None = Form(None),
    sheet: str | None = None,
    font: str | None = None,
    preview_page: int = 1,
    preview_page_size: int = 20,
):
    try:
        cache_namespace = _get_upload_cache_namespace(request)
        drive_service = None
        if docx_drive_file_id or xlsx_drive_file_id:
            user_key = _require_session_user_key(request)
            creds = load_user_credentials(user_key)
            drive_service = build("drive", "v3", credentials=creds, cache_discovery=False)

        xlsx_name, xlsx_content = await _read_input_source(
            kind="xlsx",
            upload_file=xlsx_file,
            drive_file_id=xlsx_drive_file_id,
            drive_service=drive_service,
        )
        docx_content: bytes | None = None
        cache_info: dict[str, str] | None = None
        if docx_file or docx_drive_file_id:
            docx_name, docx_content = await _read_input_source(
                kind="docx",
                upload_file=docx_file,
                drive_file_id=docx_drive_file_id,
                drive_service=drive_service,
            )
            cache_info = _write_upload_cache(
                namespace=cache_namespace,
                docx_name=docx_name,
                docx_content=docx_content,
                xlsx_name=xlsx_name,
                xlsx_content=xlsx_content,
            )
            request.session["latest_cache_id"] = cache_info["cache_id"]
        return _build_preview_payload(
            docx_content=docx_content,
            template_html=template_html,
            xlsx_content=xlsx_content,
            sheet=sheet,
            font=font,
            cache_info=cache_info,
            preview_page=preview_page,
            preview_page_size=preview_page_size,
        )
    except Exception as e:
        traceback.print_exc()
        return JSONResponse(status_code=500, content={"error": str(e)})

# 為了讓 Vercel 以外的環境也能執行
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=port)
