from fastapi import FastAPI, UploadFile, File, HTTPException, Request
from fastapi.responses import JSONResponse
from starlette.middleware.sessions import SessionMiddleware
from pydantic import BaseModel
import pandas as pd
import io
from api.core.processor import convert_docx_to_html, inject_variables
from api.core.gmail_svc import (
    create_draft,
    exchange_code_for_token,
    get_auth_url,
    load_user_credentials,
)
import os
port = int(os.environ.get("PORT", 8000))
SESSION_SECRET = os.environ.get("SESSION_SECRET")
if not SESSION_SECRET:
    raise RuntimeError("Missing required environment variable: SESSION_SECRET")

app = FastAPI()
app.add_middleware(
    SessionMiddleware,
    secret_key=SESSION_SECRET,
    same_site="lax",
    https_only=os.environ.get("SESSION_HTTPS_ONLY", "false").lower() == "true",
)

@app.get("/api/health")
def health_check():
    return {"status": "ok", "message": "Gmail Replicator API is running"}

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

@app.get("/api/auth/google/callback")
def google_auth_callback(request: Request, code: str, state: str):
    try:
        expected_state = request.session.get("oauth_state")
        if not expected_state or state != expected_state:
            raise HTTPException(status_code=400, detail="Invalid OAuth state")
        user_key = request.session.get("oauth_user_key")
        if not user_key:
            raise HTTPException(status_code=401, detail="Missing session user_key")
        creds = exchange_code_for_token(code=code, state=state, user_key=user_key)
        request.session.pop("oauth_state", None)
        request.session.pop("oauth_user_key", None)
        return {
            "status": "ok",
            "scopes": creds.scopes,
        }
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

@app.post("/api/process")
async def process_files(
    docx_file: UploadFile = File(...),
    xlsx_file: UploadFile = File(...)
):
    try:
        # 1. 讀取 Word 模板並轉為 HTML
        docx_content = await docx_file.read()
        html_template = convert_docx_to_html(docx_content)

        # 2. 使用 Pandas 讀取 Excel 數據
        xlsx_content = await xlsx_file.read()
        df = pd.read_excel(io.BytesIO(xlsx_content))
        
        # 將 DataFrame 轉為字典列表，方便處理自定義欄位
        rows = df.to_dict(orient="records")

        # 3. 產生第一筆預覽 (作為測試)
        preview = ""
        if rows:
            preview = inject_variables(html_template, rows[0])

        return {
            "total_records": len(rows),
            "headers": list(df.columns),
            "preview_first_row": preview
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"error": str(e)})

# 為了讓 Vercel 以外的環境也能執行
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=port)
