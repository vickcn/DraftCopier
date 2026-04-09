import base64
import json
import os
import secrets
import mimetypes
from email import encoders
from email.header import Header
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Optional, Protocol, Sequence

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build

DEFAULT_SCOPES = ["https://www.googleapis.com/auth/gmail.compose"]

# Detact Vercel environment for writable /tmp directory
IS_VERCEL = os.getenv("VERCEL") == "1"
if IS_VERCEL:
    TOKEN_DIR = Path("/tmp/gmail_tokens")
else:
    TOKEN_DIR = Path(os.getenv("GMAIL_TOKEN_DIR", ".secrets/gmail_tokens"))


class TokenStore(Protocol):
    """Backend-only token storage interface."""

    def save(self, user_key: str, creds: Credentials) -> None: ...

    def load(self, user_key: str) -> Optional[Credentials]: ...

    def delete(self, user_key: str) -> None: ...


class FileTokenStore:
    """
    Backend-only local token store.

    Production should prefer encrypted DB / KMS / secret manager.
    This implementation still keeps tokens on the server side only and never
    returns them to the frontend.
    """

    def __init__(self, root: Path | str = TOKEN_DIR):
        self.root = Path(root)
        try:
            self.root.mkdir(parents=True, exist_ok=True)
            # Only attempt chmod if not in serverless/readonly environments where it might fail
            if not IS_VERCEL:
                os.chmod(self.root, 0o700)
        except Exception as e:
            print(f"[!] Warning: Failed to prepare token directory {self.root}: {e}")

    def _path(self, user_key: str) -> Path:
        safe_name = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in user_key)
        return self.root / f"{safe_name}.json"

    def save(self, user_key: str, creds: Credentials) -> None:
        path = self._path(user_key)
        path.write_text(creds.to_json(), encoding="utf-8")
        try:
            os.chmod(path, 0o600)
        except OSError:
            pass

    def load(self, user_key: str) -> Optional[Credentials]:
        path = self._path(user_key)
        if not path.exists():
            return None
        return Credentials.from_authorized_user_file(str(path), scopes=DEFAULT_SCOPES)

    def delete(self, user_key: str) -> None:
        path = self._path(user_key)
        if path.exists():
            path.unlink()


# Replace this with your encrypted DB implementation in production.
token_store: TokenStore = FileTokenStore()


def _require_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def _client_config() -> dict:
    return {
        "web": {
            "client_id": _require_env("GOOGLE_CLIENT_ID"),
            "client_secret": _require_env("GOOGLE_CLIENT_SECRET"),
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
            "redirect_uris": [_require_env("GOOGLE_REDIRECT_URI")],
        }
    }


def _scopes(scopes: Optional[Sequence[str]] = None) -> list[str]:
    return list(scopes or DEFAULT_SCOPES)


def get_auth_url(
        state: Optional[str] = None,
        scopes: Optional[Sequence[str]] = None,
    ) -> tuple[str, str, str | None]:
    """
    Generate the Google OAuth2 authorization URL.

    Returns:
        (auth_url, state)

    Notes:
    - redirect_uri is fixed on the backend from environment variables.
    - only the auth URL and state are returned to the frontend.
    - tokens are exchanged and stored on the server only.
    """
    redirect_uri = _require_env("GOOGLE_REDIRECT_URI")
    flow = Flow.from_client_config(
        _client_config(),
        scopes=_scopes(scopes),
        state=state or secrets.token_urlsafe(24),
    )
    flow.redirect_uri = redirect_uri

    auth_url, resolved_state = flow.authorization_url(
        access_type="offline",
        include_granted_scopes="true",
        prompt="consent",
    )
    code_verifier = getattr(flow, 'code_verifier', None)
    return auth_url, resolved_state, code_verifier


def exchange_code_for_token(
        code: str,
        *,
        user_key: str,
        state: Optional[str] = None,
        scopes: Optional[Sequence[str]] = None,
        code_verifier: Optional[str] = None,
    ) -> Credentials:
    """
    Exchange an OAuth2 code for credentials and save them server-side.
    """
    redirect_uri = _require_env("GOOGLE_REDIRECT_URI")
    flow = Flow.from_client_config(
        _client_config(),
        scopes=_scopes(scopes),
        state=state,
    )
    flow.redirect_uri = redirect_uri
    if code_verifier:
        try:
            flow.fetch_token(code=code, code_verifier=code_verifier)
        except TypeError:
            flow.fetch_token(code=code)
    else:
        flow.fetch_token(code=code)

    creds = flow.credentials
    token_store.save(user_key, creds)
    return creds


def load_user_credentials(user_key: str) -> Credentials:
    """
    Load backend-stored credentials for a user and refresh if needed.
    """
    creds = token_store.load(user_key)
    if creds is None:
        raise RuntimeError(f"No stored Gmail credentials for user_key={user_key!r}")

    if creds.expired and creds.refresh_token:
        creds.refresh(Request())
        token_store.save(user_key, creds)

    return creds


def create_draft(
        creds: Credentials,
        to: str,
        subject: str,
        body_html: str,
        attachments: Optional[Sequence[dict[str, str | bytes]]] = None,
    ) -> dict:
    """
    Create a Gmail draft using HTML body content.

    Args:
        creds: OAuth2 credentials already obtained and stored server-side.
        to: recipient email address.
        subject: email subject.
        body_html: HTML email body.

    Returns:
        Gmail API response payload for the created draft.
    """
    if attachments:
        message = MIMEMultipart()
        message["to"] = to
        message["subject"] = subject
        message.attach(MIMEText(body_html, "html", "utf-8"))

        for attachment in attachments:
            filename = str(attachment.get("filename", "attachment"))
            content = attachment.get("content", b"")
            if isinstance(content, str):
                content = content.encode("utf-8")
            mime_type = str(attachment.get("mime_type", "")) or mimetypes.guess_type(filename)[0]
            if not mime_type:
                mime_type = "application/octet-stream"
            maintype, subtype = mime_type.split("/", 1)
            part = MIMEBase(maintype, subtype)
            part.set_payload(content)
            encoders.encode_base64(part)
            part.add_header(
                "Content-Disposition",
                "attachment",
                filename=str(Header(filename, "utf-8")),
            )
            message.attach(part)
    else:
        message = MIMEText(body_html, "html", "utf-8")
        message["to"] = to
        message["subject"] = subject

    encoded_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    body = {"message": {"raw": encoded_message}}

    service = build("gmail", "v1", credentials=creds, cache_discovery=False)
    return service.users().drafts().create(userId="me", body=body).execute()


__all__ = [
    "DEFAULT_SCOPES",
    "FileTokenStore",
    "TokenStore",
    "create_draft",
    "exchange_code_for_token",
    "get_auth_url",
    "load_user_credentials",
    "revoke_user_credentials",
    "token_store",
]


def revoke_user_credentials(user_key: str) -> None:
    """
    Delete stored credentials for a user. Frontend should also提醒使用者到
    Google 帳號的「第三方存取」頁面撤銷存取權。
    """
    token_store.delete(user_key)
