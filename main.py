# main.py — Outlook delegated OAuth microservice (FastAPI + MSAL)
# Run (Render): uvicorn main:app --host 0.0.0.0 --port $PORT

import base64
import os
import threading
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

import msal
import requests
from fastapi import FastAPI, HTTPException, Header, Request
from fastapi.responses import RedirectResponse, JSONResponse
from pydantic import BaseModel, Field

# -----------------------------
# Environment / configuration
# -----------------------------
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

CLIENT_ID = os.getenv("CLIENT_ID", "").strip()
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "").strip()
AUTHORITY = os.getenv("AUTHORITY", "https://login.microsoftonline.com/common").strip()
REDIRECT_URI = os.getenv("REDIRECT_URI", "").strip()
API_KEY = os.getenv("API_KEY", "").strip()  # optional shared secret

if not (CLIENT_ID and CLIENT_SECRET and REDIRECT_URI):
    raise RuntimeError("Set CLIENT_ID, CLIENT_SECRET, REDIRECT_URI env vars.")

# ---- token cache path with safe fallback ----
DEFAULT_TOKEN_PATH = "/data/ms_tokens.json"   # persistent if a Render Disk is mounted at /data
TOKEN_PATH = os.getenv("TOKEN_PATH", DEFAULT_TOKEN_PATH)
try:
    os.makedirs(os.path.dirname(TOKEN_PATH), exist_ok=True)
except Exception:
    TOKEN_PATH = "/tmp/ms_tokens.json"        # ephemeral (lost on restart)
    os.makedirs(os.path.dirname(TOKEN_PATH), exist_ok=True)

_cache_lock = threading.Lock()

# OAuth scopes for delegated access
SCOPES = ["Mail.Read", "offline_access", "openid", "profile"]

# -----------------------------
# MSAL helpers
# -----------------------------
def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(TOKEN_PATH):
        try:
            with open(TOKEN_PATH, "r") as f:
                cache.deserialize(f.read())
        except Exception:
            pass
    return cache

def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        with open(TOKEN_PATH, "w") as f:
            f.write(cache.serialize())

def _build_app(cache: Optional[msal.SerializableTokenCache] = None) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
        token_cache=cache,
    )

def _ensure_token() -> str:
    """Return a valid access token, refreshing silently via refresh token."""
    with _cache_lock:
        cache = _load_cache()
        app = _build_app(cache)
        accounts = app.get_accounts()
        result = None
        if accounts:
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if not result or "access_token" not in result:
            raise HTTPException(status_code=401, detail="Not authorized. Visit /auth/start to sign in.")
        _save_cache(cache)
        return result["access_token"]

def _graph_headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {_ensure_token()}"}

# -----------------------------
# Models
# -----------------------------
class SearchRequest(BaseModel):
    sender_email: Optional[str] = None
    subject_contains: Optional[str] = None
    days_back: int = Field(30, ge=0, le=365)
    top: int = Field(25, ge=1, le=100)
    folder: str = Field("inbox")
    has_attachments: Optional[bool] = None  # true/false filter

class AttachmentInfo(BaseModel):
    attachmentId: str
    name: str
    size: int
    contentType: str

class MessageItem(BaseModel):
    messageId: str
    subject: str
    from_: str = Field(..., alias="from")
    fromName: str
    receivedAt: str
    webLink: Optional[str] = None
    hasAttachments: bool
    attachments: List[AttachmentInfo] = []

class SearchResponse(BaseModel):
    items: List[MessageItem]
    summary: Dict[str, Any] = {}
    debug: Dict[str, Any] = {}

class DownloadRequest(BaseModel):
    message_id: str
    attachment_id: str

class DownloadResponse(BaseModel):
    filename: str
    content_type: str
    size: int
    content_base64: str

# -----------------------------
# FastAPI app & utilities
# -----------------------------
app = FastAPI(title="Outlook Delegated Microservice", version="1.2.0")

def _check_api_key(x_api_key: Optional[str]) -> None:
    if API_KEY and (x_api_key or "").strip() != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API key")

@app.get("/")
def root():
    return {
        "ok": True,
        "service": "Outlook Delegated Microservice",
        "docs": "/docs",
        "auth_start": "/auth/start",
        "health": "/healthz"
    }

@app.get("/healthz")
def healthz():
    return {"ok": True}

# -----------------------------
# Auth endpoints
# -----------------------------
@app.get("/auth/start")
def auth_start():
    cache = _load_cache()
    appc = _build_app(cache)
    auth_url = appc.get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
        response_mode="query",
    )
    return RedirectResponse(auth_url)

@app.get("/auth/callback")
def auth_callback(request: Request):
    code = request.query_params.get("code")
    if not code:
        err = request.query_params.get("error_description", "No code provided")
        return JSONResponse({"error": err}, status_code=400)
    with _cache_lock:
        cache = _load_cache()
        appc = _build_app(cache)
        result = appc.acquire_token_by_authorization_code(code, scopes=SCOPES, redirect_uri=REDIRECT_URI)
        if "access_token" not in result:
            return JSONResponse({"error": result}, status_code=400)
        _save_cache(cache)
    return JSONResponse({"ok": True, "message": "Authorized. You can close this tab."})

# -----------------------------
# Helpers
# -----------------------------
def _normalize_message(m: Dict[str, Any]) -> MessageItem:
    ea = ((m.get("from") or {}).get("emailAddress") or {})
    return MessageItem(
        messageId=m.get("id", ""),
        subject=m.get("subject", "") or "",
        **{"from": (ea.get("address", "") or "").lower()},
        fromName=ea.get("name", "") or "",
        receivedAt=m.get("receivedDateTime", "") or "",
        webLink=m.get("webLink", ""),
        hasAttachments=bool(m.get("hasAttachments", False)),
        attachments=[],
    )

# -----------------------------
# Business endpoints
# -----------------------------
@app.post("/search", response_model=SearchResponse)
def search(req: SearchRequest, x_api_key: Optional[str] = Header(None)):
    _check_api_key(x_api_key)
    headers = _graph_headers()

    since = (datetime.now(timezone.utc) - timedelta(days=req.days_back)).isoformat()
    filters = [f"receivedDateTime ge {since}"]

    if req.sender_email:
        addr = req.sender_email.replace("'", "''").lower()
        filters.append(f"from/emailAddress/address eq '{addr}'")
    if req.subject_contains:
        sub = req.subject_contains.replace("'", "''")
        filters.append(f"contains(subject,'{sub}')")
    if req.has_attachments is True:
        filters.append("hasAttachments eq true")
    if req.has_attachments is False:
        filters.append("hasAttachments eq false")

    params = {
        "$select": "id,subject,from,receivedDateTime,webLink,hasAttachments",
        "$orderby": "receivedDateTime desc",
        "$top": str(req.top),
        "$filter": " and ".join(filters),
    }

    url = f"{GRAPH_BASE}/me/mailFolders/{req.folder or 'inbox'}/messages"
    r = requests.get(url, headers=headers, params=params, timeout=30)
    if r.status_code >= 400:
        raise HTTPException(status_code=502, detail=f"Graph list error {r.status_code}: {r.text}")

    raw = r.json().get("value", []) or []
    items: List[MessageItem] = []
    for m in raw:
        item = _normalize_message(m)
        if item.hasAttachments:
            aurl = f"{GRAPH_BASE}/me/messages/{item.messageId}/attachments?$select=id,name,contentType,size"
            ar = requests.get(aurl, headers=headers, timeout=30)
            if ar.status_code < 400:
                item.attachments = [
                    AttachmentInfo(
                        attachmentId=a.get("id", ""),
                        name=a.get("name", "") or "",
                        size=int(a.get("size", 0) or 0),
                        contentType=a.get("contentType", "") or "",
                    )
                    for a in ar.json().get("value", []) if isinstance(a, dict)
                ]
        items.append(item)

    items.sort(key=lambda x: x.receivedAt or "", reverse=True)
    return SearchResponse(
        items=items,
        summary={
            "totalMessages": len(items),
            "totalAttachments": sum(len(i.attachments) for i in items)
        },
        debug={"folder": req.folder, "count": len(items)},
    )

@app.post("/download", response_model=DownloadResponse)
def download(req: DownloadRequest, x_api_key: Optional[str] = Header(None)):
    _check_api_key(x_api_key)
    headers = _graph_headers()

    meta_url = f"{GRAPH_BASE}/me/messages/{req.message_id}/attachments/{req.attachment_id}?$select=id,name,contentType,size"
    meta = requests.get(meta_url, headers=headers, timeout=30)
    if meta.status_code >= 400:
        raise HTTPException(status_code=502, detail=f"Graph meta error {meta.status_code}: {meta.text}")
    mj = meta.json()
    filename = mj.get("name") or "attachment.bin"
    content_type = mj.get("contentType") or "application/octet-stream"
    size = int(mj.get("size", 0) or 0)

    bin_url = f"{GRAPH_BASE}/me/messages/{req.message_id}/attachments/{req.attachment_id}/$value"
    br = requests.get(bin_url, headers=headers, timeout=60)
    if br.status_code >= 400:
        raise HTTPException(status_code=502, detail=f"Graph download error {br.status_code}: {br.text}")

    b64 = base64.b64encode(br.content).decode("ascii")
    return DownloadResponse(filename=filename, content_type=content_type, size=size, content_base64=b64)


# Optional local run
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", "8000"))
    uvicorn.run("main:app", host="0.0.0.0", port=port)
