import base64
import os
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

import msal
import requests
from fastapi import FastAPI, Header, HTTPException
from pydantic import BaseModel, Field

# ------------------ Config ------------------

GRAPH_BASE = os.getenv("GRAPH_BASE", "https://graph.microsoft.com/v1.0")

TENANT_ID = os.getenv("TENANT_ID", "").strip()
CLIENT_ID = os.getenv("CLIENT_ID", "").strip()
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "").strip()

DEFAULT_USER_EMAIL = os.getenv("DEFAULT_USER_EMAIL", "").strip()  # fallback mailbox
API_KEY = os.getenv("API_KEY", "").strip()  # optional: simple header auth

if not (TENANT_ID and CLIENT_ID and CLIENT_SECRET):
    raise RuntimeError("TENANT_ID, CLIENT_ID, CLIENT_SECRET must be set as env vars")

# ------------------ Auth ------------------

def acquire_token() -> str:
    """
    App-only (client credentials) token. Requires:
    - App registration in same Entra tenant as the target mailbox
    - Microsoft Graph Application permission: Mail.Read (admin consent)
    """
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
    )
    res = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
    if "access_token" not in res:
        raise HTTPException(status_code=500, detail=f"Token error: {res}")
    return res["access_token"]

# ------------------ Models ------------------

class SearchRequest(BaseModel):
    user_email: Optional[str] = Field(None, description="Target mailbox; defaults to DEFAULT_USER_EMAIL")
    sender_email: Optional[str] = Field(None, description="Filter: exact from address")
    subject_contains: Optional[str] = Field(None, description="Filter: subject contains text")
    days_back: int = Field(7, ge=0, le=365)
    top: int = Field(25, ge=1, le=100)
    folder: str = Field("inbox", description="Well-known name or folder id")
    has_attachments: Optional[bool] = Field(None, description="If true, only messages with attachments")

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
    user_email: Optional[str]
    message_id: str
    attachment_id: str

class DownloadResponse(BaseModel):
    filename: str
    content_type: str
    size: int
    content_base64: str

# ------------------ FastAPI ------------------

app = FastAPI(title="Outlook Attachment Microservice", version="1.0.0")

def check_api_key(x_api_key: Optional[str]):
    if API_KEY and (x_api_key or "").strip() != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API key")

def graph_headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {acquire_token()}"}

def normalize_message(m: Dict[str, Any]) -> MessageItem:
    ea = ((m.get("from") or {}).get("emailAddress") or {})
    return MessageItem(
        messageId=m.get("id", ""),
        subject=m.get("subject", "") or "",
        **{"from": ea.get("address", "") or ""},
        fromName=ea.get("name", "") or "",
        receivedAt=m.get("receivedDateTime", "") or "",
        webLink=m.get("webLink", ""),
        hasAttachments=bool(m.get("hasAttachments", False)),
        attachments=[],
    )

def list_messages(
    user_email: str,
    folder: str,
    top: int,
    days_back: int,
    sender_email: Optional[str],
    subject_contains: Optional[str],
    has_attachments: Optional[bool],
) -> List[Dict[str, Any]]:
    since = (datetime.now(timezone.utc) - timedelta(days=days_back)).isoformat()
    filters = [f"receivedDateTime ge {since}"]
    if sender_email:
        addr = sender_email.replace("'", "''").lower()
        filters.append(f"from/emailAddress/address eq '{addr}'")
    if subject_contains:
        val = subject_contains.replace("'", "''")
        filters.append(f"contains(subject,'{val}')")
    if has_attachments is True:
        filters.append("hasAttachments eq true")
    if has_attachments is False:
        filters.append("hasAttachments eq false")

    params = {
        "$select": "id,subject,from,receivedDateTime,webLink,hasAttachments",
        "$orderby": "receivedDateTime desc",
        "$top": str(top),
        "$filter": " and ".join(filters),
    }
    url = f"{GRAPH_BASE}/users/{user_email}/mailFolders/{folder}/messages"
    r = requests.get(url, headers=graph_headers(), params=params, timeout=30)
    if r.status_code >= 400:
        raise HTTPException(status_code=502, detail=f"Graph list error {r.status_code}: {r.text}")
    return r.json().get("value", []) or []

def list_attachments(user_email: str, message_id: str) -> List[Dict[str, Any]]:
    url = f"{GRAPH_BASE}/users/{user_email}/messages/{message_id}/attachments"
    params = {"$select": "id,name,contentType,size"}
    r = requests.get(url, headers=graph_headers(), params=params, timeout=30)
    if r.status_code >= 400:
        # don't fail the whole request; return empty
        return []
    return r.json().get("value", []) or []

@app.get("/healthz")
def healthz():
    return {"ok": True}

@app.post("/search", response_model=SearchResponse)
def search(req: SearchRequest, x_api_key: Optional[str] = Header(None)):
    check_api_key(x_api_key)
    user_email = (req.user_email or DEFAULT_USER_EMAIL).strip()
    if not user_email:
        raise HTTPException(status_code=400, detail="user_email missing (and DEFAULT_USER_EMAIL not set)")

    raw = list_messages(
        user_email=user_email,
        folder=req.folder or "inbox",
        top=req.top,
        days_back=req.days_back,
        sender_email=req.sender_email,
        subject_contains=req.subject_contains,
        has_attachments=req.has_attachments,
    )

    items: List[MessageItem] = []
    for m in raw:
        item = normalize_message(m)
        if item.hasAttachments:
            atts = list_attachments(user_email, item.messageId)
            item.attachments = [
                AttachmentInfo(
                    attachmentId=a.get("id", ""),
                    name=a.get("name", "") or "",
                    size=int(a.get("size", 0) or 0),
                    contentType=a.get("contentType", "") or "",
                )
                for a in atts if isinstance(a, dict)
            ]
        items.append(item)

    items.sort(key=lambda x: x.receivedAt or "", reverse=True)
    return SearchResponse(
        items=items,
        summary={"totalMessages": len(items), "totalAttachments": sum(len(i.attachments) for i in items)},
        debug={"user": user_email, "folder": req.folder, "count": len(items)},
    )

@app.post("/download", response_model=DownloadResponse)
def download(req: DownloadRequest, x_api_key: Optional[str] = Header(None)):
    check_api_key(x_api_key)
    user_email = (req.user_email or DEFAULT_USER_EMAIL).strip()
    if not (user_email and req.message_id and req.attachment_id):
        raise HTTPException(status_code=400, detail="user_email, message_id, attachment_id are required")

    # 1) Get metadata (filename, contentType, size)
    meta_url = f"{GRAPH_BASE}/users/{user_email}/messages/{req.message_id}/attachments/{req.attachment_id}"
    meta = requests.get(meta_url, headers=graph_headers(), params={"$select": "id,name,contentType,size"}, timeout=30)
    if meta.status_code >= 400:
        raise HTTPException(status_code=502, detail=f"Graph meta error {meta.status_code}: {meta.text}")
    meta_json = meta.json()
    filename = meta_json.get("name") or "attachment.bin"
    content_type = meta_json.get("contentType") or "application/octet-stream"
    size = int(meta_json.get("size", 0) or 0)

    # 2) Download bytes
    bin_url = f"{GRAPH_BASE}/users/{user_email}/messages/{req.message_id}/attachments/{req.attachment_id}/$value"
    bin_res = requests.get(bin_url, headers=graph_headers(), timeout=60)
    if bin_res.status_code >= 400:
        raise HTTPException(status_code=502, detail=f"Graph download error {bin_res.status_code}: {bin_res.text}")

    b64 = base64.b64encode(bin_res.content).decode("ascii")
    return DownloadResponse(filename=filename, content_type=content_type, size=size, content_base64=b64)
