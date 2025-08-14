# main.py — Outlook delegated OAuth microservice (FastAPI + MSAL)
# Start (Render): uvicorn main:app --host 0.0.0.0 --port $PORT

import base64
import os
import threading
from datetime import datetime, timedelta, timezone
from typing import Any, Dict, List, Optional

import msal
import requests
from fastapi import FastAPI, HTTPException, Header, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import RedirectResponse, JSONResponse
from pydantic import BaseModel, Field


# -----------------------------
# Environment / configuration
# -----------------------------
GRAPH_BASE = "https://graph.microsoft.com/v1.0"

CLIENT_ID = os.getenv("CLIENT_ID", "").strip()
CLIENT_SECRET = os.getenv("CLIENT_SECRET", "").strip()
TENANT_ID = os.getenv("TENANT_ID", "").strip()
# If TENANT_ID provided, prefer it. Otherwise allow explicit AUTHORITY or default to "common".
AUTHORITY = (
    f"https://login.microsoftonline.com/{TENANT_ID}"
    if TENANT_ID
    else os.getenv("AUTHORITY", "https://login.microsoftonline.com/common").strip()
)
REDIRECT_URI = os.getenv("REDIRECT_URI", "").strip()

# Optional shared secret to protect /search and /download
API_KEY = os.getenv("API_KEY", "").strip()

if not (CLIENT_ID and CLIENT_SECRET and REDIRECT_URI):
    raise RuntimeError("Set CLIENT_ID, CLIENT_SECRET, and REDIRECT_URI environment variables.")

# Token cache path with safe fallback (/data if mounted; else /tmp)
DEFAULT_TOKEN_PATH = "/data/ms_tokens.json"
TOKEN_PATH = os.getenv("TOKEN_PATH", DEFAULT_TOKEN_PATH)
try:
    os.makedirs(os.path.dirname(TOKEN_PATH), exist_ok=True)
except Exception:
    TOKEN_PATH = "/tmp/ms_tokens.json"
    os.makedirs(os.path.dirname(TOKEN_PATH), exist_ok=True)

_cache_lock = threading.Lock()

# Delegated scopes - specify full Graph API resource URLs for correct audience
# This ensures tokens have the right audience (https://graph.microsoft.com)
SCOPES = ["https://graph.microsoft.com/Mail.Read"]


# -----------------------------
# MSAL helpers
# -----------------------------
def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    print(f"📂 Loading token cache from: {TOKEN_PATH}")
    if os.path.exists(TOKEN_PATH):
        try:
            with open(TOKEN_PATH, "r") as f:
                cache_data = f.read()
                cache.deserialize(cache_data)
                print(f"✅ Token cache loaded successfully ({len(cache_data)} chars)")
        except Exception as e:
            print(f"❌ Failed to load token cache: {str(e)}")
    else:
        print(f"⚠️  Token cache file does not exist: {TOKEN_PATH}")
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    print(f"💾 Saving token cache to: {TOKEN_PATH}")
    print(f"📊 Cache has state changed: {cache.has_state_changed}")
    if cache.has_state_changed:
        try:
            # Ensure directory exists
            os.makedirs(os.path.dirname(TOKEN_PATH), exist_ok=True)
            with open(TOKEN_PATH, "w") as f:
                cache_data = cache.serialize()
                f.write(cache_data)
                print(f"✅ Token cache saved successfully ({len(cache_data)} chars)")
        except Exception as e:
            print(f"❌ Failed to save token cache: {str(e)}")
            # Try fallback path
            fallback_path = "/tmp/ms_tokens.json"
            print(f"🔄 Trying fallback path: {fallback_path}")
            try:
                with open(fallback_path, "w") as f:
                    f.write(cache.serialize())
                print(f"✅ Token cache saved to fallback path")
            except Exception as e2:
                print(f"❌ Fallback save also failed: {str(e2)}")
    else:
        print(f"ℹ️  No cache changes to save")


def _build_app(cache: Optional[msal.SerializableTokenCache] = None) -> msal.ConfidentialClientApplication:
    return msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        client_credential=CLIENT_SECRET,
        authority=AUTHORITY,
        token_cache=cache,
    )


def _ensure_token() -> str:
    """Return a valid access token, refreshing silently (via refresh token) if possible."""
    with _cache_lock:
        cache = _load_cache()
        app = _build_app(cache)
        accounts = app.get_accounts()
        
        print(f"🔍 Found {len(accounts)} accounts in cache")
        
        result = None
        if accounts:
            print(f"👤 Using account: {accounts[0].get('username', 'unknown')}")
            result = app.acquire_token_silent(SCOPES, account=accounts[0])
            
            if result and "access_token" in result:
                print("✅ Token acquired silently (refreshed)")
            else:
                print(f"❌ Silent token acquisition failed: {result}")
        
        if not result or "access_token" not in result:
            error_msg = "Token expired or invalid. Please re-authorize at /auth/start"
            if result and "error" in result:
                error_msg += f" (MSAL Error: {result.get('error_description', result.get('error'))})"
            
            print(f"🚫 {error_msg}")
            raise HTTPException(status_code=401, detail=error_msg)
        
        _save_cache(cache)
        print("🎯 Returning valid access token")
        return result["access_token"]


def _graph_headers() -> Dict[str, str]:
    token = _ensure_token()
    print(f"🔑 Using access token: {token[:20]}...{token[-10:] if len(token) > 30 else ''}")
    
    # Decode JWT token for debugging (without verification)
    try:
        import base64
        import json
        # Split JWT token and decode payload
        parts = token.split('.')
        if len(parts) >= 2:
            # Add padding if needed
            payload = parts[1]
            payload += '=' * (4 - len(payload) % 4)
            decoded = base64.b64decode(payload)
            token_data = json.loads(decoded)
            print(f"🎫 Token audience: {token_data.get('aud', 'NOT_SET')}")
            print(f"🎫 Token scopes: {token_data.get('scp', 'NOT_SET')}")
            print(f"🎫 Token issuer: {token_data.get('iss', 'NOT_SET')}")
            print(f"🎫 Token expires: {token_data.get('exp', 'NOT_SET')}")
    except Exception as e:
        print(f"⚠️  Could not decode token for debugging: {str(e)}")
    
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/json",
    }


# -----------------------------
# Pydantic models
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
# FastAPI app
# -----------------------------
app = FastAPI(title="Outlook Delegated Microservice", version="1.3.0")

# CORS (adjust origins if you want to restrict)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)


def _check_api_key(x_api_key: Optional[str]) -> None:
    if API_KEY and (x_api_key or "").strip() != API_KEY:
        raise HTTPException(status_code=401, detail="Invalid API key")


# FIX 1: Handle both GET and HEAD requests for health checks
@app.get("/")
@app.head("/")
def root():
    return {
        "ok": True,
        "service": "Outlook Delegated Microservice",
        "docs": "/docs",
        "auth_start": "/auth/start",
        "health": "/healthz",
        "authority": AUTHORITY,
        "redirect_uri": REDIRECT_URI,
        "status": "running"
    }


# FIX 2: Also handle HEAD requests for health endpoint
@app.get("/healthz")
@app.head("/healthz")
def healthz():
    return {"ok": True, "status": "healthy"}


# -----------------------------
# OAuth endpoints
# -----------------------------
@app.get("/auth/start")
def auth_start():
    try:
        # Debug logging
        print(f"🔐 Starting OAuth flow...")
        print(f"📍 Authority: {AUTHORITY}")
        print(f"🆔 Client ID: {CLIENT_ID[:8]}...")
        print(f"🔄 Redirect URI: {REDIRECT_URI}")
        print(f"📝 Scopes: {SCOPES}")
        
        cache = _load_cache()
        appc = _build_app(cache)
        
        auth_url = appc.get_authorization_request_url(
            scopes=SCOPES,
            redirect_uri=REDIRECT_URI,
            prompt="select_account",
            response_mode="query",
        )
        
        print(f"✅ Generated auth URL: {auth_url[:100]}...")
        return RedirectResponse(auth_url)
        
    except Exception as e:
        print(f"❌ Error in auth_start: {str(e)}")
        print(f"📊 Error type: {type(e).__name__}")
        import traceback
        traceback.print_exc()
        
        return JSONResponse({
            "error": "OAuth initialization failed",
            "details": str(e),
            "authority": AUTHORITY,
            "client_id_prefix": CLIENT_ID[:8] if CLIENT_ID else "NOT_SET",
            "redirect_uri": REDIRECT_URI
        }, status_code=500)


# alias
@app.get("/login")
def login_alias():
    return auth_start()


@app.get("/debug/config")
def debug_config():
    """Debug endpoint to check configuration (remove in production)"""
    return {
        "authority": AUTHORITY,
        "client_id_set": bool(CLIENT_ID),
        "client_id_prefix": CLIENT_ID[:8] if CLIENT_ID else "NOT_SET",
        "client_secret_set": bool(CLIENT_SECRET),
        "redirect_uri": REDIRECT_URI,
        "redirect_uri_set": bool(REDIRECT_URI),
        "tenant_id": TENANT_ID or "NOT_SET",
        "api_key_set": bool(API_KEY),
        "token_path": TOKEN_PATH,
        "scopes": SCOPES
    }


@app.get("/debug/graph-test")
def debug_graph_test(x_api_key: Optional[str] = Header(None)):
    """Test basic Graph API connectivity with /me endpoint"""
    _check_api_key(x_api_key)
    
    try:
        headers = _graph_headers()
        
        # Test with simplest Graph API endpoint
        test_url = f"{GRAPH_BASE}/me"
        print(f"🧪 Testing Graph API connectivity: {test_url}")
        
        r = requests.get(test_url, headers=headers, timeout=30)
        print(f"📊 Graph /me Response: {r.status_code}")
        
        if r.status_code >= 400:
            print(f"❌ Graph /me error: {r.text}")
            print(f"🔍 Response headers: {dict(r.headers)}")
            
            try:
                error_json = r.json()
                print(f"📋 Graph /me error details: {error_json}")
                return {"error": "Graph API test failed", "status": r.status_code, "details": error_json}
            except:
                return {"error": "Graph API test failed", "status": r.status_code, "raw_response": r.text}
        
        user_info = r.json()
        print(f"✅ Graph API test successful: {user_info.get('userPrincipalName', 'unknown')}")
        
        return {
            "success": True,
            "user_principal_name": user_info.get("userPrincipalName"),
            "display_name": user_info.get("displayName"),
            "id": user_info.get("id"),
            "status": r.status_code
        }
        
    except Exception as e:
        print(f"💥 Graph API test error: {str(e)}")
        return {"error": f"Graph API test failed: {str(e)}"}


@app.get("/auth/callback")
def auth_callback(request: Request):
    try:
        print(f"🔄 OAuth callback received")
        code = request.query_params.get("code")
        error = request.query_params.get("error")
        error_desc = request.query_params.get("error_description")
        
        if error:
            print(f"❌ OAuth error: {error} - {error_desc}")
            return JSONResponse({"error": f"{error}: {error_desc}"}, status_code=400)
            
        if not code:
            print(f"❌ No authorization code received")
            return JSONResponse({"error": "No authorization code provided"}, status_code=400)

        print(f"✅ Authorization code received: {code[:20]}...")

        with _cache_lock:
            cache = _load_cache()
            appc = _build_app(cache)
            
            print(f"🔄 Acquiring token by authorization code...")
            result = appc.acquire_token_by_authorization_code(
                code, 
                scopes=SCOPES, 
                redirect_uri=REDIRECT_URI
            )

            print(f"📊 Token acquisition result keys: {list(result.keys())}")
            
            if "access_token" not in result:
                print(f"❌ Token acquisition failed: {result}")
                return JSONResponse({"error": "Token acquisition failed", "details": result}, status_code=400)

            # Log successful token info (without exposing the actual token)
            token_info = {
                "has_access_token": "access_token" in result,
                "has_refresh_token": "refresh_token" in result,
                "expires_in": result.get("expires_in"),
                "scope": result.get("scope"),
                "token_type": result.get("token_type")
            }
            print(f"✅ Token acquired successfully: {token_info}")
            
            _save_cache(cache)
            print(f"💾 Token cache saved")

        return JSONResponse({
            "ok": True, 
            "message": "Authorized successfully. You can close this tab.",
            "token_info": token_info
        })
        
    except Exception as e:
        print(f"💥 OAuth callback error: {str(e)}")
        import traceback
        traceback.print_exc()
        return JSONResponse({"error": f"OAuth callback failed: {str(e)}"}, status_code=500)


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
# Search & Download
# -----------------------------
@app.post("/search", response_model=SearchResponse)
def search(req: SearchRequest, x_api_key: Optional[str] = Header(None)):
    _check_api_key(x_api_key)
    
    try:
        headers = _graph_headers()
        print(f"🔍 Searching emails: days_back={req.days_back}, top={req.top}")
        
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
        print(f"🌐 Graph API URL: {url}")
        
        r = requests.get(url, headers=headers, params=params, timeout=30)
        print(f"📊 Graph API Response: {r.status_code}")
        
        if r.status_code >= 400:
            error_detail = f"Graph list error {r.status_code}: {r.text}"
            print(f"❌ {error_detail}")
            
            # Log response headers for debugging
            print(f"🔍 Response headers: {dict(r.headers)}")
            
            # Try to parse JSON error for more details
            try:
                error_json = r.json()
                print(f"📋 Graph API error details: {error_json}")
            except:
                print(f"📋 Raw error response: {r.text}")
                
            raise HTTPException(status_code=502, detail=error_detail)

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
        print(f"✅ Found {len(items)} messages")
        
        return SearchResponse(
            items=items,
            summary={
                "totalMessages": len(items),
                "totalAttachments": sum(len(i.attachments) for i in items),
            },
            debug={"folder": req.folder, "count": len(items)},
        )
    
    except HTTPException:
        raise
    except Exception as e:
        print(f"💥 Unexpected error in search: {str(e)}")
        import traceback
        traceback.print_exc()
        raise HTTPException(status_code=500, detail=f"Search failed: {str(e)}")


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


# FIX 3: Add startup event to log configuration
@app.on_event("startup")
async def startup_event():
    print(f"🚀 Service starting up...")
    print(f"📍 Authority: {AUTHORITY}")
    print(f"🔄 Redirect URI: {REDIRECT_URI}")
    print(f"🗂️ Token cache path: {TOKEN_PATH}")
    if API_KEY:
        print("🔐 API key protection enabled")
    else:
        print("⚠️  No API key protection")


# -----------------------------
# Local run - FIX 4: Proper port handling
# -----------------------------
if __name__ == "__main__":
    import uvicorn
    # Use PORT environment variable (what Render sets), fallback to 8000 for local dev
    port = int(os.environ.get("PORT", 8000))
    print(f"Starting server on port {port}")
    uvicorn.run("main:app", host="0.0.0.0", port=port)