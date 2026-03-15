import os
import json
from typing import List, Optional
from fastapi import FastAPI, Request, Depends, HTTPException, status, BackgroundTasks
from fastapi.responses import RedirectResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware
from dotenv import load_dotenv

from auth import get_auth_url, get_token_from_code, get_token_silent, authenticate_oauth2
from imap_service import OutlookService
from filters import SenderNameFilter
from main import clean_inbox

# Load environment variables
load_dotenv()

app = FastAPI(title="Outlook Cleaner Web")
app.add_middleware(SessionMiddleware, secret_key=os.getenv("SESSION_SECRET", "default-secret-change-me"))

# Configuration
CLIENT_ID = os.getenv("AZURE_CLIENT_ID")
CLIENT_SECRET = os.getenv("AZURE_CLIENT_SECRET")
TENANT_ID = os.getenv("TENANT_ID", "consumers")
REDIRECT_URI = os.getenv("REDIRECT_URI")
API_KEY = os.getenv("X_API_KEY")
USER_EMAIL = os.getenv("USER_EMAIL")

# Token cache persistence
CACHE_FILE = os.getenv("CACHE_FILE", "token_cache.json")

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            return f.read()
    return None

def save_cache(cache_serialized):
    with open(CACHE_FILE, "w") as f:
        f.write(cache_serialized)

# Templates
templates = Jinja2Templates(directory="templates")

def get_current_user_token():
    """Dependency to get a valid token from cache/refresh"""
    cache_json = load_cache()
    if not cache_json:
        return None
    
    result, new_cache = get_token_silent(CLIENT_ID, CLIENT_SECRET, cache_json, TENANT_ID)
    if result:
        save_cache(new_cache)
        return result.get("access_token")
    return None

@app.get("/", response_class=HTMLResponse)
async def index(request: Request, token: str = Depends(get_current_user_token)):
    if not token:
        return templates.TemplateResponse("login.html", {"request": request})
    
    # Load configuration from config.json or SENDER_NAMES env var
    senders = []
    env_senders = os.getenv("SENDER_NAMES")
    if env_senders:
        senders = [s.strip() for s in env_senders.split(",")]
    else:
        try:
            from config import load_config
            config = load_config()
            senders = config.get("cleaning", {}).get("sender_names_to_search", [])
        except:
            pass
        
    return templates.TemplateResponse("dashboard.html", {
        "request": request, 
        "user_email": USER_EMAIL,
        "senders": senders,
        "api_key": API_KEY # We pass it here for convenience in the demo
    })

@app.get("/login")
async def login():
    auth_url = get_auth_url(CLIENT_ID, CLIENT_SECRET, REDIRECT_URI, TENANT_ID)
    return RedirectResponse(auth_url)

@app.get("/callback")
async def callback(code: str):
    # Use a real cache object during the exchange
    from auth import build_msal_app
    import msal
    app_cache = msal.SerializableTokenCache()
    app_with_cache = msal.ConfidentialClientApplication(
        CLIENT_ID, 
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        client_credential=CLIENT_SECRET,
        token_cache=app_cache
    )
    result = app_with_cache.acquire_token_by_authorization_code(
        code, 
        scopes=["https://outlook.office.com/IMAP.AccessAsUser.All"], 
        redirect_uri=REDIRECT_URI
    )
    
    if "access_token" in result:
        save_cache(app_cache.serialize())
        return RedirectResponse("/")
    else:
        raise HTTPException(status_code=400, detail=f"Login failed: {result.get('error_description')}")

@app.post("/trigger")
async def ui_trigger(background_tasks: BackgroundTasks, token: str = Depends(get_current_user_token)):
    """Trigger for the Web UI (authenticated by session)"""
    if not token:
        raise HTTPException(status_code=401, detail="Not authenticated")
    
    background_tasks.add_task(run_cleanup_task, token)
    return {"status": "accepted"}

@app.post("/api/v1/trigger-clean")
async def api_trigger(request: Request, background_tasks: BackgroundTasks):
    """Trigger for External API (authenticated by API Key)"""
    # Verify API Key
    header_key = request.headers.get("X-API-KEY")
    if header_key != API_KEY:
        raise HTTPException(status_code=403, detail="Invalid API Key")
    
    # Check for valid token
    token = get_current_user_token()
    if not token:
        raise HTTPException(status_code=401, detail="Service not authenticated. Please login via web UI.")
    
    # Run cleanup in background
    background_tasks.add_task(run_cleanup_task, token)
    
    return {"status": "accepted", "message": "Cleanup process started in background"}

def run_cleanup_task(token):
    from config import load_config
    config = load_config()
    
    # Get senders from env var or config file
    env_senders = os.getenv("SENDER_NAMES")
    if env_senders:
        senders = [s.strip() for s in env_senders.split(",")]
    else:
        senders = config.get("cleaning", {}).get("sender_names_to_search", [])
        
    email_filter = SenderNameFilter(senders)
    
    clean_inbox(
        USER_EMAIL,
        email_filter,
        client_id=CLIENT_ID,
        tenant_id=TENANT_ID,
        access_token=token,
        move_to_deleted=config.get("cleaning", {}).get("move_to_deleted", True)
    )

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
