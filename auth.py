"""OAuth2 authentication module"""
import imaplib
import base64

# For OAuth2, install: pip install msal
try:
    import msal
    OAUTH2_AVAILABLE = True
except ImportError:
    OAUTH2_AVAILABLE = False
    print("[WARN] msal is not installed. To use OAuth2, run: pip install msal")


def _read_imap_response(imap_conn):
    """Helper function to read and decode IMAP response"""
    response = imap_conn.readline()
    if isinstance(response, bytes):
        return response.decode('utf-8', errors='ignore').strip()
    return response.strip() if response else ''


def _read_final_auth_response(imap_conn, tag_str, max_attempts=10):
    """Read the final authentication response from server"""
    response_lines = []
    attempts = 0
    
    while attempts < max_attempts:
        response = _read_imap_response(imap_conn)
        if not response:
            break
        response_lines.append(response)
        if response.startswith(tag_str):
            break
        attempts += 1
    
    return ' '.join(response_lines)


def _validate_auth_response(full_response, tag_str):
    """Validate if authentication response indicates success"""
    if tag_str in full_response and 'OK' in full_response.upper():
        return True
    error_msg = f"XOAUTH2 authentication failed. Server response: {full_response}"
    raise imaplib.IMAP4.error(error_msg)


def _authenticate_oauth2_manual(imap_conn, auth_bytes):
    """
    Manual implementation of XOAUTH2 protocol (fallback)
    """
    tag = imap_conn._new_tag()
    command = f'{tag} AUTHENTICATE XOAUTH2'
    
    try:
        # Send AUTHENTICATE XOAUTH2 command
        imap_conn.send(f'{command}\r\n'.encode('utf-8'))
        
        # Read initial server response
        response = _read_imap_response(imap_conn)
        
        # Server should respond with "+" or "+ " (with space)
        if not (response.startswith('+') or response == '+'):
            error_msg = f"Unexpected server response: {response}. Expected '+' to continue."
            raise imaplib.IMAP4.error(error_msg)
        
        # Send base64 encoded string
        imap_conn.send(f'{auth_bytes}\r\n'.encode('utf-8'))
        
        # Read and validate final server response
        tag_str = str(tag)
        full_response = _read_final_auth_response(imap_conn, tag_str)
        
        if _validate_auth_response(full_response, tag_str):
            # Manually update connection state
            # imaplib uses 'AUTH' as state after successful authentication
            imap_conn.state = 'AUTH'
            return True
            
    except Exception as e:
        if isinstance(e, imaplib.IMAP4.error):
            raise
        raise imaplib.IMAP4.error(f"Error during XOAUTH2 authentication: {str(e)}")


def authenticate_oauth2(imap_conn, email_address, access_token):
    """
    Authenticates using OAuth2 with XOAUTH2 method
    """
    # Build XOAUTH2 authentication string
    auth_string = f"user={email_address}\x01auth=Bearer {access_token}\x01\x01"
    auth_bytes = base64.b64encode(auth_string.encode('utf-8')).decode('utf-8')
    
    # Try to use imaplib's authenticate method if available
    try:
        # The authenticate method expects a function that receives the server response
        # and returns the authentication string
        def auth_mechanism(response):
            # response can be bytes or string, but we always return the encoded string
            return auth_bytes
        
        # Use imaplib's authenticate method (available in Python 3.9+)
        typ, data = imap_conn.authenticate('XOAUTH2', auth_mechanism)
        if typ == 'OK':
            return True
        else:
            error_msg = f"XOAUTH2 authentication failed. Response: {data}"
            raise imaplib.IMAP4.error(error_msg)
    except (AttributeError, TypeError) as e:
        # If authenticate is not available or fails, use manual implementation
        print(f"[WARN] Using manual XOAUTH2 implementation: {e}")
        return _authenticate_oauth2_manual(imap_conn, auth_bytes)
    except Exception as e:
        # If there's another error, try manual implementation as fallback
        print(f"[WARN] Error with authenticate(), trying manual method: {e}")
        return _authenticate_oauth2_manual(imap_conn, auth_bytes)


def build_msal_app(client_id, client_secret=None, tenant_id="consumers"):
    """Builds the MSAL application instance"""
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    if client_secret:
        return msal.ConfidentialClientApplication(
            client_id,
            authority=authority,
            client_credential=client_secret
        )
    return msal.PublicClientApplication(
        client_id,
        authority=authority
    )


def get_auth_url(client_id, client_secret, redirect_uri, tenant_id="consumers"):
    """Generates the authorization URL for the web flow"""
    app = build_msal_app(client_id, client_secret, tenant_id)
    scopes = ["https://outlook.office.com/IMAP.AccessAsUser.All"]
    return app.get_authorization_request_url(scopes, redirect_uri=redirect_uri)


def get_token_from_code(client_id, client_secret, redirect_uri, code, tenant_id="consumers"):
    """Exchanges an authorization code for tokens"""
    app = build_msal_app(client_id, client_secret, tenant_id)
    scopes = ["https://outlook.office.com/IMAP.AccessAsUser.All"]
    result = app.acquire_token_by_authorization_code(
        code, 
        scopes=scopes, 
        redirect_uri=redirect_uri
    )
    return result


def get_token_silent(client_id, client_secret, token_cache_json, tenant_id="consumers"):
    """Attempts to get a token silently using the cache (refresh token)"""
    cache = msal.SerializableTokenCache()
    if token_cache_json:
        cache.deserialize(token_cache_json)
    
    app = msal.ConfidentialClientApplication(
        client_id,
        authority=f"https://login.microsoftonline.com/{tenant_id}",
        client_credential=client_secret,
        token_cache=cache
    )
    
    scopes = ["https://outlook.office.com/IMAP.AccessAsUser.All"]
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(scopes, account=accounts[0])
        return result, cache.serialize()
    return None, None


def get_oauth2_token(client_id, tenant_id="consumers", email_address=None, force_interactive=False):
    """
    Gets an OAuth2 access token using MSAL (Original CLI implementation)
    """
    if not OAUTH2_AVAILABLE:
        raise ImportError("msal is not installed. Run: pip install msal")
    
    app = build_msal_app(client_id, tenant_id=tenant_id)
    scopes = ["https://outlook.office.com/IMAP.AccessAsUser.All"]
    
    result = None
    if not force_interactive:
        accounts = app.get_accounts()
        if accounts:
            account = next((acc for acc in accounts if acc.get("username") == email_address), accounts[0])
            result = app.acquire_token_silent(scopes, account=account)
    
    if not result:
        print("[*] Opening browser window for OAuth2 authentication...")
        result = app.acquire_token_interactive(scopes, login_hint=email_address)
    
    if "access_token" in result:
        return result["access_token"]
    else:
        error_desc = result.get('error_description', 'Unknown error')
        raise RuntimeError(f"Error getting OAuth2 token: {error_desc}")

