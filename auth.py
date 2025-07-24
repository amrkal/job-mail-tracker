from datetime import datetime
import json
import os
from msal import PublicClientApplication, ConfidentialClientApplication

TOKEN_FILE = "tokens.json"

def load_config():
    scopes_raw = os.environ.get("SCOPES", "Mail.Read")
    scopes = [s for s in scopes_raw.split() if s not in {"offline_access", "openid", "profile"}]
    return {
        "client_id": os.environ.get("CLIENT_ID"),
        "client_secret": os.environ.get("CLIENT_SECRET"),
        "tenant_id": os.environ.get("TENANT_ID"),
        "scopes": scopes,
        "output_excel": os.environ.get("OUTPUT_EXCEL", "job_applications.xlsx"),
        "report_output_folder": os.environ.get("REPORT_OUTPUT_FOLDER", "reports"),
        "user_email": os.environ.get("USER_EMAIL", "kalany.a@hotmail.com")
    }

def save_token(token):
    # Add expires_on if missing (valid for 1 hour by default)
    if "expires_on" not in token:
        token["expires_on"] = int((datetime.utcnow() + timedelta(hours=1)).timestamp())
    with open(TOKEN_FILE, "w") as f:
        json.dump(token, f)

def load_token():
    try:
        with open(TOKEN_FILE, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return None
    

def authenticate_graph(config):
    is_ci = os.environ.get("CI") == "true"

    if is_ci:
        app = ConfidentialClientApplication(
            client_id=config["client_id"],
            client_credential=config["client_secret"],
            authority="https://login.microsoftonline.com/common"
        )
        result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            return result["access_token"]
        raise Exception(f"Authentication failed in CI: {result.get('error_description', result)}")

    # Local interactive login (device code)
    app = PublicClientApplication(
        config["client_id"],
        authority="https://login.microsoftonline.com/common"  # supports personal + org accounts
    )

    # ✅ Try token reuse first
    # ✅ Load cached token and validate expiration
    cached_token = load_token()
    if cached_token and "access_token" in cached_token:
        expires_on = cached_token.get("expires_on")
        if expires_on:
            expiry = int(expires_on)
            now = int(datetime.utcnow().timestamp())
            if now < expiry - 60:
                return cached_token["access_token"]
            else:
                print("🔁 Cached token expired, re-authenticating...")
        else:
            print("⚠️ Cached token has no expiration. Re-authenticating...")


            
    accounts = app.get_accounts()
    result = app.acquire_token_silent(config["scopes"], account=accounts[0]) if accounts else None

    if not result:
        flow = app.initiate_device_flow(scopes=config["scopes"])
        if "user_code" not in flow:
            raise Exception(f"Device flow failed. Response: {flow}")
        print("🔐 DEVICE LOGIN REQUIRED")
        print("Visit this URL in your browser:", flow["verification_uri"])
        print("Enter the code:", flow["user_code"])
        result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        save_token(result)
        return result["access_token"]
    else:
        print("❌ Authentication failed:", result.get("error_description", result))
        # Force reset token file
        if os.path.exists("tokens.json"):
            os.remove("tokens.json")
            print("🧹 Deleted corrupt token file. Please re-run to re-authenticate.")
        raise Exception("Authentication failed and token file was reset.")



# def authenticate_graph(config):
#     is_ci = os.environ.get("CI") == "true"

#     if is_ci:
#         app = ConfidentialClientApplication(
#             client_id=config["client_id"],
#             client_credential=config["client_secret"],
#             authority="https://login.microsoftonline.com/common"
#         )
#         result = app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
#         if "access_token" in result:
#             return result["access_token"]
#         raise Exception(f"Authentication failed in CI: {result.get('error_description', result)}")

#     # Local interactive login (device code)
#     app = PublicClientApplication(
#         config["client_id"],
#         authority="https://login.microsoftonline.com/common"  # use common for personal + org accounts
#     )

#     accounts = app.get_accounts()
#     result = app.acquire_token_silent(config["scopes"], account=accounts[0]) if accounts else None

#     if not result:
#         flow = app.initiate_device_flow(scopes=config["scopes"])
#         if "user_code" not in flow:
#             raise Exception(f"Device flow failed. Response: {flow}")
#         print("🔐 DEVICE LOGIN REQUIRED")
#         print("Visit this URL in your browser:", flow["verification_uri"])
#         print("Enter the code:", flow["user_code"])
#         result = app.acquire_token_by_device_flow(flow)

#     if "access_token" in result:
#         save_token(result)
#         return result["access_token"]
#     else:
#         raise Exception(f"Authentication failed locally: {result.get('error_description', result)}")
