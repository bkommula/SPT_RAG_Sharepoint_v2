import os
import requests
import json
from datetime import datetime, timedelta, timezone
from dotenv import load_dotenv

load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")
SITE_URL_NEW = os.getenv("SITE_URL_NEW")
NGROK_URL = os.getenv("NGROK_URL")  # example: https://abc123.ngrok-free.app

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def register_subscription():
    token = get_access_token()
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }

    site_name = SITE_URL_NEW.split("/")[-1]

    # STEP 1: Get Site ID
    site_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/root:/sites/{site_name}", headers=headers)
    site_res.raise_for_status()
    site_id = site_res.json()["id"]

    # STEP 2: Get Drive ID
    drive_res = requests.get(f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers=headers)
    drive_res.raise_for_status()
    drive_id = drive_res.json()["id"]

    # STEP 3: Register subscription on the root of the drive
    expire = (datetime.now(timezone.utc) + timedelta(minutes=42300)).isoformat()    #29.4 days

    payload = {
        "changeType": "updated",
        "notificationUrl": NGROK_URL + "/webhook",
        "resource": f"drives/{drive_id}/root",  # ✅ CORRECT & SUPPORTED
        "expirationDateTime": expire,
        "clientState": "test123"
    }

    res = requests.post("https://graph.microsoft.com/v1.0/subscriptions", headers=headers, data=json.dumps(payload))
    if res.status_code == 201:
        print("✅ Subscription created successfully:")
        print(json.dumps(res.json(), indent=2))
    else:
        print("❌ Failed to create subscription:", res.status_code)
        print(res.text)

if __name__ == "__main__":
    register_subscription()
