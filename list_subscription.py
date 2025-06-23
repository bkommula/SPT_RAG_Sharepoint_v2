import requests
import os
from dotenv import load_dotenv

# Load .env
load_dotenv()

TENANT_ID = os.getenv("TENANT_ID")
CLIENT_ID = os.getenv("CLIENT_ID")
CLIENT_SECRET = os.getenv("CLIENT_SECRET")

def get_access_token():
    url = f"https://login.microsoftonline.com/{TENANT_ID}/oauth2/v2.0/token"
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": "https://graph.microsoft.com/.default"
    }
    headers = { "Content-Type": "application/x-www-form-urlencoded" }

    res = requests.post(url, data=data, headers=headers)
    res.raise_for_status()
    return res.json()["access_token"]

def list_subscriptions(token):
    headers = {
        "Authorization": f"Bearer {token}"
    }

    res = requests.get("https://graph.microsoft.com/v1.0/subscriptions", headers=headers)

    if res.status_code == 200:
        data = res.json()
        if data.get("value"):
            print("✅ Active Subscriptions:\n")
            for sub in data["value"]:
                print(f"ID: {sub['id']}")
                print(f"Resource: {sub['resource']}")
                print(f"Change Type: {sub['changeType']}")
                print(f"Expiration: {sub['expirationDateTime']}")
                print(f"Notification URL: {sub['notificationUrl']}")
                print("-" * 40)
        else:
            print("ℹ️ No active subscriptions found.")
    else:
        print(f"❌ Failed to fetch subscriptions: {res.status_code}")
        print(res.text)

if __name__ == "__main__":
    token = get_access_token()
    list_subscriptions(token)