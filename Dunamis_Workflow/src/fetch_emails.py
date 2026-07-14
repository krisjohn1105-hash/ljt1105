import msal
import requests
import os
from config import CLIENT_ID, TENANT_ID
import base64
import datetime as dt

today = dt.date.today().strftime('%Y-%m-%d') + 'T00:00:00Z'

# # Your app registration details
CACHE_FILE = "token_cache.bin"
SCOPES = ["Mail.Read"]
SAVE_FOLDER = '../data/input/pre-trade'


def get_token():
    # Load cache from disk
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        cache.deserialize(open(CACHE_FILE, "r").read())

    app = msal.PublicClientApplication(
        CLIENT_ID,
        authority=f"https://login.microsoftonline.com/{TENANT_ID}",
        token_cache=cache
    )

    # Try silent auth first
    accounts = app.get_accounts()
    result = None
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])

    # Only do device flow if no cached token
    if not result:
        print("First time login required...")
        flow = app.initiate_device_flow(scopes=SCOPES)
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)

    # Save updated cache to disk
    if cache.has_state_changed:
        open(CACHE_FILE, "w").write(cache.serialize())

    if "access_token" not in result:
        raise Exception(f"Auth failed: {result.get('error_description')}")

    return result["access_token"]


def get_folder_id(access_token, folder_path):
    headers = {"Authorization": f"Bearer {access_token}"}

    # Start from top-level folders
    url = "https://graph.microsoft.com/v1.0/me/mailFolders"
    response = requests.get(url, headers=headers)
    folder = response.json().get("value", [])

    folder_id = None

    def search_folder(folderName, folders_):
        match = next((f for f in folders_ if f["displayName"].lower() == folderName.lower()), None)

        if not match:
            print(f"Folder '{folderName}' not found. Available at this level:")
            for f in folders_:
                print(f"  - {f['displayName']}")
            return None

        folder_id = match["id"]

        return folder_id

    def child_folders(folder_id):
        # Go one level deeper
        folders = []
        url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/childFolders"
        while url:
            response = requests.get(url, headers=headers, params={'$top': 100})
            data = response.json()
            folders.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        # print(folders)
        return folders

    for folder_name in folder_path:
        folder_id = search_folder(folder_name, folder)
        folder = child_folders(folder_id)

    return folder_id


def get_today_emails(access_token, folder_path, top=25, since=None):
    headers = {"Authorization": f"Bearer {access_token}"}

    folder_id = get_folder_id(access_token, folder_path)
    if not folder_id:
        return []

    url = f"https://graph.microsoft.com/v1.0/me/mailFolders/{folder_id}/messages"
    params = {
        "$top": top,
        "$select": "subject,from,receivedDateTime,bodyPreview,isRead",
        "$orderby": "receivedDateTime desc"
    }

    response = requests.get(url, headers=headers, params=params)

    if not response.ok:
        print("Error:", response.json())
        response.raise_for_status()

    # Add Filters
    email_list = response.json().get("value", [])
    if since:
        email_list = [e for e in email_list if e["receivedDateTime"] >= since]

    return get_latest_emails_per_sender_subject(email_list)


def get_latest_emails_per_sender_subject(emails):
    latest = {}

    for email in emails:
        sender = email["from"]["emailAddress"]["address"].lower() #.split("@").split('.')[0] # Email Structure of PB
        subject = email["subject"].strip().lower()
        date = email["receivedDateTime"][:10]  # YYYY-MM-DD
        key = (sender, subject, date)

        # Keep only the most recent email per key
        if key not in latest or email["receivedDateTime"] > latest[key]["receivedDateTime"]:
            latest[key] = email

    return list(latest.values())


def download_attachments(access_token, message_id, save_folder):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/me/messages/{message_id}/attachments"

    response = requests.get(url, headers=headers)
    if not response.ok:
        print("Error fetching attachments:", response.json())
        return

    attachments = response.json().get("value", [])
    for attachment in attachments:
        if attachment.get("@odata.type") != "#microsoft.graph.fileAttachment":
            continue

        filename = attachment["name"]
        file_data = base64.b64decode(attachment["contentBytes"])
        save_path = os.path.join(save_folder, filename)

        with open(save_path, "wb") as f:
            f.write(file_data)
    return


# Run
access_token = get_token()
emails = get_today_emails(access_token, folder_path=['Inbox', 'Trade Files', 'Pre'], since=today)

for email in emails:
    download_attachments(access_token, email["id"], SAVE_FOLDER)
    print(f"Fetched:\n{email['receivedDateTime']} | {email['subject']}")

# Create a function to delete over a week old files

