import json
import requests
import msal
import config

# All credentials and SharePoint settings loaded from config.py
TENANT_ID     = config.SP_TENANT_ID
CLIENT_ID     = config.SP_CLIENT_ID
CLIENT_SECRET = config.SP_CLIENT_SECRET

TENANT_HOST = config.TENANT_HOST
SITE_NAME   = config.SITE_NAME
DRIVE_NAME  = config.DRIVE_NAME
FOLDER_PATH = config.FOLDER_PATH
FILE_NAME   = config.FILE_NAME

AUTHORITY = config.SP_AUTHORITY
SCOPES    = config.GRAPH_SCOPE

def get_token():
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET
    )
    return app.acquire_token_for_client(scopes=SCOPES)["access_token"]

def get_json(url, headers):
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json()

def read_vendor_data(sheet_name: str) -> dict:
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}

    site = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}",
        headers
    )
    site_id = site["id"]

    drives = get_json(
        f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives",
        headers
    )["value"]

    drive_id = next(d["id"] for d in drives if d["name"] == DRIVE_NAME)

    item = get_json(
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{FOLDER_PATH}/{FILE_NAME}",
        headers
    )

    range_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item['id']}"
        f"/workbook/worksheets('{sheet_name}')/range(address='A2:B7')"
    )

    rows = get_json(range_url, headers).get("values", [])

    return {r[0]: r[1] for r in rows if len(r) >= 2}
