"""
read_email.py
=============
Fetches emails from Outlook via Microsoft Graph API.

Public API
----------
  fetch_emails_last_n_days(days=2)  → list of raw email dicts
  process_email(message)            → enriched dict with parsed tables + downloads
  process_latest_email()            → backwards-compat: process only the newest email

Each returned dict includes:
  message_id, received_at, title, body_html,
  downloaded_files, parsed_tables
"""

import os
import json
import re
import msal
import requests
from datetime import datetime, timezone, timedelta
from urllib.parse import unquote, urlparse

from classified_html_table_parser import parse_html_tables
import config

# ─────────────────────────────────────────────────────────────────────────────
# Configuration  (all values loaded from config.py)
# ─────────────────────────────────────────────────────────────────────────────
TENANT_ID     = config.GRAPH_TENANT_ID
CLIENT_ID     = config.GRAPH_CLIENT_ID
CLIENT_SECRET = config.GRAPH_CLIENT_SECRET
USER_EMAIL    = config.USER_EMAIL

ATTACHMENT_DIR = config.ATTACHMENT_DIR
AUTHORITY      = config.GRAPH_AUTHORITY
SCOPE          = config.GRAPH_SCOPE

# ─────────────────────────────────────────────────────────────────────────────
# Keyword that must appear somewhere in the subject to qualify as a DSM/NIS
# automation email.  Matching is case-insensitive (so "dsm-nis-booking",
# "DSM-NIS-BOOKING", "Re: DSM-NIS-Booking weekly" all match).
# ─────────────────────────────────────────────────────────────────────────────
DSM_SUBJECT_KEYWORD = config.DSM_SUBJECT_KEYWORD


# ─────────────────────────────────────────────────────────────────────────────
# Auth
# ─────────────────────────────────────────────────────────────────────────────
def _get_access_token() -> str:
    app = msal.ConfidentialClientApplication(
        CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET,
    )
    token = app.acquire_token_for_client(scopes=SCOPE)
    if "access_token" not in token:
        raise RuntimeError(f"MSAL token error: {token}")
    return token["access_token"]


# ─────────────────────────────────────────────────────────────────────────────
# PDF helpers
# ─────────────────────────────────────────────────────────────────────────────
def _extract_pdf_links(html: str) -> list:
    pdfs    = []
    pattern = re.compile(
        r'<a[^>]+href=["\']([^"\']+\.pdf[^"\']*)["\'][^>]*>(.*?)</a>',
        re.IGNORECASE | re.DOTALL,
    )
    for url, text in pattern.findall(html):
        filename = text.strip()
        if not filename.lower().endswith(".pdf"):
            filename = os.path.basename(urlparse(url).path) or "attachment.pdf"
        pdfs.append({"url": unquote(url), "filename": filename})
    return pdfs


def _download_url(url: str, filename: str, subdir: str) -> str:
    os.makedirs(subdir, exist_ok=True)
    path = os.path.join(subdir, filename)
    resp = requests.get(url, stream=True, timeout=60)
    resp.raise_for_status()
    with open(path, "wb") as f:
        for chunk in resp.iter_content(chunk_size=8192):
            if chunk:
                f.write(chunk)
    return path


# ─────────────────────────────────────────────────────────────────────────────
# Subject filter
# ─────────────────────────────────────────────────────────────────────────────
def is_dsm_email(subject: str) -> bool:
    """
    Return True when the subject contains DSM_SUBJECT_KEYWORD anywhere,
    case-insensitively.

    Match   : "AUTO-DSM-NIS-Booking"          ✅
    Match   : "dsm-nis-booking week 34"        ✅
    Match   : "Re: DSM-NIS-BOOKING summary"    ✅
    Match   : "FWD: Dsm-Nis-Booking"           ✅
    No-match: "DSM report"                     ❌
    No-match: "NIS weekly update"              ❌
    """
    return DSM_SUBJECT_KEYWORD.lower() in (subject or "").lower()


def has_required_tables(parsed_tables: dict) -> bool:
    """
    Return True only when BOTH the metadata table AND spv_dsm table are
    present, non-empty, and contain an spv_name value.
    """
    spv  = parsed_tables.get("spv_dsm_table",  [])
    meta = parsed_tables.get("metadata_table", [])
    if not spv or not meta:
        return False
    if not spv[0].get("spv_name", "").strip():
        return False
    return True


# ─────────────────────────────────────────────────────────────────────────────
# Core: process a single raw Graph API message dict
# ─────────────────────────────────────────────────────────────────────────────
def process_email(message: dict, headers: dict, base_url: str) -> dict:
    """
    Given a raw Graph API message object (already fetched), download
    attachments, parse tables, and return an enriched dict.
    """
    message_id  = message["id"]
    received_at = message.get("receivedDateTime", "")
    subject     = message.get("subject", "")
    body_html   = message.get("body", {}).get("content", "")

    # per-message attachment directory so files don't collide across emails
    safe_id    = re.sub(r"[^a-zA-Z0-9_-]", "_", message_id)[:40]
    att_subdir = os.path.join(ATTACHMENT_DIR, safe_id)
    os.makedirs(att_subdir, exist_ok=True)

    downloaded_files = []

    # ── 1. Outlook PDF attachments (normal + inline) ──────────────────────
    try:
        att_resp = requests.get(
            f"{base_url}/messages/{message_id}/attachments",
            headers=headers,
        )
        att_resp.raise_for_status()
        for att in att_resp.json().get("value", []):
            name = att.get("name", "")
            if not name.lower().endswith(".pdf"):
                continue
            att_id    = att["id"]
            file_resp = requests.get(
                f"{base_url}/messages/{message_id}/attachments/{att_id}/$value",
                headers=headers,
            )
            if file_resp.status_code == 200:
                path = os.path.join(att_subdir, name)
                with open(path, "wb") as f:
                    f.write(file_resp.content)
                downloaded_files.append(path)
    except Exception as exc:
        print(f"  [WARN] Attachment download failed for {message_id[:20]}… : {exc}")

    # ── 2. PDF links embedded in HTML body ────────────────────────────────
    for pdf in _extract_pdf_links(body_html):
        try:
            path = _download_url(pdf["url"], pdf["filename"], att_subdir)
            downloaded_files.append(path)
        except Exception as exc:
            print(f"  [WARN] PDF link download failed: {pdf['url']} → {exc}")

    # ── 3. Parse HTML tables ──────────────────────────────────────────────
    parsed_tables = parse_html_tables(body_html)

    return {
        "message_id":       message_id,
        "received_at":      received_at,
        "title":            subject,
        "body_html":        body_html,
        "downloaded_files": downloaded_files,
        "parsed_tables":    parsed_tables,
        "attachment_dir":   att_subdir,
    }


# ─────────────────────────────────────────────────────────────────────────────
# Public: fetch ALL emails from the last N days
# ─────────────────────────────────────────────────────────────────────────────
def fetch_emails_last_n_days(days: int = 2) -> list:
    """
    Return a list of raw Graph API message dicts received in the last `days` days.
    Only fetches subject + body (no attachments yet).
    Pagination is followed so ALL matching emails are returned.
    """
    token    = _get_access_token()
    headers  = {"Authorization": f"Bearer {token}"}
    base_url = f"https://graph.microsoft.com/v1.0/users/{USER_EMAIL}"

    since_dt = (datetime.now(timezone.utc) - timedelta(days=days)).strftime(
        "%Y-%m-%dT%H:%M:%SZ"
    )

    url = (
        f"{base_url}/messages"
        f"?$filter=receivedDateTime ge {since_dt}"
        "&$orderby=receivedDateTime desc"
        "&$select=id,subject,body,receivedDateTime"
        "&$top=50"          # fetch up to 50 per page; follow @odata.nextLink
    )

    messages = []
    while url:
        resp = requests.get(url, headers=headers)
        resp.raise_for_status()
        data     = resp.json()
        messages.extend(data.get("value", []))
        url = data.get("@odata.nextLink")   # None when last page

    print(f"[EMAIL] Fetched {len(messages)} email(s) from the last {days} day(s).")
    return messages, headers, base_url


# ─────────────────────────────────────────────────────────────────────────────
# Public: process only the newest email (backwards-compat for direct runs)
# ─────────────────────────────────────────────────────────────────────────────
def process_latest_email() -> dict:
    messages, headers, base_url = fetch_emails_last_n_days(days=1)
    if not messages:
        return {}
    return process_email(messages[0], headers, base_url)


# ─────────────────────────────────────────────────────────────────────────────
# Standalone test
# ─────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    msgs, hdrs, base = fetch_emails_last_n_days(days=2)
    for m in msgs:
        result = process_email(m, hdrs, base)
        print(json.dumps({
            "message_id": result["message_id"],
            "title":      result["title"],
            "has_tables": has_required_tables(result["parsed_tables"]),
        }, indent=2))