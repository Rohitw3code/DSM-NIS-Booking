import os
import io
import sys
import json
import time
import requests
import msal
import math
import config

# =========================
# CONFIG  (all values loaded from config.py)
# =========================
TENANT_ID     = config.SP_TENANT_ID
CLIENT_ID     = config.SP_CLIENT_ID
CLIENT_SECRET = config.SP_CLIENT_SECRET

TENANT_HOST  = config.TENANT_HOST
SITE_NAME    = config.SITE_NAME
DRIVE_NAME   = config.DRIVE_NAME
FOLDER_PATH  = config.FOLDER_PATH
FILE_NAME    = config.FILE_NAME
TARGET_SHEET = config.TARGET_SHEET

AUTHORITY = config.SP_AUTHORITY
SCOPES    = config.GRAPH_SCOPE

# =========================
# SHEET LAYOUT
# =========================
# Columns D to O (12 columns total):
#   D  = Week_No                            ← week_no (from metadata_table)
#   E  = (untouched / blank)
#   F  = Week                               ← "from_date TO to_date" (from metadata_table)
#   G  = (untouched / blank)
#   H  = DSM_Charges                        ← total_dsm_charges_payable (from spv_dsm_table)
#   I  = Revenue_Diff_amount                ← revenue_diff (from spv_dsm_table)
#   J  = (untouched / blank)
#   K  = (untouched / blank)
#   L  = Due_Date                           ← due_date (from metadata_table)
#   M  = (untouched / blank)
#   N  = (untouched / blank)
#   O  = (untouched / blank)

HEADER_ROW      = 2
START_DATA_ROW  = HEADER_ROW + 1
START_COL, END_COL = "D", "O"
MAX_SCAN_ROW    = 5000

EXPECTED_COLS = ord(END_COL) - ord(START_COL) + 1  # 12 columns D..O


# =========================
# AUTH
# =========================
def get_token(retries: int = 0, backoff_sec: float = 1.5) -> str:
    app = msal.ConfidentialClientApplication(
        client_id=CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )
    last_err = None
    for attempt in range(retries + 1):
        result = app.acquire_token_for_client(scopes=SCOPES)
        if "access_token" in result:
            return result["access_token"]
        last_err = result
        if result.get("error") == "invalid_client":
            break
        if attempt < retries:
            time.sleep(backoff_sec * (2 ** attempt))
    raise RuntimeError(
        "Failed to acquire token. If error is 'invalid_client', verify the client secret.\n"
        f"MSAL response:\n{json.dumps(last_err, indent=2)}"
    )


# =========================
# HTTP HELPERS
# =========================
def get_json(url: str, headers: dict) -> dict:
    r = requests.get(url, headers=headers)
    try:
        r.raise_for_status()
    except Exception:
        print("GET failed:", url)
        print("Status:", r.status_code)
        print(r.text[:2048])
        raise
    try:
        return r.json()
    except Exception:
        return {}


def patch_json(url: str, headers: dict, payload: dict) -> dict:
    r = requests.patch(url, headers={**headers, "Content-Type": "application/json"}, json=payload)
    try:
        r.raise_for_status()
    except Exception:
        print("PATCH failed:", url)
        print("Status:", r.status_code)
        print(r.text[:2048])
        raise
    try:
        return r.json()
    except Exception:
        return {}


# =========================
# GRAPH HELPERS
# =========================
def resolve_site_and_drive(headers_auth: dict):
    site_url = f"https://graph.microsoft.com/v1.0/sites/{TENANT_HOST}:/sites/{SITE_NAME}"
    site = get_json(site_url, headers_auth) or {}
    site_id = site.get("id")
    if not site_id:
        raise RuntimeError("Failed to resolve Site ID.")

    drives_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives"
    drives = get_json(drives_url, headers_auth) or {}
    all_drives = drives.get("value") or []

    drive_id = None
    for d in all_drives:
        if d.get("name") == DRIVE_NAME:
            drive_id = d.get("id")
            break
    if not drive_id:
        default_drive = get_json(
            f"https://graph.microsoft.com/v1.0/sites/{site_id}/drive", headers_auth
        ) or {}
        drive_id = default_drive.get("id")
    if not drive_id:
        raise RuntimeError("Failed to resolve Drive ID.")
    return site_id, drive_id


def resolve_item_id_by_path(drive_id: str, item_path: str, headers_auth: dict) -> str:
    item_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{item_path}"
    item = get_json(item_url, headers_auth) or {}
    item_id = item.get("id")
    if not item_id:
        raise RuntimeError("Unable to resolve item ID for the Excel file.")
    return item_id


def ensure_sheet_name(drive_id: str, item_id: str, headers_auth: dict, target: str) -> str:
    ws_url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets"
    )
    sheets = get_json(ws_url, headers_auth) or {}
    names = [s.get("name") for s in (sheets.get("value") or []) if s.get("name")]
    if target in names:
        return target
    cand = [n for n in names if target.lower() in n.lower()]
    if cand:
        return cand[0]
    raise RuntimeError(f"Worksheet '{target}' not found. Available: {names}")


def read_col_D_values(
    drive_id: str, item_id: str, sheet_name: str, headers_auth: dict, max_row: int = MAX_SCAN_ROW
):
    address = f"D1:D{max_row}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{address}')"
    )
    resp = get_json(url, headers_auth) or {}
    col_vals = resp.get("values")
    if not isinstance(col_vals, list):
        return []
    return col_vals


def write_row_D_to_O(
    drive_id: str,
    item_id: str,
    sheet_name: str,
    row_index: int,
    values: list,
    headers_auth: dict,
):
    address = f"{START_COL}{row_index}:{END_COL}{row_index}"
    url = (
        f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
        f"/workbook/worksheets('{sheet_name}')/range(address='{address}')"
    )
    payload = {"values": [values]}
    patch_json(url, headers_auth, payload)
    print(f"[OK] Wrote row at {address}")

    fmt_url = url + "/format"
    patch_json(fmt_url, headers_auth, {
        "horizontalAlignment": "Center",
        "verticalAlignment": "Center",
    })

    font_url = fmt_url + "/font"
    patch_json(font_url, headers_auth, {"size": 12})
    print(f"[OK] Applied font size 12 and centre alignment to {address}")


# =========================
# ROW COMPUTATION
# =========================
def is_blank_cell(val) -> bool:
    return (val is None) or (isinstance(val, str) and val.strip() == "")


def try_parse_int(val):
    if val is None:
        return None
    if isinstance(val, int):
        return int(val)
    if isinstance(val, float):
        if math.isnan(val):
            return None
        return int(val) if float(val).is_integer() else None
    if isinstance(val, str):
        s = val.strip()
        if s.isdigit():
            try:
                return int(s)
            except Exception:
                return None
    return None


def find_first_blank_row_in_D(col_vals: list, start_row: int, max_row: int) -> int:
    n = len(col_vals)
    row = start_row
    while row <= max_row:
        if row <= n:
            row_vals = col_vals[row - 1]
            cell_val = (
                row_vals[0]
                if (isinstance(row_vals, list) and len(row_vals) > 0)
                else None
            )
            if is_blank_cell(cell_val):
                return row
        else:
            return row
        row += 1
    return max_row


def compute_next_week_no(col_vals: list, start_row: int) -> int:
    max_seen = 0
    for idx, row in enumerate(col_vals, start=1):
        if idx < start_row:
            continue
        cell_val = (
            row[0] if (isinstance(row, list) and len(row) > 0) else None
        )
        v = try_parse_int(cell_val)
        if v is not None and v > max_seen:
            max_seen = v
    return max_seen + 1 if max_seen >= 0 else 1


# =========================
# VALUE HELPERS
# =========================
def safe_float(val):
    """Convert a string like '1,016.00' or '4517646' to float. Returns 0.0 on failure."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    cleaned = str(val).replace(",", "").strip()
    try:
        return float(cleaned)
    except (ValueError, TypeError):
        return 0.0


def build_row_values(
    metadata_row: dict,
    spv_dsm_row: dict,
    checklist_id: str = "",
) -> list:
    """
    Build the 12-value list for columns D..O.

    Column mapping (0-based index → spreadsheet column):
      0  D  Week_No                     ← week_no
      1  E  Week                        ← "from_date to to_date"
      2  F  DSM Statement Published Date
      3  G  (blank)
      4  H  DSM_Charges                 ← total_dsm_charges_payable
      5  I  Revenue_Diff                ← revenue_diff
      6  J  Net Payable/Receivable      ← net_payable_receivable
      7  K  (blank)
      8  L  Due_Date                    ← due_date
      9  M  NIS Checklist ID            ← checklist_id (captured from portal)
     10  N  (blank)
     11  O  (blank)
    """
    # ---- metadata fields ----
    week_no   = str(metadata_row.get("week_no", "")).strip()
    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date", "")).strip()
    due_date  = str(metadata_row.get("due_date", "")).strip()
    dsm_statement_published_date = str(metadata_row.get("dsm_statement_published_date", "")).strip()

    # Column E: "DD-MM-YYYY to DD-MM-YYYY"
    week_range = f"{from_date} to {to_date}" if from_date and to_date else ""

    # ---- spv dsm fields ----
    dsm_charges  = safe_float(spv_dsm_row.get("total_dsm_charges_payable"))
    revenue_diff = safe_float(spv_dsm_row.get("revenue_diff"))
    net_payable_receivable = safe_float(spv_dsm_row.get("net_payable_receivable"))

    return [
        week_no,        # D  – Week_No
        week_range,     # E  – Week ("from_date to to_date")
        dsm_statement_published_date,  # F  – DSM Statement Published Date
        "",             # G  – (blank)
        dsm_charges,    # H  – DSM_Charges
        revenue_diff,   # I  – Revenue_Diff
        net_payable_receivable,        # J  – Net Payable/Receivable
        "",             # K  – (blank)
        due_date,       # L  – Due_Date
        checklist_id or "",            # M  – NIS Checklist ID
        "",             # N  – (blank)
        "",             # O  – (blank)
    ]


# =========================
# MAIN PUBLIC FUNCTION
# =========================
def add_incremental_week_row(
    metadata_row: dict,
    spv_dsm_row: dict,
    sheet_name: str = TARGET_SHEET,
    checklist_id: str = "",
):
    """
    Adds one new row to the NIS booking Excel sheet using live data from the
    parsed email tables.

    Args:
        metadata_row  : first dict from parsed_tables["metadata_table"]
        spv_dsm_row   : matching dict from parsed_tables["spv_dsm_table"]
        sheet_name    : target worksheet name (defaults to TARGET_SHEET constant)
        checklist_id  : NIS checklist/booking number captured from the portal
                        (written to column M)
    """
    token = get_token()
    headers_auth = {"Authorization": f"Bearer {token}"}

    _, drive_id = resolve_site_and_drive(headers_auth)
    item_path   = f"{FOLDER_PATH}/{FILE_NAME}".strip("/")
    item_id     = resolve_item_id_by_path(drive_id, item_path, headers_auth)
    resolved_sheet = ensure_sheet_name(drive_id, item_id, headers_auth, sheet_name)

    col_vals   = read_col_D_values(drive_id, item_id, resolved_sheet, headers_auth, max_row=MAX_SCAN_ROW)
    target_row = find_first_blank_row_in_D(col_vals, START_DATA_ROW, MAX_SCAN_ROW)

    row_values = build_row_values(metadata_row, spv_dsm_row, checklist_id=checklist_id)

    if len(row_values) != EXPECTED_COLS:
        raise ValueError(
            f"row_values must have exactly {EXPECTED_COLS} items (D..O); got {len(row_values)}."
        )

    write_row_D_to_O(drive_id, item_id, resolved_sheet, target_row, row_values, headers_auth)
    print(f"[INFO] Inserted WEEK NO={metadata_row.get('week_no', '?')} at row {target_row} on sheet '{resolved_sheet}'")


# =========================
# STANDALONE ENTRY POINT
# =========================
def main():
    """
    Standalone test: uses hard-coded example data so the script can be run
    independently without the full email-reading pipeline.
    """
    print("\n=== Insert Row with Incremental WEEK NO (D..O) – standalone test ===")
    print("TENANT_HOST :", TENANT_HOST)
    print("SITE_NAME   :", SITE_NAME)
    print("DRIVE_NAME  :", DRIVE_NAME)
    print("FOLDER_PATH :", FOLDER_PATH)
    print("FILE_NAME   :", FILE_NAME)
    print("TARGET_SHEET:", TARGET_SHEET)
    print("=======================================================================\n")

    # Example data matching the email parser output format
    example_metadata = {
        "week_no":                       "34",
        "from_date":                     "17-11-2025",
        "to_date":                       "23-11-2025",
        "dsm_statement_published_date":  "2-Dec-25",
        "due_date":                      "12-Dec-25",
    }

    example_spv_dsm = {
        "spv_name":                  "AWEMP1PL",
        "total_dsm_charges_payable":  1016,
        "drawl_charges_payable":      0,
        "revenue_diff":               4517646,
        "revenue_loss":               0,
        "net_payable_receivable":     4518662.00,
    }

    add_incremental_week_row(
        metadata_row=example_metadata,
        spv_dsm_row=example_spv_dsm,
        sheet_name=TARGET_SHEET,
        checklist_id="NIS-2025-0001234",   # example checklist ID
    )


if __name__ == "__main__":
    main()