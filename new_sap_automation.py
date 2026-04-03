import os
import time
import subprocess
import win32com.client
from datetime import datetime
import config

# -------------------------------------------------
# Helpers: resilient UI actions & waits
# -------------------------------------------------
def wait(seconds: float = 1.0):
    time.sleep(seconds)

def try_press(session, ids):
    """Try pressing the first existing control from a list of IDs."""
    for _id in ids:
        try:
            session.findById(_id).press()
            return True
        except Exception:
            pass
    return False

def try_select(session, ids):
    """Try selecting (e.g., radio/tab) the first existing control from a list of IDs."""
    for _id in ids:
        try:
            session.findById(_id).select()
            return True
        except Exception:
            pass
    return False

def try_set_text(session, _id, value):
    try:
        session.findById(_id).text = str(value)
        return True
    except Exception:
        return False

def try_set_key(session, _id, key):
    """For combo boxes where .key is the right property."""
    try:
        session.findById(_id).key = str(key)
        return True
    except Exception:
        return False

def send_enter(session):
    try:
        session.findById("wnd[0]").sendVKey(0)
        return True
    except Exception:
        return False

def handle_default_popup_ok(session):
    """Press OK/Continue on common popups if they appear (wnd[1])."""
    try:
        session.findById("wnd[1]/tbar[0]/btn[0]").press()  # OK / Continue
        return True
    except Exception:
        return False

def handle_yes_popup(session):
    """Press Yes/Enter on confirmation popup (wnd[1]) if present."""
    return try_press(session, [
        "wnd[1]/tbar[0]/btn[8]",
        "wnd[1]/tbar[0]/btn[0]",
    ])

# -------------------------------------------------
# SAP bootstrap: launch, attach, connect, login
# -------------------------------------------------

def launch_sap_logon(saplogon_path: str):
    """Launch SAP Logon if not already running."""
    try:
        subprocess.Popen(saplogon_path)
        wait(5)
    except Exception:
        pass

def attach_to_scripting(system_name: str):
    SapGuiAuto = win32com.client.GetObject("SAPGUI")
    application = SapGuiAuto.GetScriptingEngine
    connection = application.OpenConnection(system_name, True)
    session = connection.Children(0)
    return application, connection, session

def login(session, client: str, username: str, password: str, language: str = "EN"):
    session.findById("wnd[0]/usr/txtRSYST-MANDT").text = client
    session.findById("wnd[0]/usr/txtRSYST-BNAME").text = username
    session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = password
    session.findById("wnd[0]/usr/txtRSYST-LANGU").text = language
    send_enter(session)
    wait(3)

def go_to_tcode(session, tcode: str):
    session.findById("wnd[0]/tbar[0]/okcd").text = tcode
    send_enter(session)
    wait(2)

# -------------------------------------------------
# Task flows (configurable)
# -------------------------------------------------
def select_non_po_and_create(session):
    """
    Select the Non-PO Based option (if present) then trigger Create.
    Multiple fallbacks are attempted as different systems expose different IDs.
    """
    try_select(session, [
        "wnd[0]/usr/radNON_PO_BASED",
        "wnd[1]/usr/radNON_PO_BASED",
        "wnd[0]/usr/radRB3",
    ])

    created = try_press(session, [
        "wnd[0]/usr/btnCREATE_CHECKLIST",
        "wnd[0]/usr/btnCREATE",
        "wnd[0]/tbar[1]/btn[8]",
    ])
    print("created")
    handle_cancel_popup(session)
    if created:
        wait(0.8)


def fill_non_po_form(session, params: dict):
    """Fill fields for the Non-PO flow (config-driven)."""
    try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-CHK_DOC_TYP", params.get("CHK_DOC_TYP", ""))
    try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-BUKRS", params.get("BUKRS", ""))
    try_set_key(session,  "wnd[0]/usr/cmbZFIBPT11001-DIG_SIGN_INV", params.get("DIG_SIGN_INV", ""))

    if params.get("GJAHR"):
        try_set_text(session, "wnd[0]/usr/txtZFIBPT11001-GJAHR", params["GJAHR"])
        send_enter(session)
        wait(0.7)
    if params.get("LIFNR"):
        try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-LIFNR", params["LIFNR"])

    try_press(session, ["wnd[0]/tbar[1]/btn[8]"])
    handle_default_popup_ok(session)

    try:
        tbl = session.findById(
            "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1/"
            "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9012/tblSAPMZFIBP11001ZTC_9013"
        )
        tbl.verticalScrollbar.position = 1
    except Exception:
        pass

    if params.get("DEB_NOT_REF"):
        try_set_text(session, "wnd[0]/usr/txtZFIBPT11001-DEB_NOT_REF", params["DEB_NOT_REF"])

def apply_subcategory_checkbox_table_and_save(session, params: dict):
    """
    Implements the additional steps:
    - Fill Debit Note Date and Document Sub-Category
    - Go to the 'Documents' tab and tick a checkbox
    - Return to the first tab and fill table fields (Amount, Cost Center, Plant)
    - (Optional) Preview
    - Save and confirm popups
    """
    if params.get("DEB_NOT_DATE"):
        try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-DEB_NOT_DAT", params["DEB_NOT_DATE"])
    if params.get("DOC_SUB_CAT"):
        try_set_text(session, "wnd[0]/usr/ctxtZFIBPT11001-DOC_SUB_CAT", params["DOC_SUB_CAT"])

    try:
        session.findById("wnd[0]/usr/txtGV_DOC_SUB_TEXT").setFocus()
        session.findById("wnd[0]/usr/txtGV_DOC_SUB_TEXT").caretPosition = 0
    except Exception:
        pass

    # Navigate to checkbox tab (FC3) and tick the checkbox
    try_select(session, ["wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC3"])
    try:
        chk_id = (
            "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC3/"
            "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9005/tblSAPMZFIBP11001ZTC_9005/"
            "chkGW_DOCS-CHECK[1,0]"
        )
        cell = session.findById(chk_id)
        cell.selected = True
        cell.setFocus()
    except Exception:
        pass

    try_select(session, ["wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC2"])
    try_select(session, ["wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1"])

    # Fill table row (row 0) fields
    base_tbl = (
        "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1/"
        "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9012/tblSAPMZFIBP11001ZTC_9013/"
    )

    if params.get("AMOUNT"):
        try_set_text(session, base_tbl + "txtGW_RE_MODULE-DEB_NOT_AMT[3,0]", params["AMOUNT"])
    if params.get("COST_CENTER"):
        try_set_text(session, base_tbl + "ctxtGW_RE_MODULE-KOSTL[4,0]", params["COST_CENTER"])
    if params.get("PLANT"):
        try_set_text(session, base_tbl + "ctxtGW_RE_MODULE-WERKS[5,0]", params["PLANT"])
        try:
            session.findById(base_tbl + "ctxtGW_RE_MODULE-WERKS[5,0]").setFocus()
            session.findById(base_tbl + "ctxtGW_RE_MODULE-WERKS[5,0]").caretPosition = len(str(params["PLANT"]))
        except Exception:
            pass

    try:
        tbl = session.findById(
            "wnd[0]/usr/tabsZTABSTRIP/tabpZTABSTRIP_FC1/"
            "ssubZTABSTRIP_SCA:SAPMZFIBP11001:9012/tblSAPMZFIBP11001ZTC_9013"
        )
        tbl.verticalScrollbar.position = 0
    except Exception:
        pass

    wait(0.5)

    if str(params.get("DO_PREVIEW", "true")).lower() in ("1", "true", "yes", "y"):
        preview_pressed = try_press(session, [
            "wnd[0]/usr/btnPREVIEW",
            "wnd[0]/tbar[1]/btn[5]",
            "wnd[0]/tbar[0]/btn[5]",
        ])
        if preview_pressed:
            wait(1.0)
            handle_default_popup_ok(session)

    saved = try_press(session, [
        "wnd[0]/tbar[0]/btn[11]",
        "wnd[0]/tbar[1]/btn[11]",
        "wnd[0]/usr/btnSAVE",
    ])
    if saved:
        wait(0.8)
        handle_yes_popup(session)
        handle_default_popup_ok(session)

def handle_cancel_popup(session):
    """Press 'Cancel' on a typical modal (wnd[1]) if it appears."""
    try:
        session.findById("wnd[1]").sendVKey(12)  # F12 = Cancel
        return True
    except Exception:
        pass

    for cancel_id in [
        "wnd[1]/tbar[0]/btn[12]",
        "wnd[1]/tbar[0]/btn[2]",
        "wnd[1]/tbar[0]/btn[21]",
    ]:
        try:
            session.findById(cancel_id).press()
            return True
        except Exception:
            continue

    try:
        session.findById("wnd[1]")
        return False
    except Exception:
        return True

def open_print_dialog(session):
    """
    Opens SAP print dialog safely without relying on preview shell
    """
    # Try toolbar print button
    if try_press(session, [
        "wnd[0]/tbar[0]/btn[86]",   # Print
        "wnd[0]/tbar[1]/btn[5]"     # Alternate print
    ]):
        time.sleep(2)
        return True

    # Try menu path (List → Print)
    try:
        session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
        time.sleep(2)
        return True
    except Exception:
        pass

    raise RuntimeError("❌ Unable to open SAP Print dialog")


def save_pdf(session):
    """
    Trigger PDF! in SAP → SAP auto-saves the PDF to the Windows Temp folder.
    Find the most recently created/modified PDF in that folder and move it
    to the target directory with a timestamped filename.
    """
    import glob
    import shutil

    temp_dir   = config.SAP_PDF_TEMP_DIR
    target_dir = config.SAP_PDF_TARGET_DIR
    os.makedirs(target_dir, exist_ok=True)

    # ------------------------------------------------------------------
    # STEP 1 – Snapshot the newest PDF in Temp BEFORE triggering PDF!
    #          so we can detect exactly which file SAP creates afterwards.
    # ------------------------------------------------------------------
    def get_newest_pdf(folder):
        pdfs = glob.glob(os.path.join(folder, "*.pdf")) + \
               glob.glob(os.path.join(folder, "**", "*.pdf"), recursive=True)
        if not pdfs:
            return None, 0
        newest = max(pdfs, key=os.path.getmtime)
        return newest, os.path.getmtime(newest)

    _, mtime_before = get_newest_pdf(temp_dir)
    snapshot_time   = time.time()

    # ------------------------------------------------------------------
    # STEP 2 – Trigger PDF! in SAP (Enter after setting okcd)
    # ------------------------------------------------------------------
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = "PDF!"
        session.findById("wnd[0]").sendVKey(0)   # Enter
        print("📄 PDF! triggered — waiting for SAP to write file to Temp...")
        time.sleep(5)                             # give SAP time to render & write
    except Exception as e:
        print(f"❌ Failed to trigger PDF!: {e}")
        return

    # ------------------------------------------------------------------
    # STEP 3 – Poll Temp folder until a new PDF appears (up to 30 s)
    # ------------------------------------------------------------------
    new_pdf = None
    for _ in range(30):
        candidate, mtime_after = get_newest_pdf(temp_dir)
        if candidate and mtime_after > mtime_before and os.path.getmtime(candidate) >= snapshot_time:
            new_pdf = candidate
            break
        time.sleep(1)

    if not new_pdf:
        print("❌ No new PDF found in Temp folder after 30 seconds.")
        return

    print(f"📥  New PDF detected in Temp: {new_pdf}")

    # ------------------------------------------------------------------
    # STEP 4 – Move & rename the file to the target directory
    # ------------------------------------------------------------------
    filename    = os.path.basename(new_pdf)
    destination = os.path.join(target_dir, filename)

    try:
        shutil.copy2(new_pdf, destination)
        print(f"✅ PDF copied successfully → {destination}")
        return destination   # return the destination path so caller can rename it
    except Exception as e:
        print(f"❌ Failed to copy PDF: {e}")
        return None



def close_pdf_preview_and_exit(session):
    """
    1. Close the PDF Preview popup (wnd[1]) — equivalent to clicking X on the popup.
    2. Maximize wnd[0] and press btn[15] (Back/Exit) to return to the checklist screen.
       Mirrors the VBScript:
           session.findById("wnd[0]").maximize
           session.findById("wnd[0]/tbar[0]/btn[15]").press
    """
    # --- Close the PDF Preview popup (wnd[1]) ---
    try:
        session.findById("wnd[1]").close()
        print("🪟  PDF Preview popup closed.")
        time.sleep(1)
    except Exception:
        # Fallback: send F3 / F12 to dismiss
        try:
            session.findById("wnd[1]").sendVKey(12)   # F12 = Cancel/Close
            time.sleep(1)
            print("🪟  PDF Preview popup dismissed via F12.")
        except Exception:
            pass

    # --- Maximize main window and press btn[15] (Back) ---
    try:
        session.findById("wnd[0]").maximize()
    except Exception:
        pass

    try:
        session.findById("wnd[0]/tbar[0]/btn[15]").press()
        print("🔙  Pressed btn[15] (Back/Exit) on wnd[0].")
        time.sleep(2)
    except Exception as e:
        print(f"⚠️  Could not press btn[15]: {e}")


def capture_checklist_data(session):
    """
    Implements the exact VBS sequence, then reads the checklist summary data.

    VBS order (per spec):
        session.findById("wnd[1]").close                   ← close PDF-preview popup
        session.findById("wnd[0]/tbar[0]/btn[3]").press   ← Back button
        session.findById("wnd[0]/sbar").doubleClick        ← opens checklist-number popup

    The checklist number is copied from the popup that opens on doubleClick.
    Fields captured afterwards (standard ZDCC / checklist output area):
      • Checklist Number   (from the doubleClick popup, with sbar-text fallback)
      • Plant, Company Code, Vendor, GST fields, Value, Doc No(s)
    """
    import time
    import re

    # ── STEP 1: Close the PDF-preview popup (wnd[1]) ─────────────────────────
    try:
        session.findById("wnd[1]").close()
        print("🪟  wnd[1] closed (PDF-preview popup).")
        time.sleep(1)
    except Exception as e:
        print(f"ℹ️  wnd[1] not present or already closed: {e}")

    # ── STEP 2: Press Back (btn[3]) — return to the saved checklist screen ────
    try:
        session.findById("wnd[0]/tbar[0]/btn[3]").press()
        print("🔙 Back button (btn[3]) pressed — returning to checklist screen.")
        time.sleep(2)
    except Exception as e:
        print(f"⚠️  Could not press Back btn[3]: {e}")

    # ── STEP 3: Double-click the status bar — checklist number is in this popup
    checklist_number = None
    try:
        session.findById("wnd[0]/sbar").doubleClick()
        print("🖱️  Status bar double-clicked — waiting for checklist popup...")
        time.sleep(1.5)

        # Try every common SAP message-popup text element (wnd[1] re-opens here)
        popup_text_ids = [
            "wnd[1]/usr/txtMESSAGE",
            "wnd[1]/usr/txtMESSTXT-STEXT",
            "wnd[1]/usr/subSCREEN:SAPLMSG_DISPLAY:0100/txtMESSAGE",
            "wnd[1]/usr/subSCREEN_HEADER:SAPLMSG_DISPLAY:0100/txtMESSAGE",
        ]
        for ptxt_id in popup_text_ids:
            try:
                popup_text = session.findById(ptxt_id).text.strip()
                if popup_text:
                    m = re.search(r'\b(\d{7,})\b', popup_text)
                    if m:
                        checklist_number = m.group(1)
                        print(f"✅ Checklist number read from doubleClick popup: {checklist_number}")
                        break
            except Exception:
                continue

        # Fallback: try the popup window title / text property itself
        if not checklist_number:
            try:
                wnd1_text = session.findById("wnd[1]").text.strip()
                m = re.search(r'\b(\d{7,})\b', wnd1_text)
                if m:
                    checklist_number = m.group(1)
                    print(f"✅ Checklist number read from popup title: {checklist_number}")
            except Exception:
                pass

        # Close the checklist-number popup
        for close_id in [
            "wnd[1]/tbar[0]/btn[0]",    # OK / Continue
            "wnd[1]/tbar[0]/btn[12]",   # Cancel / Close
        ]:
            try:
                session.findById(close_id).press()
                time.sleep(0.5)
                print("🪟  Checklist popup closed.")
                break
            except Exception:
                continue

    except Exception as e:
        print(f"⚠️  Status bar doubleClick failed: {e}")

    # ── STEP 3: Fallback — read directly from status bar text ─────────────────
    if not checklist_number:
        try:
            sbar_text = session.findById("wnd[0]/sbar").text.strip()
            if sbar_text:
                m = re.search(r'\b(\d{7,})\b', sbar_text)
                if m:
                    checklist_number = m.group(1)
                    print(f"✅ Checklist number read from status bar text: {checklist_number}")
        except Exception:
            pass

    # ── STEP 4: Read all other checklist fields from the form ─────────────────
    print("\n" + "=" * 60)
    print("  CHECKLIST DATA CAPTURED FROM SAP")
    print("=" * 60)

    data = {}

    # --- Header fields (text fields in the main form) ---
    field_map = {
        "Checklist No"       : "wnd[0]/usr/txtZFIBPT11001-CHK_NO",
        "Plant"              : "wnd[0]/usr/ctxtZFIBPT11001-WERKS",
        "Company Code"       : "wnd[0]/usr/ctxtZFIBPT11001-BUKRS",
        "Vendor"             : "wnd[0]/usr/ctxtZFIBPT11001-LIFNR",
        "Adani GSTIN"        : "wnd[0]/usr/txtZFIBPT11001-GSTIN",
        "GST Partner"        : "wnd[0]/usr/txtZFIBPT11001-GST_PART",
        "GST Partner Reg No" : "wnd[0]/usr/txtZFIBPT11001-GST_REG",
        "Doc Category"       : "wnd[0]/usr/ctxtZFIBPT11001-CHK_DOC_TYP",
        "Value of Checklist" : "wnd[0]/usr/txtZFIBPT11001-CHK_VALUE",
        "Debit Note Date"    : "wnd[0]/usr/ctxtZFIBPT11001-DEB_NOT_DAT",
        "Debit Note Ref"     : "wnd[0]/usr/txtZFIBPT11001-DEB_NOT_REF",
        "Doc Sub Category"   : "wnd[0]/usr/ctxtZFIBPT11001-DOC_SUB_CAT",
    }

    for label, ctrl_id in field_map.items():
        try:
            val = session.findById(ctrl_id).text.strip()
            data[label] = val if val else "—"
        except Exception:
            data[label] = "—"

    # --- Status bar: supplement data dict with already-captured checklist number ---
    try:
        status_msg = session.findById("wnd[0]/sbar").text.strip()
        if status_msg:
            data["Status Bar"] = status_msg
            # If the VBS preamble didn't get a number, try once more here
            if not checklist_number:
                match = re.search(r'\b(\d{7,})\b', status_msg)
                if match:
                    checklist_number = match.group(1)
    except Exception:
        pass

    if checklist_number:
        data["Checklist No"] = checklist_number

    # --- Print captured fields ---
    for label, val in data.items():
        print(f"  {label:<22}: {val}")

    if checklist_number:
        print(f"\n  ✅ Checklist Number : {checklist_number}")
    else:
        print("\n  ⚠️  Checklist number could not be parsed from status bar.")

    print("=" * 60 + "\n")

    # ------------------------------------------------------------------
    # Close all open SAP windows — handle logoff confirmation popup
    # ------------------------------------------------------------------
    print("🔒 Closing all SAP windows...")
    for i in range(9, -1, -1):
        try:
            session.findById(f"wnd[{i}]").close()
            time.sleep(0.5)
            print(f"   Closed wnd[{i}]")
        except Exception:
            pass

        # After each close attempt, check if a logoff/confirmation popup appeared
        # and click Yes (btn[8] = Yes, btn[0] = OK — try both)
        for popup in range(5, -1, -1):
            try:
                pop = session.findById(f"wnd[{popup}]")
                # Try Yes button first (btn[8]), then OK (btn[0])
                confirmed = False
                for btn_id in [
                    f"wnd[{popup}]/tbar[0]/btn[8]",   # Yes
                    f"wnd[{popup}]/tbar[0]/btn[0]",   # OK / Continue
                ]:
                    try:
                        session.findById(btn_id).press()
                        confirmed = True
                        print(f"   ✔ Clicked Yes/OK on popup wnd[{popup}]")
                        time.sleep(0.5)
                        break
                    except Exception:
                        continue
            except Exception:
                break   # popup window doesn't exist — move on

    # ------------------------------------------------------------------
    # Handle the "Log Off" popup — default focus is on "No", shift to "Yes"
    # Use Ctrl+Left arrow to move focus to Yes, then press Enter
    # ------------------------------------------------------------------
    time.sleep(1)
    try:
        # Send Ctrl+Left to move focus from No → Yes, then Enter to confirm
        session.findById("wnd[1]").sendVKey(0)   # make sure popup has focus
    except Exception:
        pass

    try:
        from pywinauto import Desktop
        from pywinauto.keyboard import send_keys

        logoff_dlg = Desktop(backend="uia").window(title_re=".*Log Off.*|.*Log off.*")
        logoff_dlg.wait("visible", timeout=5)
        logoff_dlg.set_focus()
        time.sleep(0.3)
        send_keys("{LEFT}")       # move focus from No → Yes
        time.sleep(0.3)
        send_keys("{ENTER}")      # confirm Yes
        print("   ✔ Navigated to 'Yes' and confirmed Log Off popup.")
        time.sleep(1.5)
    except Exception as e:
        print(f"   ⚠️  Could not handle Log Off popup via keyboard: {e}")

    # ------------------------------------------------------------------
    # Close the final remaining SAP window via its X (close) button
    # ------------------------------------------------------------------
    time.sleep(1)
    try:
        session.findById("wnd[0]").close()
        print("   ✔ Final SAP window closed via scripting.")
        time.sleep(0.5)
    except Exception:
        pass

    # Fallback: pywinauto close any remaining SAP windows
    try:
        from pywinauto import Desktop
        for win in Desktop(backend="uia").windows(title_re=".*SAP.*"):
            try:
                win.close()
                print(f"   ✔ Closed remaining SAP window: '{win.window_text()}'")
                time.sleep(0.5)
            except Exception:
                pass
    except Exception:
        pass

    print("✅ All SAP windows closed.")
    return checklist_number



    
# -------------------------------------------------
# PUBLIC ENTRY POINT (called from main.py)
# -------------------------------------------------
def run_sap_automation(
    company_code: str,
    vendor_code: str,
    amount: float,
    cost_center: str,
    plant: str,
    deb_not_date: str = None,
    print_pdf: bool = True,
):
    """
    Execute SAP Non-PO automation with values sourced from the email/SPV tables.

    Args:
        company_code  : BUKRS  – from vendor master data
        vendor_code   : LIFNR  – from vendor master data
        amount        : Absolute value of net_payable_receivable from SPV DSM table
                        (negative amounts are converted to positive before passing in)
        cost_center   : KOSTL  – from vendor master data
        plant         : WERKS  – from vendor master data
        deb_not_date  : DEB_NOT_DATE in SAP format DD.MM.YYYY (defaults to today)
        print_pdf     : When True (default), trigger the PDF print flow after saving.
                        Pass False to skip printing (e.g., during dry-run tests).
    """
    # --- Derive SAP-formatted date (DD.MM.YYYY) ---
    if not deb_not_date:
        deb_not_date = datetime.today().strftime("%d.%m.%Y")

    print(f"\n=== SAP Automation triggered ===")
    print(f"  Company Code  (BUKRS)       : {company_code}")
    print(f"  Vendor Code   (LIFNR)       : {vendor_code}")
    print(f"  Amount                      : {amount}")
    print(f"  Cost Center   (KOSTL)       : {cost_center}")
    print(f"  Plant         (WERKS)       : {plant}")
    print(f"  Debit Note Date             : {deb_not_date}")
    print(f"  Print PDF after save        : {print_pdf}")
    print(f"================================\n")

    # --- Static / env-driven config (loaded from config.py) ---
    saplogon_path = config.SAP_LOGON_PATH
    system_name   = config.SAP_SYSTEM
    sap_client    = config.SAP_CLIENT
    sap_user      = config.SAP_USERNAME
    sap_pass      = config.SAP_PASSWORD
    sap_lang      = config.SAP_LANGUAGE

    # --- Build params dict (only the fields that come from email data are overridden) ---
    params = {
        # Fields sourced from email / vendor master (passed in as arguments)
        "BUKRS":        company_code,
        "LIFNR":        vendor_code,
        "AMOUNT":       str(amount),          # already made positive by caller
        "COST_CENTER":  cost_center,
        "PLANT":        plant,
        "DEB_NOT_DATE": deb_not_date,

        # Static / env fields (loaded from config.py)
        "CHK_DOC_TYP":  config.SAP_CHK_DOC_TYP,
        "DIG_SIGN_INV": config.SAP_DIG_SIGN_INV,
        "GJAHR":        config.SAP_GJAHR,
        "DEB_NOT_REF":  config.SAP_DEB_NOT_REF,
        "DOC_SUB_CAT":  config.SAP_DOC_SUB_CAT,
        "DO_PREVIEW":   config.SAP_DO_PREVIEW,
    }

    # --- Orchestration ---
    launch_sap_logon(saplogon_path)
    app, conn, session = attach_to_scripting(system_name)

    try:
        session.findById("wnd[0]").maximize()
    except Exception:
        pass

    if not sap_user or not sap_pass:
        raise RuntimeError("SAP_USERNAME and SAP_PASSWORD environment variables must be set.")

    login(session, sap_client, sap_user, sap_pass, sap_lang)
    go_to_tcode(session, "ZDCC")

    select_non_po_and_create(session)
    fill_non_po_form(session, params)
    apply_subcategory_checkbox_table_and_save(session, params)

    pdf_dest = save_pdf(session)

    checklist_number   = None
    renamed_pdf_path   = pdf_dest   # will be updated after rename

    if pdf_dest:
        checklist_number = capture_checklist_data(session)

        # ── Rename: {original_stem}_{checklist_number}.pdf ──────────────────
        if checklist_number and pdf_dest:
            import pathlib as _pl
            _p       = _pl.Path(pdf_dest)
            _new_name = f"{_p.stem}_{checklist_number}{_p.suffix}"
            _new_path = _p.parent / _new_name
            try:
                if _new_path != _p:
                    os.rename(str(_p), str(_new_path))
                    print(f"📄 Checklist PDF renamed: '{_p.name}' → '{_new_name}'")
                renamed_pdf_path = str(_new_path)
            except Exception as _rename_err:
                print(f"⚠️  PDF rename failed ({_rename_err}) — "
                      f"keeping original: '{_p.name}'")
                renamed_pdf_path = pdf_dest
        elif pdf_dest:
            print(f"⚠️  checklist_number unavailable — "
                  f"PDF kept as: '{os.path.basename(pdf_dest)}'")

    print("✅ SAP automation flow completed.")
    return checklist_number, renamed_pdf_path


# -------------------------------------------------
# Standalone entry point (for isolated testing)
# -------------------------------------------------
def main():
    """
    Standalone test using hard-coded example values.
    In production this function is NOT called; run_sap_automation() is used instead.
    """
    run_sap_automation(
        company_code="6060",
        vendor_code="215944",
        amount=1,
        cost_center="6067OPC1",
        plant="6067",
        deb_not_date=None,   # defaults to today
        print_pdf=True,      # set to False to skip PDF printing during tests
    )

if __name__ == "__main__":
    main()