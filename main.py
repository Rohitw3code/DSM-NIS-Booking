"""
main.py
=======
DSM-NIS Automation Pipeline  –  Entry Point

Run modes
─────────
  python main.py --mode scheduled [--interval 300] [--lookback 2]
      Polls every N seconds.
      On each poll: scans the last `lookback` days of emails, finds any that
      (a) match DSM keywords or contain the required tables, AND
      (b) are not yet registered in SQLite (or are registered but not done).
      Then runs / resumes the pipeline for each such email.

  python main.py --mode immediate [--lookback 2]
      Single-shot: scans the last `lookback` days, processes all unfinished
      emails, then exits.

Pipeline steps (in order)
─────────────────────────
  1. read_email                – raw fetch done; body_html stored in DB
  2. email_extraction          – parse metadata_table + spv_dsm_table from HTML
  3. email_classification      – bucket every SPV field as +ve / -ve / zero
  4. email_validation          – confirm all required values are present
  5. read_vendor_master        – pull vendor row from SharePoint Excel
  6. nis_booking_<field>       – one step PER positive SPV field (Playwright)
  7. sap_automation_<field>    – one step PER negative SPV field (SAP GUI)
  8. update_excel_tracker      – append new row to SharePoint booking sheet

Each step is recorded in SQLite with status pending/running/done/failed/skipped.
If the process crashes mid-run, restarting will resume from the first non-done step.

python main.py --mode scheduled --interval 300 --lookback 2
python main.py --mode immediate --lookback 2
"""

import argparse
import json
import os
import shutil
import time
import traceback
from datetime import datetime
from pathlib import Path

import db_manager as db
from read_email import (
    fetch_emails_last_n_days,
    process_email,
    is_dsm_email,
    has_required_tables,
)
from read_vendor_master_data import read_vendor_data
from nis_booking             import book_nis
from add_new_row_data_nis    import add_incremental_week_row
from sap_automation          import run_sap_automation

# ─────────────────────────────────────────────────────────────────────────────
# Constants
# ─────────────────────────────────────────────────────────────────────────────
DEFAULT_POLL_INTERVAL = 300   # 5 minutes
DEFAULT_LOOKBACK_DAYS = 2

# Maximum number of attempts for NIS booking and SAP automation steps.
# On each failure the step is reset to 'pending' and retried automatically.
# Once this limit is reached the step is permanently marked failed and the
# pipeline stops for that run.
MAX_STEP_RETRIES = 3

# SPV amount fields that drive NIS / SAP routing
SPV_AMOUNT_FIELDS = [
    ("total_dsm_charges_payable", "Total DSM Charges Payable"),
    ("drawl_charges_payable",     "Drawl Charges Payable"),
    ("revenue_diff",              "Revenue Diff"),
    ("revenue_loss",              "Revenue Loss"),
]


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────
def safe_float(val) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def _format_invoice_date(raw: str) -> str:
    """
    Convert any of the common date formats found in DSM emails into the
    DD.MM.YYYY format required by the NIS invoice date field.

    Handles (case-insensitive month abbreviations):
      • DD-MM-YYYY       e.g. "02-12-2025"
      • D-MMM-YY         e.g. "2-Dec-25"
      • D-MMM-YYYY       e.g. "2-Dec-2025"
      • D MMM YYYY       e.g. "2 Dec 2025"
      • YYYY-MM-DD       e.g. "2025-12-02"
      • DD/MM/YYYY       e.g. "02/12/2025"
      • DD.MM.YYYY       already correct – returned as-is

    Falls back to the raw string if no format matches (so the pipeline
    never hard-crashes on an unexpected date format).
    """
    from datetime import datetime as _dt

    raw = str(raw or "").strip()
    if not raw:
        return raw

    _FMTS = [
        "%d.%m.%Y",   # already correct – check first
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%b-%y",   # 2-Dec-25
        "%d-%b-%Y",   # 2-Dec-2025
        "%d %b %Y",   # 2 Dec 2025
        "%d %b %y",   # 2 Dec 25
        "%d %B %Y",   # 2 December 2025
    ]
    for fmt in _FMTS:
        try:
            return _dt.strptime(raw, fmt).strftime("%d.%m.%Y")
        except ValueError:
            continue

    print(f"[WARN] _format_invoice_date: unrecognised format '{raw}' – using as-is")
    return raw


def normalize_vendor_data(raw: dict) -> dict:
    out = {}
    for k, v in raw.items():
        key = k.strip().upper()
        if key in ("VENDOR", "VENDOR CODE"):
            out["vendor_code"]  = str(v)
        elif key in ("COMPANY", "COMPANY CODE"):
            out["company_code"] = str(v)
        elif key == "COST CENTER":
            out["cost_center"]  = str(v)
        elif key == "PLANT":
            out["plant"]        = str(v)
        elif key == "BANK KEY":
            out["bank_key"]     = str(v)
        elif key == "GL":
            out["gl_account"]   = str(v)
        elif key == "PURPOSE":
            out["purpose"]      = str(v)
    return out


def _ts() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ─────────────────────────────────────────────────────────────────────────────
# Booking folder  –  save email artefacts after successful extraction
# ─────────────────────────────────────────────────────────────────────────────
BOOKING_ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "booking")


def _has_all_extracted_values(metadata_row: dict, spv_dsm_rows: list) -> bool:
    """
    Return True only when the extraction produced ALL the values we care about:
      • metadata: week_no, from_date, to_date, due_date
      • spv_dsm : at least one row with a non-empty spv_name
    """
    for field in ("week_no", "from_date", "to_date", "due_date"):
        if not str(metadata_row.get(field, "")).strip():
            return False
    if not spv_dsm_rows:
        return False
    if not any(str(r.get("spv_name", "")).strip() for r in spv_dsm_rows):
        return False
    return True


def _sanitize_folder_name(name: str) -> str:
    """Remove or replace characters that are unsafe in directory names."""
    # Replace path-unsafe chars with underscores, collapse whitespace
    import re as _re
    name = _re.sub(r'[\\/:*?"<>|]', "_", name)
    name = _re.sub(r"\s+", " ", name).strip()
    return name


def _save_email_to_booking_folder(
    metadata_row: dict,
    spv_dsm_rows: list,
    body_html: str,
    attachment_dir: str,
    subject: str = "",
) -> str:
    """
    Create  booking/<from_date> to <to_date>/  and save:
      • email.pdf          – the email body rendered to PDF (via Playwright)
      • all PDF files from attachment_dir (copied)

    Returns the absolute path to the created booking folder,
    or "" if the folder could not be created.
    """
    from_date = str(metadata_row.get("from_date", "")).strip()
    to_date   = str(metadata_row.get("to_date",   "")).strip()

    folder_name = _sanitize_folder_name(f"{from_date} to {to_date}")
    booking_dir = os.path.join(BOOKING_ROOT, folder_name)
    os.makedirs(booking_dir, exist_ok=True)

    # ── Save email body as PDF using Playwright ───────────────────────────
    email_pdf_path = os.path.join(booking_dir, "email.pdf")
    try:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            browser = pw.chromium.launch(headless=True)
            page = browser.new_page()
            page.set_content(body_html, wait_until="networkidle")
            page.pdf(path=email_pdf_path, format="A4", print_background=True)
            browser.close()
        print(f"  [BOOKING] Saved email as PDF → {email_pdf_path}")
    except Exception as exc:
        print(f"  [WARN] Could not convert email to PDF: {exc}")
        # Fallback: save as HTML so we don't lose the email content
        html_fallback = os.path.join(booking_dir, "email.html")
        try:
            with open(html_fallback, "w", encoding="utf-8") as f:
                f.write(body_html)
            print(f"  [BOOKING] Fallback – saved email body as HTML → {html_fallback}")
        except Exception as exc2:
            print(f"  [WARN] Could not save email HTML fallback: {exc2}")

    # ── Copy PDF attachments into the booking folder ──────────────────────
    if attachment_dir and os.path.isdir(attachment_dir):
        for fname in os.listdir(attachment_dir):
            if fname.lower().endswith(".pdf"):
                src = os.path.join(attachment_dir, fname)
                dst = os.path.join(booking_dir, fname)
                try:
                    shutil.copy2(src, dst)
                    print(f"  [BOOKING] Copied PDF → {dst}")
                except Exception as exc:
                    print(f"  [WARN] Could not copy {fname}: {exc}")

    print(f"  [BOOKING] Folder created: {booking_dir}")
    return booking_dir


# ─────────────────────────────────────────────────────────────────────────────
# Step 2 – Email extraction
# ─────────────────────────────────────────────────────────────────────────────
def _step_extraction(run_id: int, body_html: str) -> tuple:
    """
    Parse HTML tables.
    Returns (metadata_row: dict, spv_dsm_rows: list[dict]).
    Raises on failure.
    """
    from classified_html_table_parser import parse_html_tables

    if db.is_step_done(run_id, db.STEP_EMAIL_EXTRACTION):
        print("  [RESUME] email_extraction already done")
        run = db.get_run_by_id(run_id)
        metadata_row  = json.loads(run["metadata_json"]  or "{}")
        spv_dsm_rows  = json.loads(run["spv_dsm_json"]   or "[]")
        return metadata_row, spv_dsm_rows

    db.step_start(run_id, db.STEP_EMAIL_EXTRACTION)
    try:
        from classified_html_table_parser import parse_html_tables, print_raw_tables, print_raw_tables

        # ── Always print raw structure so mapping issues are immediately visible ──
        print_raw_tables(body_html)

        tables       = parse_html_tables(body_html)
        spv_dsm_rows = tables.get("spv_dsm_table",  [])
        metadata_row = (tables.get("metadata_table", []) or [{}])[0]

        # ── Show mapped result so you can compare against raw above ──────────
        print("[EXTRACTED] metadata_table (mapped):")
        print(f"  {metadata_row}")
        print("[EXTRACTED] spv_dsm_table (mapped):")
        for r in spv_dsm_rows:
            print(f"  {r}")
        print()

        if not spv_dsm_rows:
            raise ValueError("spv_dsm_table is empty after extraction")
        if not metadata_row:
            raise ValueError("metadata_table is empty after extraction")

        db.save_extracted_data(run_id, metadata_row, spv_dsm_rows)
        db.step_done(run_id, db.STEP_EMAIL_EXTRACTION,
                     detail=f"{len(spv_dsm_rows)} SPV row(s)  "
                            f"week={metadata_row.get('week_no','?')}")
        return metadata_row, spv_dsm_rows

    except Exception as exc:
        db.step_failed(run_id, db.STEP_EMAIL_EXTRACTION, str(exc))
        raise


# ─────────────────────────────────────────────────────────────────────────────
# Step 3 – Classification  (per SPV row)
# ─────────────────────────────────────────────────────────────────────────────
def _step_classification(run_id: int, spv_row: dict) -> tuple:
    """
    Classify each SPV amount field as positive or negative.
    Returns (positive_fields: dict, negative_fields: dict).
    """
    step = db.STEP_EMAIL_CLASSIFICATION

    if db.is_step_done(run_id, step):
        print("  [RESUME] email_classification already done")
        run = db.get_run_by_id(run_id)
        pos = json.loads(run["positive_fields_json"] or "{}")
        neg = json.loads(run["negative_fields_json"] or "{}")
        return pos, neg

    db.step_start(run_id, step)
    try:
        positive = {}
        negative = {}

        for field_key, label in SPV_AMOUNT_FIELDS:
            amount = safe_float(spv_row.get(field_key))
            if amount > 0:
                positive[field_key] = amount
                print(f"    [+] {label} = {amount}  → NIS booking")
            elif amount < 0:
                negative[field_key] = abs(amount)
                print(f"    [-] {label} = {amount}  → SAP automation (abs={abs(amount)})")
            else:
                print(f"    [0] {label} = 0 / blank  → skipped")

        db.save_classification(run_id, positive, negative)

        # Seed dynamic step rows now that we know which fields exist
        for field_key in positive:
            db.ensure_dynamic_step(run_id, db.step_nis(field_key))
        for field_key in negative:
            db.ensure_dynamic_step(run_id, db.step_sap(field_key))

        db.step_done(run_id, step,
                     detail=f"positive={list(positive.keys())}  "
                            f"negative={list(negative.keys())}")
        return positive, negative

    except Exception as exc:
        db.step_failed(run_id, step, str(exc))
        raise


# ─────────────────────────────────────────────────────────────────────────────
# Step 4 – Validation
# ─────────────────────────────────────────────────────────────────────────────
def _step_validation(run_id: int, metadata_row: dict, spv_row: dict,
                     positive: dict, negative: dict):
    step = db.STEP_EMAIL_VALIDATION

    if db.is_step_done(run_id, step):
        print("  [RESUME] email_validation already done")
        return

    db.step_start(run_id, step)
    errors = []

    # Required metadata fields
    for field in ("week_no", "from_date", "to_date", "due_date"):
        if not str(metadata_row.get(field, "")).strip():
            errors.append(f"metadata.{field} is missing")

    # SPV must have an spv_name
    if not str(spv_row.get("spv_name", "")).strip():
        errors.append("spv_dsm.spv_name is missing")

    # At least one actionable amount
    if not positive and not negative:
        errors.append("All SPV amounts are zero – nothing to process")

    if errors:
        msg = "; ".join(errors)
        db.step_failed(run_id, step, msg)
        raise ValueError(f"Validation failed: {msg}")

    db.step_done(run_id, step,
                 detail=f"week={metadata_row.get('week_no')}  "
                        f"spv={spv_row.get('spv_name')}")


# ─────────────────────────────────────────────────────────────────────────────
# Step 5 – Vendor master
# ─────────────────────────────────────────────────────────────────────────────
def _step_vendor_master(run_id: int, spv_name: str) -> dict:
    step = db.STEP_READ_VENDOR

    if db.is_step_done(run_id, step):
        print("  [RESUME] read_vendor_master already done")
        run = db.get_run_by_id(run_id)
        return json.loads(run["vendor_data_json"] or "{}")

    db.step_start(run_id, step, detail=f"sheet={spv_name}")
    try:
        raw_vendor  = read_vendor_data(spv_name)
        vendor_data = normalize_vendor_data(raw_vendor)
        db.save_vendor_data(run_id, vendor_data)
        db.step_done(run_id, step, detail=json.dumps(vendor_data))
        return vendor_data
    except Exception as exc:
        db.step_failed(run_id, step, str(exc))
        raise


# ─────────────────────────────────────────────────────────────────────────────
# Step 6 – NIS booking  (one call per positive field, up to MAX_STEP_RETRIES)
# ─────────────────────────────────────────────────────────────────────────────
def _step_nis_booking(
    run_id:       int,
    field_key:    str,
    amount:       float,
    spv_row:      dict,
    negative:     dict,
    vendor_data:  dict,
    invoice_date: str,
    attachment_dir: str,
    metadata_row: dict = None,
):
    step  = db.step_nis(field_key)
    label = next(lbl for k, lbl in SPV_AMOUNT_FIELDS if k == field_key)

    if db.is_step_done(run_id, step):
        print(f"  [RESUME] {step} already done")
        return

    # ── Guard: already exhausted retries in a previous session ───────────────
    retry_count = db.get_step_retry_count(run_id, step)
    if retry_count >= MAX_STEP_RETRIES and not db.is_step_done(run_id, step):
        msg = (f"NIS booking [{label}] exceeded maximum {MAX_STEP_RETRIES} "
               f"retries – marking permanently failed")
        print(f"  ❌ {msg}")
        db.step_failed(run_id, step, msg)
        raise RuntimeError(msg)

    # Find PDF in the per-message attachment directory
    att_path  = Path(attachment_dir)
    pdf_files = list(att_path.glob("*.pdf"))
    if not pdf_files:
        att_path  = Path("./attachments")
        pdf_files = list(att_path.glob("*.pdf"))
    if not pdf_files:
        err = f"No PDF found for NIS booking ({field_key})"
        db.step_failed(run_id, step, err)
        raise FileNotFoundError(err)

    pdf_path = str(pdf_files[0])

    # Zero-out negative fields and other positive fields in the row copy
    nis_row = dict(spv_row)
    for fk in negative:
        nis_row[fk] = 0
    for fk, _ in SPV_AMOUNT_FIELDS:
        if fk != field_key:
            nis_row[fk] = 0

    company_code = vendor_data.get("company_code", "")
    vendor_code  = vendor_data.get("vendor_code",  "")
    bank_key     = vendor_data.get("bank_key",     "")
    cost_center  = vendor_data.get("cost_center",  "")
    plant        = vendor_data.get("plant",        "")

    # Purpose of expenditure: "from_date TO to_date DSM charges"
    _meta        = metadata_row or {}
    from_date    = str(_meta.get("from_date", "")).strip()
    to_date      = str(_meta.get("to_date",   "")).strip()
    if from_date and to_date:
        purpose = f"{from_date} TO {to_date} DSM charges"
    else:
        purpose  = vendor_data.get("purpose", "NIS booking for DSM")

    # ── Retry loop ────────────────────────────────────────────────────────────
    while True:
        current_attempt = db.get_step_retry_count(run_id, step) + 1
        db.step_start(
            run_id, step,
            detail=f"attempt={current_attempt}/{MAX_STEP_RETRIES}  "
                   f"amount={amount}  pdf={Path(pdf_path).name}",
        )
        try:
            print(f"\n  ▶ NIS booking  [{label}]  amount={amount}"
                  f"  (attempt {current_attempt}/{MAX_STEP_RETRIES})")

            nis_number = book_nis(
                pdf_path     = pdf_path,
                company_code = company_code,
                vendor_code  = vendor_code,
                bank_key     = bank_key,
                invoice_date = invoice_date,
                cost_center  = cost_center,
                plant        = plant,
                purpose      = purpose,
                svp_dsm_row  = nis_row,
            )

            # ── Persist the NIS booking number captured from the portal ───────
            if nis_number:
                db.save_nis_checklist_id(run_id, nis_number)
                print(f"  [DB] NIS booking number saved: {nis_number}")
            else:
                print("  [WARN] NIS booking number was not captured from the portal.")

            # ── Persist the email attachment PDF blob ─────────────────────────
            db.save_downloaded_pdfs(run_id, pdf_files, checklist_pdf_name="")

            db.step_done(
                run_id, step,
                detail=f"amount={amount}  attempts={current_attempt}"
                       f"  nis_number={nis_number or 'not_captured'}",
            )
            print(f"  ✅ NIS done  [{label}]  NIS booking number: {nis_number or '(not captured)'}")
            return  # ← success: exit retry loop

        except Exception as exc:
            err_str = str(exc)
            new_count = db.increment_step_retry(run_id, step, err_str)
            print(f"  ⚠️  NIS booking attempt {current_attempt} failed: {err_str[:120]}")

            if new_count >= MAX_STEP_RETRIES:
                # All retries exhausted – mark permanently failed and propagate
                final_msg = (
                    f"NIS booking [{label}] failed after {MAX_STEP_RETRIES} "
                    f"attempt(s). Last error: {err_str[:200]}"
                )
                db.step_failed(run_id, step, final_msg)
                print(f"  ❌ {final_msg}")
                raise RuntimeError(final_msg) from exc

            # Still have retries left – brief pause then loop
            wait_sec = 5 * current_attempt   # 5s, 10s between attempts
            print(f"  🔁 Retrying in {wait_sec}s "
                  f"({MAX_STEP_RETRIES - new_count} attempt(s) remaining) …")
            time.sleep(wait_sec)


# ─────────────────────────────────────────────────────────────────────────────
# Step 7 – SAP automation  (one call per negative field, up to MAX_STEP_RETRIES)
# ─────────────────────────────────────────────────────────────────────────────
def _step_sap_automation(
    run_id:      int,
    field_key:   str,
    abs_amount:  float,
    vendor_data: dict,
    booking_dir: str = "",
):
    step  = db.step_sap(field_key)
    label = next(lbl for k, lbl in SPV_AMOUNT_FIELDS if k == field_key)

    if db.is_step_done(run_id, step):
        print(f"  [RESUME] {step} already done")
        return

    # ── Guard: already exhausted retries in a previous session ───────────────
    retry_count = db.get_step_retry_count(run_id, step)
    if retry_count >= MAX_STEP_RETRIES and not db.is_step_done(run_id, step):
        msg = (f"SAP automation [{label}] exceeded maximum {MAX_STEP_RETRIES} "
               f"retries – marking permanently failed")
        print(f"  ❌ {msg}")
        db.step_failed(run_id, step, msg)
        raise RuntimeError(msg)

    company_code = vendor_data.get("company_code", "")
    vendor_code  = vendor_data.get("vendor_code",  "")
    cost_center  = vendor_data.get("cost_center",  "")
    plant        = vendor_data.get("plant",        "")
    deb_not_date = datetime.today().strftime("%d.%m.%Y")

    # ── Retry loop ────────────────────────────────────────────────────────────
    while True:
        current_attempt = db.get_step_retry_count(run_id, step) + 1
        db.step_start(
            run_id, step,
            detail=f"attempt={current_attempt}/{MAX_STEP_RETRIES}  "
                   f"amount={abs_amount}  date={deb_not_date}",
        )
        try:
            print(f"\n  ▶ SAP automation  [{label}]  amount={abs_amount}"
                  f"  (attempt {current_attempt}/{MAX_STEP_RETRIES})")
            sap_result = run_sap_automation(
                company_code = company_code,
                vendor_code  = vendor_code,
                amount       = abs_amount,
                cost_center  = cost_center,
                plant        = plant,
                deb_not_date = deb_not_date,
            )

            # ── Persist SAP checklist value + PDF name immediately ────────────
            # run_sap_automation returns (checklist_number, renamed_pdf_path)
            sap_checklist_number = ""
            sap_pdf_name         = ""
            if isinstance(sap_result, tuple) and len(sap_result) >= 2:
                sap_checklist_number = str(sap_result[0] or "")
                sap_pdf_name         = str(Path(sap_result[1]).name) if sap_result[1] else ""
            elif isinstance(sap_result, str):
                sap_checklist_number = sap_result or ""

            if sap_checklist_number:
                db.save_checklist_value(run_id, sap_checklist_number)
                print(f"  [DB] SAP checklist number saved: {sap_checklist_number}")
            if sap_pdf_name:
                # Save only the PDF filename (not the blob) as requested
                db._update_run(run_id, {"checklist_pdf_name": sap_pdf_name})
                print(f"  [DB] SAP checklist PDF name saved: {sap_pdf_name}")

            # ── Move SAP checklist PDF into the booking folder ────────────────
            sap_pdf_full_path = ""
            if isinstance(sap_result, tuple) and len(sap_result) >= 2:
                sap_pdf_full_path = str(sap_result[1] or "")

            if booking_dir and sap_pdf_full_path and os.path.isfile(sap_pdf_full_path):
                dst = os.path.join(booking_dir, os.path.basename(sap_pdf_full_path))
                try:
                    shutil.move(sap_pdf_full_path, dst)
                    print(f"  [BOOKING] Moved SAP checklist PDF → {dst}")
                except Exception as mv_exc:
                    print(f"  [WARN] Could not move SAP checklist PDF to booking folder: {mv_exc}")

            db.step_done(
                run_id, step,
                detail=(
                    f"amount={abs_amount}  attempts={current_attempt}"
                    f"  checklist={sap_checklist_number or 'n/a'}"
                    f"  pdf={sap_pdf_name or 'n/a'}"
                ),
            )
            print(f"  ✅ SAP done  [{label}]")
            return  # ← success: exit retry loop

        except Exception as exc:
            err_str = str(exc)
            new_count = db.increment_step_retry(run_id, step, err_str)
            print(f"  ⚠️  SAP automation attempt {current_attempt} failed: {err_str[:120]}")

            if new_count >= MAX_STEP_RETRIES:
                final_msg = (
                    f"SAP automation [{label}] failed after {MAX_STEP_RETRIES} "
                    f"attempt(s). Last error: {err_str[:200]}"
                )
                db.step_failed(run_id, step, final_msg)
                print(f"  ❌ {final_msg}")
                raise RuntimeError(final_msg) from exc

            wait_sec = 5 * current_attempt   # 5s, 10s between attempts
            print(f"  🔁 Retrying in {wait_sec}s "
                  f"({MAX_STEP_RETRIES - new_count} attempt(s) remaining) …")
            time.sleep(wait_sec)


# ─────────────────────────────────────────────────────────────────────────────
# Step 8 – Update Excel tracker
# ─────────────────────────────────────────────────────────────────────────────
def _step_update_tracker(
    run_id:      int,
    metadata_row: dict,
    spv_row:     dict,
    spv_name:    str,
):
    step = db.STEP_UPDATE_TRACKER

    if db.is_step_done(run_id, step):
        print("  [RESUME] update_excel_tracker already done")
        return

    db.step_start(run_id, step, detail=f"sheet={spv_name}")
    try:
        # Pull checklist_id saved during NIS booking step (may be empty on first run)
        run_row    = db.get_run_by_id(run_id)
        checklist_id = run_row.get("nis_checklist_id") or ""

        add_incremental_week_row(
            metadata_row = metadata_row,
            spv_dsm_row  = spv_row,
            sheet_name   = spv_name,
            checklist_id = checklist_id,
        )
        db.step_done(run_id, step,
                     detail=f"week={metadata_row.get('week_no')}  spv={spv_name}"
                            f"  checklist={checklist_id or 'n/a'}")
    except Exception as exc:
        db.step_failed(run_id, step, str(exc))
        raise


# ─────────────────────────────────────────────────────────────────────────────
# Master pipeline runner  (resume-aware)
# ─────────────────────────────────────────────────────────────────────────────
def run_pipeline(run_id: int, email_result: dict = None):
    """
    Execute or resume the complete pipeline for run_id.

    email_result is only needed on the first run (before extraction is done).
    On resume it is loaded from the DB.
    """
    db.mark_run_started(run_id)
    run = db.get_run_by_id(run_id)

    print(f"\n{'─'*65}")
    print(f"  PIPELINE  run_id={run_id}  →  {run['subject']}")
    print(f"{'─'*65}")

    try:
        # ── STEP 1: read_email ────────────────────────────────────────────
        # The body_html was stored in the DB during registration; just mark done.
        if not db.is_step_done(run_id, db.STEP_READ_EMAIL):
            db.step_start(run_id, db.STEP_READ_EMAIL)
            db.step_done(run_id, db.STEP_READ_EMAIL,
                         detail=f"msg_id={run['message_id'][:30]}")
        else:
            print("  [RESUME] read_email already done")

        # Retrieve body_html (from DB or from the freshly-fetched email_result)
        body_html = run.get("body_html") or ""
        if not body_html and email_result:
            body_html = email_result.get("body_html", "")

        # attachment_dir: from email_result on first run; fall back to global
        attachment_dir = ""
        if email_result:
            attachment_dir = email_result.get("attachment_dir", "")

        # ── STEP 2: email_extraction ──────────────────────────────────────
        metadata_row, spv_dsm_rows = _step_extraction(run_id, body_html)

        # ── Save email to booking folder if ALL extracted values are present ──
        booking_folder = ""
        if _has_all_extracted_values(metadata_row, spv_dsm_rows):
            booking_folder = _save_email_to_booking_folder(
                metadata_row   = metadata_row,
                spv_dsm_rows   = spv_dsm_rows,
                body_html      = body_html,
                attachment_dir = attachment_dir,
                subject        = run.get("subject", ""),
            )
            # Update attachment_dir so later steps (NIS/SAP) use the booking folder
            if booking_folder:
                attachment_dir = booking_folder
        else:
            print("  [BOOKING] Skipped folder creation – not all extracted values present")

        # We process SPV rows one by one; most emails have exactly one
        # but the loop handles multiple SPV rows gracefully
        for spv_row in spv_dsm_rows:
            spv_name = spv_row.get("spv_name", "").strip()
            if not spv_name:
                print("  [WARN] SPV row has no spv_name – skipping row")
                continue

            db.save_spv_name(run_id, spv_name)

            # ── STEP 3: email_classification ─────────────────────────────
            print(f"\n  ── Classifying SPV: {spv_name} ──")
            positive, negative = _step_classification(run_id, spv_row)

            # ── STEP 4: email_validation ──────────────────────────────────
            _step_validation(run_id, metadata_row, spv_row, positive, negative)

            # ── STEP 5: read_vendor_master ────────────────────────────────
            vendor_data = _step_vendor_master(run_id, spv_name)

            invoice_date = _format_invoice_date(
                metadata_row.get("dsm_statement_published_date") or
                metadata_row.get("from_date") or
                "10.02.2020"
            )

            # ── STEP 6: SAP automation (per negative field) ───────────────
            if negative:
                print(f"\n  ── SAP automation for {list(negative.keys())} ──")
                for field_key, abs_amount in negative.items():
                    _step_sap_automation(run_id, field_key, abs_amount,
                                        vendor_data, booking_dir=booking_folder)
            else:
                # Ensure any pre-seeded sap steps are marked skipped
                for field_key, _ in SPV_AMOUNT_FIELDS:
                    s = db.step_sap(field_key)
                    step_row = db.get_step(run_id, s)
                    if step_row and not db.is_step_done(run_id, s):
                        db.step_skip(run_id, s, "no negative fields")
                print("  [INFO] No negative fields – SAP automation skipped")

            # ── STEP 7: NIS booking (per positive field) ──────────────────
            if positive:
                print(f"\n  ── NIS booking for {list(positive.keys())} ──")
                for field_key, amount in positive.items():
                    _step_nis_booking(
                        run_id, field_key, amount,
                        spv_row, negative, vendor_data,
                        invoice_date, attachment_dir,
                        metadata_row=metadata_row,
                    )
            else:
                for field_key, _ in SPV_AMOUNT_FIELDS:
                    s = db.step_nis(field_key)
                    step_row = db.get_step(run_id, s)
                    if step_row and not db.is_step_done(run_id, s):
                        db.step_skip(run_id, s, "no positive fields")
                print("  [INFO] No positive fields – NIS booking skipped")

            # ── STEP 8: update Excel tracker ──────────────────────────────
            _step_update_tracker(run_id, metadata_row, spv_row, spv_name)

        # ── All steps done ────────────────────────────────────────────────
        db.mark_run_done(run_id)
        db.print_run_summary(run_id)
        print(f"✅  Pipeline COMPLETE  run_id={run_id}\n")

    except Exception as exc:
        db.mark_run_failed(run_id, str(exc))
        db.print_run_summary(run_id)
        print(f"❌  Pipeline FAILED  run_id={run_id}\n{traceback.format_exc()}")
        raise


# ─────────────────────────────────────────────────────────────────────────────
# Scan + dispatch  (called by both modes)
# ─────────────────────────────────────────────────────────────────────────────
def scan_and_dispatch(lookback_days: int = DEFAULT_LOOKBACK_DAYS):
    """
    1. Fetch all emails from the last `lookback_days` days.
    2. For each email:
       a. Skip if subject is clearly unrelated AND tables are absent.
       b. If already registered + done → skip.
       c. If already registered + not done → resume.
       d. If not registered → register + run.
    """
    print(f"\n[{_ts()}] Scanning last {lookback_days} day(s) of email …")

    try:
        messages, headers, base_url = fetch_emails_last_n_days(days=lookback_days)
    except Exception as exc:
        print(f"[ERROR] Could not fetch emails: {exc}")
        return

    if not messages:
        print("[INFO] No emails found in the scan window.")
        return

    print(f"[INFO] {len(messages)} email(s) to evaluate.\n")
    processed = skipped = resumed = 0

    for msg in messages:
        message_id = msg["id"]
        subject    = msg.get("subject", "")

        # ── Pre-filter by subject keywords (cheap check before full parse) ──
        if not is_dsm_email(subject):
            print(f"  [SKIP] Subject does not match DSM keywords: {subject[:60]}")
            skipped += 1
            continue

        # ── Already done → skip completely ───────────────────────────────────
        if db.is_email_registered(message_id):
            existing = db.get_run_by_message_id(message_id)
            if existing["status"] == db.STATUS_DONE:
                print(f"  [DEDUP] Already done  run_id={existing['id']}  "
                      f"{subject[:50]}")
                skipped += 1
                continue
            # Registered but not done → resume
            print(f"  [RESUME] run_id={existing['id']}  status={existing['status']}  "
                  f"{subject[:50]}")
            try:
                run_pipeline(existing["id"])
                resumed += 1
            except Exception:
                pass
            continue

        # ── New email: full process → parse tables → decide relevance ─────────
        print(f"  [NEW] Processing: {subject[:60]}")
        try:
            email_result = process_email(msg, headers, base_url)
        except Exception as exc:
            print(f"  [ERROR] Could not process email {message_id[:20]}: {exc}")
            skipped += 1
            continue

        # Secondary relevance check: does it have the required tables?
        if not has_required_tables(email_result["parsed_tables"]):
            print(f"  [SKIP] No DSM tables found in body: {subject[:60]}")
            skipped += 1
            continue

        # Register + run
        run_id = db.register_email(
            message_id  = message_id,
            subject     = subject,
            received_at = email_result["received_at"],
            body_html   = email_result["body_html"],
        )
        try:
            run_pipeline(run_id, email_result=email_result)
            processed += 1
        except Exception:
            pass   # already logged inside run_pipeline

    print(f"\n[SCAN COMPLETE]  processed={processed}  resumed={resumed}  "
          f"skipped={skipped}\n")


# ─────────────────────────────────────────────────────────────────────────────
# Run modes
# ─────────────────────────────────────────────────────────────────────────────
def run_scheduled(interval_sec: int, lookback_days: int):
    print(f"\n{'═'*65}")
    print(f"  SCHEDULED MODE  |  interval={interval_sec}s  lookback={lookback_days}d")
    print(f"{'═'*65}\n")
    db.init_db()

    while True:
        # Resume any stuck runs from previous sessions first
        pending = db.get_pending_runs()
        if pending:
            print(f"[RESUME] {len(pending)} pending run(s) from previous session(s).")
            for r in pending:
                print(f"  → run_id={r['id']}  {r['subject'][:50]}")
                try:
                    run_pipeline(r["id"])
                except Exception:
                    pass

        scan_and_dispatch(lookback_days)

        print(f"[{_ts()}] Next poll in {interval_sec}s …\n")
        time.sleep(interval_sec)


def run_immediate(lookback_days: int):
    print(f"\n{'═'*65}")
    print(f"  IMMEDIATE MODE  |  lookback={lookback_days}d")
    print(f"{'═'*65}\n")
    db.init_db()

    # Resume stuck runs first
    pending = db.get_pending_runs()
    if pending:
        print(f"[RESUME] {len(pending)} pending run(s) found.")
        for r in pending:
            try:
                run_pipeline(r["id"])
            except Exception:
                pass

    scan_and_dispatch(lookback_days)
    print("\n=== IMMEDIATE MODE complete ===\n")


# ─────────────────────────────────────────────────────────────────────────────
# CLI
# ─────────────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="DSM-NIS Automation Pipeline",
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        "--mode",
        choices=["scheduled", "immediate"],
        default="scheduled",
        help=(
            "scheduled  – poll continuously on a fixed interval (default)\n"
            "immediate  – scan once and exit"
        ),
    )
    parser.add_argument(
        "--interval",
        type=int,
        default=DEFAULT_POLL_INTERVAL,
        metavar="SECONDS",
        help=f"Poll interval for scheduled mode (default: {DEFAULT_POLL_INTERVAL})",
    )
    parser.add_argument(
        "--lookback",
        type=int,
        default=DEFAULT_LOOKBACK_DAYS,
        metavar="DAYS",
        help=f"How many days back to scan for emails (default: {DEFAULT_LOOKBACK_DAYS})",
    )
    args = parser.parse_args()

    if args.mode == "immediate":
        run_immediate(args.lookback)
    else:
        run_scheduled(args.interval, args.lookback)


if __name__ == "__main__":
    main()