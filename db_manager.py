"""
db_manager.py
=============
SQLite persistence layer for the DSM-NIS automation pipeline.

Tables
------
  email_runs   – one row per unique email (dedup key = message_id)
  run_steps    – one row per (run_id, step_name) pair

Step names (in execution order)
--------------------------------
  1.  read_email                  – fetch raw email from Outlook Graph API
  2.  email_extraction            – parse HTML tables from body HTML
  3.  email_classification        – bucket each SPV amount as +ve / -ve / zero
  4.  email_validation            – confirm required fields exist and are usable
  5.  read_vendor_master          – pull vendor data from SharePoint Excel
  6.  nis_booking_<field>         – one step PER positive SPV field
  7.  sap_automation_<field>      – one step PER negative SPV field
  8.  update_excel_tracker        – write new row to SharePoint booking sheet

Status values: pending | running | done | failed | skipped
"""

import sqlite3
import json
from datetime import datetime
from contextlib import contextmanager
from typing import Optional
import config

DB_PATH = config.DB_PATH

# ── Fixed step-name constants ─────────────────────────────────────────────────
STEP_READ_EMAIL           = "read_email"
STEP_EMAIL_EXTRACTION     = "email_extraction"
STEP_EMAIL_CLASSIFICATION = "email_classification"
STEP_EMAIL_VALIDATION     = "email_validation"
STEP_READ_VENDOR          = "read_vendor_master"
STEP_UPDATE_TRACKER       = "update_excel_tracker"

# ── Dynamic step-name helpers ─────────────────────────────────────────────────
def step_nis(field: str) -> str:
    return f"nis_booking_{field}"

def step_sap(field: str) -> str:
    return f"sap_automation_{field}"

# ── Status values ─────────────────────────────────────────────────────────────
STATUS_PENDING = "pending"
STATUS_RUNNING = "running"
STATUS_DONE    = "done"
STATUS_FAILED  = "failed"
STATUS_SKIPPED = "skipped"

# ── Fixed steps seeded on every new run ──────────────────────────────────────
_FIXED_STEPS = [
    STEP_READ_EMAIL,
    STEP_EMAIL_EXTRACTION,
    STEP_EMAIL_CLASSIFICATION,
    STEP_EMAIL_VALIDATION,
    STEP_READ_VENDOR,
    STEP_UPDATE_TRACKER,
]


# ─────────────────────────────────────────────────────────────────────────────
# Connection helper
# ─────────────────────────────────────────────────────────────────────────────
@contextmanager
def _conn():
    con = sqlite3.connect(DB_PATH, timeout=15)
    con.row_factory = sqlite3.Row
    con.execute("PRAGMA journal_mode=WAL")
    con.execute("PRAGMA foreign_keys=ON")
    try:
        yield con
        con.commit()
    except Exception:
        con.rollback()
        raise
    finally:
        con.close()


# ─────────────────────────────────────────────────────────────────────────────
# Migration helper
# ─────────────────────────────────────────────────────────────────────────────
def _migrate_add_column(col_name: str, col_type: str, table: str = "email_runs"):
    """
    Add a column to an existing table if it does not already exist.
    Safe to call repeatedly (idempotent).
    """
    with _conn() as con:
        existing = [
            row[1]
            for row in con.execute(f"PRAGMA table_info({table})").fetchall()
        ]
        if col_name not in existing:
            con.execute(f"ALTER TABLE {table} ADD COLUMN {col_name} {col_type}")
            print(f"[DB] Migration: added column '{col_name}' to {table}")


# ─────────────────────────────────────────────────────────────────────────────
# Schema bootstrap
# ─────────────────────────────────────────────────────────────────────────────
def init_db():
    with _conn() as con:
        con.executescript("""
            CREATE TABLE IF NOT EXISTS email_runs (
                id                   INTEGER PRIMARY KEY AUTOINCREMENT,
                message_id           TEXT    NOT NULL UNIQUE,
                subject              TEXT,
                received_at          TEXT,
                spv_name             TEXT,
                status               TEXT    NOT NULL DEFAULT 'pending',

                -- raw body stored so extraction can be retried on resume
                body_html            TEXT,

                -- parsed data stored after extraction step
                metadata_json        TEXT,
                -- spv_dsm_json holds a JSON *list* (one dict per SPV row)
                spv_dsm_json         TEXT,

                -- normalised vendor data stored after read_vendor_master step
                vendor_data_json     TEXT,

                -- classification output (JSON objects: {field: abs_amount})
                positive_fields_json TEXT,
                negative_fields_json TEXT,

                -- attachment tracking
                downloaded_pdf_names TEXT,   -- JSON list of every PDF filename downloaded
                checklist_pdf_name   TEXT,   -- filename of the DSM checklist PDF specifically
                nis_checklist_id     TEXT,   -- NIS checklist/booking number captured from portal

                -- PDF binary storage
                downloaded_pdf_blob  BLOB,   -- binary content of the DSM email PDF attachment
                checklist_pdf_blob   BLOB,   -- binary content of the NIS checklist PDF

                -- checklist value captured from portal confirmation screen
                checklist_value      TEXT,   -- human-readable checklist/booking value from portal

                -- housekeeping
                created_at           TEXT NOT NULL,
                started_at           TEXT,
                finished_at          TEXT,
                error_message        TEXT
            );

            CREATE TABLE IF NOT EXISTS run_steps (
                id          INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id      INTEGER NOT NULL
                                REFERENCES email_runs(id) ON DELETE CASCADE,
                step_name   TEXT    NOT NULL,
                status      TEXT    NOT NULL DEFAULT 'pending',
                detail      TEXT,
                started_at  TEXT,
                finished_at TEXT,
                error       TEXT,
                retry_count INTEGER NOT NULL DEFAULT 0,
                UNIQUE(run_id, step_name)
            );
        """)
    print(f"[DB] Initialised → {DB_PATH}")

    # ── Idempotent column migrations (safe to run on existing DB) ────────────
    _migrate_add_column("downloaded_pdf_names", "TEXT")
    _migrate_add_column("checklist_pdf_name",   "TEXT")
    _migrate_add_column("nis_checklist_id",      "TEXT")
    _migrate_add_column("downloaded_pdf_blob",  "BLOB")
    _migrate_add_column("checklist_pdf_blob",   "BLOB")
    _migrate_add_column("checklist_value",      "TEXT")
    _migrate_add_column("retry_count", "INTEGER NOT NULL DEFAULT 0", table="run_steps")


# ─────────────────────────────────────────────────────────────────────────────
# email_runs helpers
# ─────────────────────────────────────────────────────────────────────────────
def is_email_registered(message_id: str) -> bool:
    with _conn() as con:
        row = con.execute(
            "SELECT 1 FROM email_runs WHERE message_id=?", (message_id,)
        ).fetchone()
        return row is not None


def register_email(
    message_id: str,
    subject:    str,
    received_at: str,
    body_html:  str = "",
) -> int:
    """
    Insert a brand-new email_run row + seed all fixed step rows as 'pending'.
    Dynamic NIS/SAP steps are added later (after classification).
    Returns run_id.
    """
    now = _now()
    with _conn() as con:
        cur = con.execute(
            """
            INSERT INTO email_runs
                (message_id, subject, received_at, body_html, status, created_at)
            VALUES (?,?,?,?,?,?)
            """,
            (message_id, subject, received_at, body_html, STATUS_PENDING, now),
        )
        run_id = cur.lastrowid
        for step in _FIXED_STEPS:
            con.execute(
                "INSERT INTO run_steps (run_id, step_name, status) VALUES (?,?,?)",
                (run_id, step, STATUS_PENDING),
            )
    print(f"[DB] Registered  run_id={run_id}  msg={message_id[:40]}…")
    return run_id


def ensure_dynamic_step(run_id: int, step_name: str):
    """Add a dynamic step row (nis_booking_X / sap_automation_X) if not present."""
    with _conn() as con:
        con.execute(
            "INSERT OR IGNORE INTO run_steps (run_id, step_name, status) VALUES (?,?,?)",
            (run_id, step_name, STATUS_PENDING),
        )


def get_run_by_message_id(message_id: str) -> Optional[dict]:
    with _conn() as con:
        row = con.execute(
            "SELECT * FROM email_runs WHERE message_id=?", (message_id,)
        ).fetchone()
        return dict(row) if row else None


def get_run_by_id(run_id: int) -> Optional[dict]:
    with _conn() as con:
        row = con.execute(
            "SELECT * FROM email_runs WHERE id=?", (run_id,)
        ).fetchone()
        return dict(row) if row else None


def get_pending_runs() -> list:
    """Return all runs not yet fully done/failed, oldest first."""
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM email_runs WHERE status IN (?,?) ORDER BY created_at",
            (STATUS_PENDING, STATUS_RUNNING),
        ).fetchall()
        return [dict(r) for r in rows]


# ── Data-save helpers (called by pipeline after each step) ───────────────────
def save_extracted_data(run_id: int, metadata_row: dict, spv_dsm_rows: list):
    _update_run(run_id, {
        "metadata_json": json.dumps(metadata_row  or {}),
        "spv_dsm_json":  json.dumps(spv_dsm_rows  or []),
    })


def save_classification(run_id: int, positive: dict, negative: dict):
    _update_run(run_id, {
        "positive_fields_json": json.dumps(positive),
        "negative_fields_json": json.dumps(negative),
    })


def save_vendor_data(run_id: int, vendor_data: dict):
    _update_run(run_id, {"vendor_data_json": json.dumps(vendor_data)})


def save_downloaded_pdfs(run_id: int, pdf_paths: list, checklist_pdf_name: str = ""):
    """
    Persist the list of downloaded PDF filenames (basenames only) and, if
    identifiable, the specific checklist PDF filename.
    Also reads and stores the binary content of the first downloaded PDF
    (email attachment) into downloaded_pdf_blob.

    Args:
        pdf_paths          : list of full or relative file paths downloaded
        checklist_pdf_name : basename of the DSM checklist PDF (empty if unknown)
    """
    import os as _os

    basenames = [str(p).replace("\\", "/").split("/")[-1] for p in (pdf_paths or [])]

    # Read binary content of first PDF (email attachment)
    pdf_blob = None
    for path in (pdf_paths or []):
        try:
            with open(path, "rb") as fh:
                pdf_blob = fh.read()
            break
        except Exception as e:
            print(f"[DB] Could not read PDF blob from {path}: {e}")

    fields: dict = {
        "downloaded_pdf_names": json.dumps(basenames),
        "checklist_pdf_name":   checklist_pdf_name or "",
    }
    if pdf_blob is not None:
        fields["downloaded_pdf_blob"] = pdf_blob

    _update_run(run_id, fields)
    print(f"[DB] Saved {len(basenames)} PDF name(s) for run_id={run_id}  "
          f"checklist='{checklist_pdf_name or 'n/a'}'"
          f"  blob={'yes' if pdf_blob else 'no'} ({len(pdf_blob) if pdf_blob else 0} bytes)")


def save_nis_checklist_id(run_id: int, checklist_id: str):
    """Persist the NIS checklist / booking number returned by the portal."""
    _update_run(run_id, {"nis_checklist_id": str(checklist_id or "")})
    print(f"[DB] Saved NIS checklist_id='{checklist_id}' for run_id={run_id}")


def save_checklist_pdf_blob(run_id: int, checklist_pdf_path: str):
    """
    Read the checklist PDF from disk and store its binary content in
    checklist_pdf_blob.  Also updates checklist_pdf_name with the basename.

    Args:
        checklist_pdf_path : full path to the checklist PDF file
    """
    from pathlib import Path as _Path
    p = _Path(checklist_pdf_path)
    try:
        blob = p.read_bytes()
        _update_run(run_id, {
            "checklist_pdf_blob": blob,
            "checklist_pdf_name": p.name,
        })
        print(f"[DB] Saved checklist PDF blob for run_id={run_id}  "
              f"file='{p.name}'  ({len(blob)} bytes)")
    except Exception as e:
        print(f"[DB] Could not read checklist PDF blob from {checklist_pdf_path}: {e}")


def save_checklist_value(run_id: int, checklist_value: str):
    """
    Persist the human-readable checklist / booking value captured from the
    NIS portal confirmation screen.

    Args:
        checklist_value : the checklist number / value string from the portal
    """
    _update_run(run_id, {"checklist_value": str(checklist_value or "")})
    print(f"[DB] Saved checklist_value='{checklist_value}' for run_id={run_id}")


def save_spv_name(run_id: int, spv_name: str):
    _update_run(run_id, {"spv_name": spv_name})


def mark_run_started(run_id: int):
    _update_run(run_id, {"status": STATUS_RUNNING, "started_at": _now()})


def mark_run_done(run_id: int):
    _update_run(run_id, {"status": STATUS_DONE, "finished_at": _now()})


def mark_run_failed(run_id: int, error: str):
    _update_run(run_id, {
        "status":        STATUS_FAILED,
        "finished_at":   _now(),
        "error_message": str(error)[:2000],
    })


def _update_run(run_id: int, fields: dict):
    set_clause = ", ".join(f"{k}=?" for k in fields)
    with _conn() as con:
        con.execute(
            f"UPDATE email_runs SET {set_clause} WHERE id=?",
            (*fields.values(), run_id),
        )


# ─────────────────────────────────────────────────────────────────────────────
# run_steps helpers
# ─────────────────────────────────────────────────────────────────────────────
def get_steps(run_id: int) -> list:
    with _conn() as con:
        rows = con.execute(
            "SELECT * FROM run_steps WHERE run_id=? ORDER BY id", (run_id,)
        ).fetchall()
        return [dict(r) for r in rows]


def get_step(run_id: int, step_name: str) -> Optional[dict]:
    with _conn() as con:
        row = con.execute(
            "SELECT * FROM run_steps WHERE run_id=? AND step_name=?",
            (run_id, step_name),
        ).fetchone()
        return dict(row) if row else None


def is_step_done(run_id: int, step_name: str) -> bool:
    s = get_step(run_id, step_name)
    return s is not None and s["status"] == STATUS_DONE


def is_step_skipped(run_id: int, step_name: str) -> bool:
    s = get_step(run_id, step_name)
    return s is not None and s["status"] == STATUS_SKIPPED


def step_start(run_id: int, step_name: str, detail: str = ""):
    with _conn() as con:
        con.execute(
            """
            UPDATE run_steps
               SET status=?, started_at=?, detail=?, error=NULL, finished_at=NULL
             WHERE run_id=? AND step_name=?
            """,
            (STATUS_RUNNING, _now(), detail, run_id, step_name),
        )
    print(f"  [DB] ▶ START  {step_name}")


def step_done(run_id: int, step_name: str, detail: str = ""):
    with _conn() as con:
        con.execute(
            """
            UPDATE run_steps
               SET status=?, finished_at=?, detail=?
             WHERE run_id=? AND step_name=?
            """,
            (STATUS_DONE, _now(), detail, run_id, step_name),
        )
    print(f"  [DB] ✅ DONE   {step_name}")


def step_failed(run_id: int, step_name: str, error: str):
    with _conn() as con:
        con.execute(
            """
            UPDATE run_steps
               SET status=?, finished_at=?, error=?
             WHERE run_id=? AND step_name=?
            """,
            (STATUS_FAILED, _now(), str(error)[:2000], run_id, step_name),
        )
    print(f"  [DB] ❌ FAILED {step_name}  →  {str(error)[:100]}")


def step_skip(run_id: int, step_name: str, reason: str = ""):
    with _conn() as con:
        con.execute(
            """
            UPDATE run_steps
               SET status=?, finished_at=?, detail=?
             WHERE run_id=? AND step_name=?
            """,
            (STATUS_SKIPPED, _now(), reason, run_id, step_name),
        )
    print(f"  [DB] ⏭  SKIP   {step_name}  ({reason})")


# ─────────────────────────────────────────────────────────────────────────────
# Retry helpers  (used by NIS booking and SAP automation)
# ─────────────────────────────────────────────────────────────────────────────
def get_step_retry_count(run_id: int, step_name: str) -> int:
    """Return the current retry_count for a step row (0 if the row does not exist)."""
    with _conn() as con:
        row = con.execute(
            "SELECT retry_count FROM run_steps WHERE run_id=? AND step_name=?",
            (run_id, step_name),
        ).fetchone()
        return int(row[0]) if row else 0


def increment_step_retry(run_id: int, step_name: str, error: str) -> int:
    """
    Increment retry_count, record the latest error, and reset the step status
    back to 'pending' so the normal resume path will attempt it again on the
    next call.

    Returns the NEW retry_count after incrementing.
    """
    with _conn() as con:
        con.execute(
            """
            UPDATE run_steps
               SET retry_count = retry_count + 1,
                   status      = ?,
                   error       = ?,
                   finished_at = ?
             WHERE run_id=? AND step_name=?
            """,
            (STATUS_PENDING, str(error)[:2000], _now(), run_id, step_name),
        )
        row = con.execute(
            "SELECT retry_count FROM run_steps WHERE run_id=? AND step_name=?",
            (run_id, step_name),
        ).fetchone()
        new_count = int(row[0]) if row else 1
    print(f"  [DB] 🔁 RETRY  {step_name}  (attempt {new_count})")
    return new_count


# ─────────────────────────────────────────────────────────────────────────────
# Console summary
# ─────────────────────────────────────────────────────────────────────────────
def print_run_summary(run_id: int):
    run   = get_run_by_id(run_id)
    steps = get_steps(run_id)
    icons = {
        STATUS_DONE:    "✅",
        STATUS_FAILED:  "❌",
        STATUS_RUNNING: "🔄",
        STATUS_SKIPPED: "⏭ ",
        STATUS_PENDING: "⏳",
    }
    print(f"\n{'═'*68}")
    print(f"  RUN #{run_id}  |  {run['subject']}")
    print(f"  Overall  : {run['status']}")
    print(f"  SPV      : {run['spv_name']}")
    print(f"  Started  : {run['started_at']}   Finished : {run['finished_at']}")
    if run["error_message"]:
        print(f"  Error    : {run['error_message'][:120]}")
    print("  Steps    :")
    for s in steps:
        icon = icons.get(s["status"], "?")
        dur  = ""
        if s["started_at"] and s["finished_at"]:
            try:
                fmt  = "%Y-%m-%dT%H:%M:%SZ"
                secs = int((
                    datetime.strptime(s["finished_at"], fmt) -
                    datetime.strptime(s["started_at"],  fmt)
                ).total_seconds())
                dur  = f" ({secs}s)"
            except Exception:
                pass
        retry_tag = f"  [retries={s.get('retry_count', 0)}]" if s.get("retry_count") else ""
        print(f"    {icon}  {s['step_name']:<38} {s['status']}{dur}{retry_tag}")
        if s["error"]:
            print(f"         └─ {s['error'][:90]}")
    print(f"{'═'*68}\n")


# ─────────────────────────────────────────────────────────────────────────────
# Utility
# ─────────────────────────────────────────────────────────────────────────────
def _now() -> str:
    return datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ")