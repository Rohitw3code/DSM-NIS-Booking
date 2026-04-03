# DSM-NIS Booking Automation Pipeline

An end-to-end automation pipeline for **AGEL (Adani Green Energy Limited)** that monitors a dedicated Outlook mailbox for DSM (Deviation Settlement Mechanism) emails, classifies the SPV financial fields, and automatically routes each transaction to either the **NIS portal** (positive amounts) or **SAP GUI** (negative amounts), then writes the result back to a SharePoint Excel tracker.

---

## Table of Contents

- [Overview](#overview)
- [Architecture](#architecture)
- [Project Structure](#project-structure)
- [Pipeline Steps](#pipeline-steps)
- [Data Flow](#data-flow)
- [Module Reference](#module-reference)
- [Configuration](#configuration)
- [Environment Variables](#environment-variables)
- [Installation](#installation)
- [Running the Pipeline](#running-the-pipeline)
- [Resume / Fault Tolerance](#resume--fault-tolerance)
- [SharePoint Excel Sheet Layout](#sharepoint-excel-sheet-layout)

---

## Overview

```
Outlook Email (DSM-NIS-BOOKING)
        │
        ▼
  Read & Parse Email
        │
        ▼
  Classify SPV Amounts
   ┌────┴────┐
   │         │
+ve amt   -ve amt
   │         │
   ▼         ▼
NIS Portal  SAP GUI
(Playwright) (COM)
   │         │
   └────┬────┘
        │
        ▼
 SharePoint Excel Tracker
  (NEW-DSM-NIS-Booking.xlsx)
```

The pipeline is **resume-aware**: every step is recorded in a local SQLite database. If the process crashes or is interrupted, restarting it automatically picks up from the first incomplete step.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────────────┐
│                          main.py  (Orchestrator)                    │
│                                                                     │
│  ┌──────────┐  ┌───────────┐  ┌──────────┐  ┌──────────────────┐  │
│  │read_email│  │classified_│  │read_vend-│  │add_new_row_      │  │
│  │   .py    │  │html_table_│  │or_master_│  │data_nis.py       │  │
│  │          │  │parser.py  │  │data.py   │  │(SharePoint write)│  │
│  └────┬─────┘  └─────┬─────┘  └────┬─────┘  └────────┬─────────┘  │
│       │              │              │                  │            │
│  ┌────▼──────────────▼──────────────▼──────────────────▼────────┐  │
│  │                    db_manager.py  (SQLite)                    │  │
│  │        email_runs table  +  run_steps table                   │  │
│  └───────────────────────────────────────────────────────────────┘  │
│                                                                     │
│  ┌─────────────────────────┐   ┌─────────────────────────────────┐ │
│  │     nis_booking.py      │   │  sap_automation.py              │ │
│  │  (Playwright / Edge)    │   │  (SAP GUI COM / win32com)       │ │
│  │                         │   │                                 │ │
│  │  1. Password login      │   │  1. Launch saplogon.exe         │ │
│  │  2. MFA / RSA OTP       │   │  2. Attach COM scripting        │ │
│  │  3. Create NIS form     │   │  3. Login → ZDCC tcode          │ │
│  │  4. Upload PDF          │   │  4. Fill Non-PO form            │ │
│  │  5. Fill vendor/invoice │   │  5. Save → Print PDF            │ │
│  │  6. Add expense items   │   │  6. Capture checklist number    │ │
│  │  7. Select approver     │   │  7. Close SAP windows           │ │
│  │  8. Submit              │   │                                 │ │
│  │  9. Capture NIS number  │   │                                 │ │
│  └─────────────────────────┘   └─────────────────────────────────┘ │
│                                                                     │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │                       config.py                              │   │
│  │   All credentials, URLs, passwords — one place to manage     │   │
│  └──────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
```

### External Systems

```
┌─────────────────────────────────────────────────────────────────┐
│                    Microsoft Cloud (Azure AD)                   │
│                                                                 │
│  ┌───────────────────┐        ┌──────────────────────────────┐  │
│  │  Outlook Mailbox  │        │  SharePoint / OneDrive       │  │
│  │  (Graph API)      │        │  AGEL-Automation site        │  │
│  │                   │        │  NEW-DSM-NIS-Booking.xlsx    │  │
│  │  agel.nis.        │        │  (vendor master + tracker)   │  │
│  │  automation@      │        │                              │  │
│  │  adani.com        │        │                              │  │
│  └───────────────────┘        └──────────────────────────────┘  │
└─────────────────────────────────────────────────────────────────┘

┌──────────────────────────┐    ┌──────────────────────────────────┐
│  NIS Fiori Launchpad     │    │  SAP GUI (Windows Desktop)       │
│  (SAP BTP / HANA Cloud)  │    │  Transaction: ZDCC               │
│  ael-nis-prd.launchpad   │    │  Non-PO Vendor Payment           │
│  .cfapps.in30.hana...    │    │  Checklist PDF printed locally   │
└──────────────────────────┘    └──────────────────────────────────┘
```

---

## Project Structure

```
start7/
│
├── main.py                          # Orchestrator & entry point
├── config.py                        # ★ All credentials, URLs & settings
│
├── read_email.py                    # Fetch DSM emails via Microsoft Graph API
├── classified_html_table_parser.py  # Parse metadata & SPV tables from email HTML
│
├── read_vendor_master_data.py       # Pull vendor row from SharePoint Excel
├── add_new_row_data_nis.py          # Write result row to SharePoint Excel tracker
│
├── nis_booking.py                   # End-to-end NIS portal automation (Playwright)
├── loginwithpassword.py             # Standalone NIS login script (dev/test use)
│
├── sap_automation.py                # SAP GUI automation via COM scripting
├── new_sap_automation.py            # Alternative / updated SAP automation flow
│
├── db_manager.py                    # SQLite persistence layer (resume support)
│
├── requirements.txt                 # Python dependencies
├── README.md                        # This file
└── automation.db                    # SQLite database (auto-created on first run)
```

---

## Pipeline Steps

Each step is tracked individually in SQLite with status: `pending → running → done / failed / skipped`.

```
Step 1 │ read_email
       │ Fetch the DSM email body HTML from Outlook via Graph API.
       │ Attachment PDFs are downloaded to ./attachments/{message_id}/
       │
Step 2 │ email_extraction
       │ Parse two HTML tables from the email body:
       │   • metadata_table  → week_no, from_date, to_date, due_date, published_date
       │   • spv_dsm_table   → spv_name, total_dsm_charges_payable, drawl_charges_payable,
       │                        revenue_diff, revenue_loss, net_payable_receivable
       │
Step 3 │ email_classification
       │ Classify each SPV amount field:
       │   > 0  → positive  → routed to NIS booking
       │   < 0  → negative  → routed to SAP automation (absolute value used)
       │   = 0  → skipped
       │
Step 4 │ email_validation
       │ Confirm all required metadata fields are present and at least one
       │ amount field is actionable. Fails fast with a clear error message.
       │
Step 5 │ read_vendor_master
       │ Read vendor master row from SharePoint Excel (sheet = spv_name):
       │   vendor_code, company_code, cost_center, plant, bank_key, purpose
       │
Step 6 │ nis_booking_<field>          [one step per positive SPV field]
       │ Playwright automation against the NIS Fiori Launchpad:
       │   • Password + MFA login (email → password → RSA OTP → Yes)
       │   • Upload attachment PDF as header document
       │   • Fill Sub Doc Type, Company Code → Process
       │   • Fill Vendor, Bank Key, Invoice Number, Invoice Date
       │   • Add expense item(s): category, sub-category, amount,
       │     cost center, plant, remarks
       │   • Navigate approval screen → select approver radio
       │   • Click Submit
       │   • Capture NIS booking number from success popup
       │   • Click OK to dismiss popup
       │
Step 7 │ sap_automation_<field>       [one step per negative SPV field]
       │ SAP GUI COM scripting:
       │   • Launch saplogon.exe → attach COM scripting engine
       │   • Login with client / username / password
       │   • Navigate to ZDCC transaction
       │   • Select Non-PO Based → Create checklist
       │   • Fill company code, vendor, year, amount, cost center, plant
       │   • Fill debit note date, doc sub-category, tick documents checkbox
       │   • Preview → Save → confirm popups
       │   • Trigger PDF! print → detect & copy PDF from Temp folder
       │   • Capture checklist number from status bar
       │   • Close all SAP windows
       │
Step 8 │ update_excel_tracker
       │ Write one new row to SharePoint Excel (columns D–O):
       │   D=Week_No, E=Week range, F=Published date, H=DSM charges,
       │   I=Revenue diff, J=Net payable, L=Due date, M=NIS booking number
```

---

## Data Flow

```
Outlook Email
    │
    │  Graph API (MSAL token)
    ▼
read_email.py
    │  body_html, attachment PDFs
    ▼
classified_html_table_parser.py
    │  metadata_row  { week_no, from_date, to_date, due_date, published_date }
    │  spv_dsm_row   { spv_name, total_dsm_charges_payable, revenue_diff, ... }
    ▼
main.py  ─── classify amounts ──► positive_fields  { field: amount }
         └──────────────────────► negative_fields  { field: abs_amount }
    │
    ├── positive fields ──────────────────────────────────────────┐
    │                                                             │
    │   read_vendor_master_data.py (SharePoint Graph API)         │
    │     └► vendor_code, company_code, bank_key,                 │
    │        cost_center, plant, purpose                          │
    │                                                             ▼
    │                                                    nis_booking.py
    │                                                       (Playwright)
    │                                                    NIS booking number
    │                                                    saved → SQLite
    │
    ├── negative fields ──────────────────────────────────────────┐
    │                                                             │
    │   read_vendor_master_data.py (SharePoint Graph API)         │
    │     └► vendor_code, company_code,                           │
    │        cost_center, plant                                   │
    │                                                             ▼
    │                                                   sap_automation.py
    │                                                     (SAP GUI COM)
    │                                                   SAP checklist number
    │                                                   + PDF saved locally
    │                                                   saved → SQLite
    │
    └── update tracker ──────────────────────────────────────────►
                                                        add_new_row_data_nis.py
                                                        (SharePoint Graph API)
                                                        Write row D..O to
                                                        NEW-DSM-NIS-Booking.xlsx
```

---

## Module Reference

| File | Role |
|---|---|
| `main.py` | Entry point. Orchestrates all 8 pipeline steps in order. Supports `--mode scheduled` (poll loop) and `--mode immediate` (single-shot). |
| `config.py` | Single source of truth for all credentials, URLs, and settings. Every value can be overridden via environment variable. |
| `read_email.py` | Fetches emails from Outlook via Microsoft Graph API using MSAL. Filters by `DSM-NIS-BOOKING` subject keyword. Downloads PDF attachments. |
| `classified_html_table_parser.py` | Pure-Python HTML parser. Extracts `metadata_table` and `spv_dsm_table` from the raw email body HTML and returns typed dicts. |
| `read_vendor_master_data.py` | Reads vendor master data (vendor code, company code, cost center, plant, bank key) from a SharePoint Excel file via Graph API. |
| `nis_booking.py` | Full Playwright browser automation for the NIS Fiori Launchpad. Handles login (password + RSA MFA), form filling, PDF upload, expense items, approver selection, submit, and NIS number capture. |
| `loginwithpassword.py` | Standalone script for testing the NIS login flow independently. |
| `sap_automation.py` | SAP GUI automation via `win32com.client` COM scripting. Fills the ZDCC Non-PO checklist form, saves, prints PDF, and captures the checklist number. |
| `new_sap_automation.py` | Updated/alternative SAP automation flow with the same public interface as `sap_automation.py`. |
| `add_new_row_data_nis.py` | Writes one new data row (columns D–O) to the SharePoint Excel tracker sheet via Graph API PATCH. |
| `db_manager.py` | SQLite persistence layer. Tracks every email run and every pipeline step with status, timestamps, retry counts, and captured values. Enables full resume on crash. |

---

## Configuration

All configuration lives in **`config.py`**. Never hard-code credentials in any other file.

| Section | Key Variables |
|---|---|
| **Graph API – Email** | `GRAPH_TENANT_ID`, `GRAPH_CLIENT_ID`, `GRAPH_CLIENT_SECRET`, `USER_EMAIL` |
| **Graph API – SharePoint** | `SP_TENANT_ID`, `SP_CLIENT_ID`, `SP_CLIENT_SECRET` |
| **SharePoint location** | `TENANT_HOST`, `SITE_NAME`, `DRIVE_NAME`, `FOLDER_PATH`, `FILE_NAME`, `TARGET_SHEET` |
| **NIS Portal** | `NIS_LOGIN_URL`, `NIS_EDGE_PATH`, `NIS_USERNAME`, `NIS_PASSWORD`, `NIS_OTP_SECRET` |
| **SAP GUI** | `SAP_LOGON_PATH`, `SAP_SYSTEM`, `SAP_CLIENT`, `SAP_USERNAME`, `SAP_PASSWORD`, `SAP_LANGUAGE` |
| **SAP Form defaults** | `SAP_CHK_DOC_TYP`, `SAP_DIG_SIGN_INV`, `SAP_GJAHR`, `SAP_DEB_NOT_REF`, `SAP_DOC_SUB_CAT` |
| **SAP PDF paths** | `SAP_PDF_TEMP_DIR`, `SAP_PDF_TARGET_DIR` |
| **Pipeline** | `DSM_SUBJECT_KEYWORD`, `DB_PATH`, `ATTACHMENT_DIR` |

---

## Environment Variables

Override any config value without touching `config.py`:

```bash
# Azure AD – Email
export GRAPH_TENANT_ID="your-tenant-id"
export GRAPH_CLIENT_ID="your-app-client-id"
export GRAPH_CLIENT_SECRET="your-secret"
export USER_EMAIL="automation@yourcompany.com"

# Azure AD – SharePoint
export SP_CLIENT_ID="your-sharepoint-app-id"
export SP_CLIENT_SECRET="your-sharepoint-secret"

# NIS Portal
export NIS_USERNAME="nis.user@yourcompany.com"
export NIS_PASSWORD="your-nis-password"
export NIS_OTP_SECRET="your-rsa-otp"
export NIS_EDGE_PATH="C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"

# SAP GUI
export SAP_LOGON_PATH="C:\Program Files (x86)\SAP\FrontEnd\SAPGUI\saplogon.exe"
export SAP_SYSTEM="Renewable Production SSO"
export SAP_CLIENT="910"
export SAP_USERNAME="AUTO10051"
export SAP_PASSWORD="your-sap-password"

# SAP PDF output
export SAP_PDF_TARGET_DIR="D:\DSM\CheckList-PDF"
```

---

## Installation

### Prerequisites
- Python 3.10 or higher
- Windows machine (required for SAP GUI automation via `pywin32` / `pywinauto`)
- Microsoft Edge browser installed at the path set in `NIS_EDGE_PATH`
- SAP GUI (SAP Logon) installed and configured with the system connection

### Steps

```bash
# 1. Clone the repository
git clone https://github.com/Rohitw3code/DSM-NIS-Booking.git
cd DSM-NIS-Booking

# 2. Create and activate a virtual environment (recommended)
python -m venv venv
venv\Scripts\activate          # Windows
# source venv/bin/activate     # Linux/macOS (non-SAP modules only)

# 3. Install Python dependencies
pip install -r requirements.txt

# 4. Install Playwright browser binaries
playwright install chromium

# 5. Edit config.py with your credentials
#    (or set the corresponding environment variables)
notepad config.py
```

---

## Running the Pipeline

### Immediate mode  *(run once and exit)*
```bash
python main.py --mode immediate --lookback 2
```

### Scheduled mode  *(poll every N seconds)*
```bash
# Poll every 5 minutes (default), scan last 2 days of email
python main.py --mode scheduled --interval 300 --lookback 2

# Poll every 10 minutes, scan last 1 day
python main.py --mode scheduled --interval 600 --lookback 1
```

### Arguments

| Argument | Default | Description |
|---|---|---|
| `--mode` | required | `immediate` or `scheduled` |
| `--interval` | `300` | Seconds between polls (scheduled mode only) |
| `--lookback` | `2` | Number of past days of email to scan |

---

## Resume / Fault Tolerance

Every email run and every pipeline step is persisted in `automation.db` (SQLite):

```
email_runs table
  id, message_id, subject, status, metadata_json, spv_dsm_json,
  vendor_data_json, nis_checklist_id, checklist_pdf_name, ...

run_steps table
  run_id, step_name, status, attempt_count, detail, error, started_at, done_at
```

**Step statuses:** `pending → running → done | failed | skipped`

If the pipeline crashes mid-run (e.g. browser crash, network error, SAP timeout):
1. Simply restart with the same command.
2. The pipeline reads `automation.db`, finds the run with status `running` or `failed`.
3. It skips all `done` steps and resumes from the first incomplete step.
4. Each step retries up to **3 times** before being marked permanently `failed`.

---

## SharePoint Excel Sheet Layout

Columns **D → O** written by `add_new_row_data_nis.py`:

| Col | Field | Source |
|---|---|---|
| D | Week No | `metadata_table.week_no` |
| E | Week Range | `"from_date to to_date"` |
| F | DSM Statement Published Date | `metadata_table.dsm_statement_published_date` |
| G | *(blank)* | — |
| H | DSM Charges | `spv_dsm_table.total_dsm_charges_payable` |
| I | Revenue Diff Amount | `spv_dsm_table.revenue_diff` |
| J | Net Payable / Receivable | `spv_dsm_table.net_payable_receivable` |
| K | *(blank)* | — |
| L | Due Date | `metadata_table.due_date` |
| M | **NIS Booking Number** | Captured from NIS portal success popup |
| N | *(blank)* | — |
| O | *(blank)* | — |