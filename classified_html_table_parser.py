import re
from html import unescape
from typing import List, Dict, Optional

# --------------------------------------------------
# Header maps  –  key = any normalized substring that
# uniquely identifies the column.
# Multiple keys can map to the same field (aliases).
# The LONGEST matching key wins (avoids "date" matching
# both "from date" and "to date" ambiguously).
# --------------------------------------------------
METADATA_HEADER_MAP = {
    # ── week number ────────────────────────────────
    "week no":          "week_no",
    "week number":      "week_no",
    "weekno":           "week_no",

    # ── period start ───────────────────────────────
    "from date":        "from_date",
    "date from":        "from_date",
    "start date":       "from_date",
    "week from":        "from_date",
    "period from":      "from_date",

    # ── period end ─────────────────────────────────
    "to date":          "to_date",
    "date to":          "to_date",
    "end date":         "to_date",
    "week to":          "to_date",
    "period to":        "to_date",
    "week end":         "to_date",

    # ── published date ──────────────────────────────
    "dsm statement published date": "dsm_statement_published_date",
    "statement published":          "dsm_statement_published_date",
    "published date":               "dsm_statement_published_date",

    # ── due date ────────────────────────────────────
    "due date":         "due_date",
    "payment due":      "due_date",
}

SPV_DSM_HEADER_MAP = {
    "spv name":                     "spv_name",
    "spv":                          "spv_name",
    "total dsm charges payable":    "total_dsm_charges_payable",
    "dsm charges":                  "total_dsm_charges_payable",
    "drawl charges payable":        "drawl_charges_payable",
    "drawl charges":                "drawl_charges_payable",
    "revenue diff":                 "revenue_diff",
    "revenue difference":           "revenue_diff",
    # "revenue_loss" column is intentionally excluded per automation spec
    "net payable":                  "net_payable_receivable",
    "net receivable":               "net_payable_receivable",
    "net amount":                   "net_payable_receivable",
}

# --------------------------------------------------
# Utility functions
# --------------------------------------------------
def normalize(text: str) -> str:
    text = unescape(text.lower())
    # ── Step 1: expand meaningful parentheticals BEFORE stripping ────────────
    # "From (Date)" → "from date", "To (Date)" → "to date"
    # Without this, the blanket strip below turns them into bare "from"/"to"
    # which never match the header-map keys "from date" / "to date".
    text = re.sub(r"\(\s*date\s*\)", " date", text)
    # ── Step 2: strip only short label/unit codes like (A), (B), (Rs.), (In) ─
    # Limit to ≤ 10 chars inside parens so long descriptive phrases survive.
    text = re.sub(r"\([^)]{0,10}\)", " ", text)
    # ── Step 3: remove remaining punctuation, collapse whitespace ─────────────
    text = re.sub(r"[^a-z0-9 ]+", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()

def clean_cell(text: str) -> str:
    if not text:
        return ""
    text = unescape(text)
    text = re.sub(r"<br\s*/?>", " ", text, flags=re.I)
    text = re.sub(r"<[^>]+>", "", text)
    return text.strip()

def extract_tables(html: str) -> List[str]:
    return re.findall(r"(?is)<table.*?</table>", html)

def extract_rows(table_html: str) -> List[str]:
    return re.findall(r"(?is)<tr.*?</tr>", table_html)

def extract_cells(row_html: str) -> List[str]:
    return re.findall(r"(?is)<t[dh].*?>(.*?)</t[dh]>", row_html)

# --------------------------------------------------
# Core table parser
# --------------------------------------------------
def parse_single_table(table_html: str) -> Optional[dict]:
    rows_html = extract_rows(table_html)
    rows = []

    for r in rows_html:
        cells = [clean_cell(c) for c in extract_cells(r)]
        if any(cells):
            rows.append(cells)

    if len(rows) < 2:
        return None

    headers = [normalize(h) for h in rows[0]]
    data_rows = rows[1:]

    return {
        "headers":     headers,
        "raw_headers": rows[0],
        "data":        data_rows,
    }

def match_headers(headers: List[str], header_map: Dict[str, str]) -> bool:
    return sum(1 for h in headers for k in header_map if k in h) >= 1

def build_table(parsed: dict, header_map: Dict[str, str]) -> List[Dict]:
    """
    Map column indices to field names using a longest-match-wins strategy:
    for each header cell, all matching keys are collected and the LONGEST
    key string is declared the winner.  This prevents short aliases like
    "date" or "spv" from stealing a column that a longer, more specific
    key (e.g. "from date", "spv name") also matches.
    """
    headers   = parsed["headers"]
    data_rows = parsed["data"]

    idx_map: Dict[int, str] = {}
    for i, h in enumerate(headers):
        best_key    = ""
        best_mapped = ""
        for key, mapped in header_map.items():
            if key in h and len(key) > len(best_key):
                best_key    = key
                best_mapped = mapped
        if best_mapped:
            idx_map[i] = best_mapped

    result = []
    for row in data_rows:
        record = {
            mapped: row[i] if i < len(row) else ""
            for i, mapped in idx_map.items()
        }
        if any(v.strip() for v in record.values()):
            result.append(record)

    return result

# --------------------------------------------------
# Debug helper – prints raw extracted table structure
# --------------------------------------------------
def print_raw_tables(body_html: str) -> None:
    """
    Pretty-print every <table> found in the HTML so you can manually verify
    what the parser actually received before any mapping is applied.

    For each table this prints:
      1. The raw HTML snippet (truncated to 2 000 chars so the log stays readable)
      2. The raw cell text of EVERY row (no cleaning, no normalisation)
      3. The normalised headers after clean_cell + normalize()
      4. Up to 5 data rows after clean_cell()
    """
    tables = extract_tables(body_html)
    print(f"\n{'='*70}")
    print(f"[RAW EXTRACT] {len(tables)} table(s) found in email body")
    print(f"{'='*70}")

    for t_idx, t in enumerate(tables):
        print(f"\n---- Table[{t_idx}] raw HTML (first 2000 chars) ----")
        print(t[:2000])
        if len(t) > 2000:
            print(f"  ... [{len(t) - 2000} more chars truncated]")

        rows_html = extract_rows(t)
        print(f"\n---- Table[{t_idx}] raw cell text (NO cleaning, {len(rows_html)} row(s)) ----")
        for r_idx, row_html in enumerate(rows_html):
            # pull raw inner text without any clean_cell processing
            raw_cells = [
                re.sub(r"<[^>]+>", "", unescape(cell)).strip()
                for cell in re.findall(r"(?is)<t[dh].*?>(.*?)</t[dh]>", row_html)
            ]
            print(f"  Row[{r_idx}] raw  : {raw_cells}")

        parsed = parse_single_table(t)
        if not parsed:
            print(f"  (skipped – fewer than 2 usable rows after cleaning)")
            continue

        print(f"\n---- Table[{t_idx}] after clean_cell + normalize ----")
        print(f"  Raw  headers : {parsed['raw_headers']}")
        print(f"  Norm headers : {parsed['headers']}")
        for r_idx, row in enumerate(parsed["data"][:5]):
            print(f"  Data row[{r_idx}]  : {row}")
        if len(parsed["data"]) > 5:
            print(f"  ... ({len(parsed['data']) - 5} more data rows)")

    print(f"\n{'='*70}\n")

# --------------------------------------------------
# ✅ PUBLIC FUNCTION
# --------------------------------------------------
def parse_html_tables(body_html: str) -> Dict[str, List[Dict]]:
    tables   = extract_tables(body_html)
    metadata = []
    spv_dsm  = []

    for t in tables:
        parsed = parse_single_table(t)
        if not parsed:
            continue

        headers = parsed["headers"]

        if match_headers(headers, METADATA_HEADER_MAP):
            metadata = build_table(parsed, METADATA_HEADER_MAP)

        if match_headers(headers, SPV_DSM_HEADER_MAP):
            spv_dsm = build_table(parsed, SPV_DSM_HEADER_MAP)

    return {
        "metadata_table": metadata,
        "spv_dsm_table":  spv_dsm,
    }