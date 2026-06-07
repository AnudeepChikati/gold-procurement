"""
Gold Procurement TXT to Excel Converter
========================================
Parses Oracle fixed-width PO Transaction Reports and converts them to
readable Excel files with summary pivot sheets.

Usage:
    python parse_gold_report.py                   # processes all *.txt in current folder
    python parse_gold_report.py "FEB'25.txt"      # single file
    python parse_gold_report.py "FEB'25.txt" "Mar'25.txt"  # multiple files
"""

import sys
import re
import os
import glob
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# Column definitions (name, friendly label)
# Order must match the separator groups in the report
# ─────────────────────────────────────────────────────────────────────────────
COLUMN_LABELS = [
    ("po_number",            "PO Number"),
    ("po_date",              "PO Date"),
    ("trans_date",           "Trans Date"),
    ("vendor_code",          "Vendor Code"),
    ("vendor_name",          "Vendor Name"),
    ("vendor_site_code",     "Vendor Site Code"),
    ("state",                "State"),
    ("receipt_number",       "Receipt Number"),
    ("receipt_date",         "Receipt Date"),
    ("material_code",        "Material Code"),
    ("material_description", "Material Description"),
    ("lot_number",           "Lot Number"),
    ("orgn_code",            "Orgn Code"),
    ("subinv_code",          "SubInv Code"),
    ("location",             "Location"),
    ("prim_uom",             "Prim UOM"),
    ("sec_uom",              "Sec UOM"),
    ("primary_quantity",     "Primary Qty (GMS)"),
    ("secondary_quantity",   "Secondary Qty"),
    ("purchase_value",       "Purchase Value (₹)"),
    ("gold_rate",            "Gold Rate"),
    ("gold_net_weight",      "Gold Net Weight"),
    ("gold_converted_weight","Gold Converted Weight"),
    ("gold_value",           "Gold Value"),
    ("stone_value",          "Stone Value"),
    ("loss_value",           "Loss Value"),
    ("labour_charges",       "Labour Charges"),
    ("total_conv_weight",    "Total Conv Weight"),
    ("total_gold_value",     "Total Gold Value"),
    ("tax_perc_1",           "Tax % (1)"),
    ("tax_perc_2",           "Tax % (2)"),
    ("tax_amount",           "Tax Amount (₹)"),
    ("gstin",                "GSTIN"),
    ("glitem_class",         "GL Item Class"),
    ("purity",               "Purity"),
    ("indicator",            "Indicator"),
    ("other_mat_wt",         "Other Mat Wt"),
    ("cfa_code",             "CFA Code"),
    ("attribute_category",   "Attribute Category"),
]

FIELD_NAMES  = [c[0] for c in COLUMN_LABELS]
HEADER_NAMES = [c[1] for c in COLUMN_LABELS]

# Columns that should be treated as numbers
NUMERIC_COLS = {
    "primary_quantity", "secondary_quantity", "purchase_value",
    "gold_rate", "gold_net_weight", "gold_converted_weight",
    "gold_value", "stone_value", "loss_value", "labour_charges",
    "total_conv_weight", "total_gold_value",
    "tax_perc_1", "tax_perc_2", "tax_amount",
    "purity", "other_mat_wt",
}

# ─────────────────────────────────────────────────────────────────────────────
# Patterns to detect noise lines (headers, separators, blanks)
# ─────────────────────────────────────────────────────────────────────────────
_RE_PAGE_HEADER   = re.compile(
    r'^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)',
    re.IGNORECASE
)
_RE_SEPARATOR     = re.compile(r'^-{5,}')
_RE_COL_LABEL     = re.compile(r'^(PO\s+PO\s+Trans|Number\s+Date|---)')
_RE_REPORT_TITLE  = re.compile(r'TJD CST PO TRANSACTIONS REPORT', re.IGNORECASE)
_RE_NON_HDR       = re.compile(r'^\s*(Non|Gold\s+Rec\s+Recover|other)\s*$', re.IGNORECASE)

# A data line starts with a 7-digit PO number
_RE_DATA_LINE     = re.compile(r'^\d{7}\s')

# ─────────────────────────────────────────────────────────────────────────────
# Helper: derive column char-position ranges from the separator line
# ─────────────────────────────────────────────────────────────────────────────
def _get_colspecs(separator_line: str) -> list[tuple[int, int]]:
    """
    Returns list of (start, end) character positions for each column
    by finding groups of dashes in the separator line.
    """
    specs = []
    for m in re.finditer(r'-+', separator_line):
        specs.append((m.start(), m.end()))
    return specs


# ─────────────────────────────────────────────────────────────────────────────
# Helper: determine vendor type from a row dict
# ─────────────────────────────────────────────────────────────────────────────
def _vendor_type(row: dict) -> str:
    mat_code  = str(row.get("material_code", "")).strip().upper()
    attr_cat  = str(row.get("attribute_category", "")).strip().upper()
    gl_class  = str(row.get("glitem_class", "")).strip().upper()
    lot       = str(row.get("lot_number", "")).strip().upper()

    if gl_class == "RAW MATL":
        return "RAW_MATERIAL"
    if attr_cat == "EXCHANGE":
        return "JEWELLER"
    # Bank purchases: lot number is '2B' and not a jeweller
    if lot == "2B" and mat_code.startswith("11GOR"):
        return "BANK"
    # Watch Division uses used targets / pure gold granules
    if "11GORYS" in mat_code or "11GORYM127" in mat_code:
        return "WATCH_DIV"
    if gl_class == "PRECMATL":
        return "BANK"
    return "OTHER"


# ─────────────────────────────────────────────────────────────────────────────
# Core: parse one TXT file → DataFrame
# ─────────────────────────────────────────────────────────────────────────────
def parse_txt(filepath: str) -> pd.DataFrame:
    path = Path(filepath)
    print(f"\n{'─'*60}")
    print(f"  Parsing: {path.name}")
    print(f"{'─'*60}")

    with open(filepath, encoding="utf-8", errors="replace") as fh:
        lines = fh.readlines()

    # ── Step 1: find the separator line and derive column specs ──
    colspecs   = None
    data_lines = []

    for raw in lines:
        line = raw.rstrip("\n")

        # Skip all noise lines
        if not line.strip():
            continue
        if _RE_PAGE_HEADER.match(line):
            continue
        if _RE_REPORT_TITLE.search(line):
            continue
        if _RE_NON_HDR.match(line):
            continue
        if line.lstrip().startswith("PO") or line.lstrip().startswith("Number"):
            continue
        if _RE_SEPARATOR.match(line.lstrip()):
            # First separator gives us column positions
            if colspecs is None:
                colspecs = _get_colspecs(line)
            continue

        # Only keep lines that look like data (start with 7-digit PO number)
        if _RE_DATA_LINE.match(line.strip()):
            data_lines.append(line)

    if colspecs is None:
        raise ValueError(f"Could not find separator line in {filepath}")

    expected_cols = len(FIELD_NAMES)
    if len(colspecs) < expected_cols:
        # Some report versions have trailing whitespace that merges last cols;
        # pad with reasonable range
        last_end = colspecs[-1][1]
        while len(colspecs) < expected_cols:
            new_start = last_end + 1
            new_end   = new_start + 30
            colspecs.append((new_start, new_end))
            last_end  = new_end

    # Trim to expected number of columns
    colspecs = colspecs[:expected_cols]

    print(f"  Detected {len(colspecs)} columns | {len(data_lines)} data lines")

    # ── Step 2: parse each data line using colspecs ──
    records = []
    for line in data_lines:
        # Pad line to at least the width needed
        padded = line.ljust(colspecs[-1][1] + 5)
        row = {}
        for (fname, _), (start, end) in zip(COLUMN_LABELS, colspecs):
            raw_val = padded[start:end].strip()
            if fname in NUMERIC_COLS:
                # Convert to float; treat blank / '.' / '.00' as 0
                raw_val = raw_val.lstrip("$").replace(",", "").strip()
                try:
                    row[fname] = float(raw_val) if raw_val and raw_val != "." else 0.0
                except ValueError:
                    row[fname] = 0.0
            else:
                row[fname] = raw_val
        records.append(row)

    df = pd.DataFrame(records, columns=FIELD_NAMES)

    # ── Step 3: derive extra columns ──
    # Month label from filename (e.g. "FEB'25.txt" → "FEB-25")
    month_label = re.sub(r"['\s]", "-", path.stem).upper()
    df["month"]       = month_label
    df["source_file"] = path.name

    # Vendor type
    df["vendor_type"] = df.apply(lambda r: _vendor_type(r), axis=1)

    # Normalise State (uppercase + strip)
    df["state_normalized"] = df["state"].str.strip().str.upper()

    # Split Trans Date into date + time
    def _split_trans_date(val):
        val = str(val).strip()
        m = re.match(r'(\d{2}-\w{3}-\d{4})\s+([\d:]+)', val)
        if m:
            return m.group(1), m.group(2)
        return val, ""

    df[["trans_date_only", "trans_time"]] = pd.DataFrame(
        df["trans_date"].apply(_split_trans_date).tolist(),
        index=df.index
    )

    # GST rate label
    def _gst_label(perc):
        if perc == 3.0:
            return "3% (Gold/Silver)"
        if perc == 18.0:
            return "18% (Raw Material)"
        if perc == 0.0:
            return "0% (Export/Exempt)"
        return f"{perc}%"

    df["gst_rate_label"] = df["tax_perc_2"].apply(_gst_label)

    print(f"  Total purchase value : ₹{df['purchase_value'].sum():,.2f}")
    print(f"  Total tax amount     : ₹{df['tax_amount'].sum():,.2f}")
    print(f"  Total qty (GMS)      : {df['primary_quantity'].sum():,.3f}")
    print(f"  Unique vendors       : {df['vendor_name'].nunique()}")

    return df


# ─────────────────────────────────────────────────────────────────────────────
# Excel styling helpers
# ─────────────────────────────────────────────────────────────────────────────
_HEADER_FILL  = PatternFill("solid", fgColor="1F4E79")   # dark blue
_ALT_FILL     = PatternFill("solid", fgColor="EBF3FF")   # very light blue
_SUMMARY_FILL = PatternFill("solid", fgColor="2E7D32")   # dark green
_TOTAL_FILL   = PatternFill("solid", fgColor="FFF9C4")   # light yellow
_THIN_BORDER  = Border(
    left  =Side(style="thin", color="BDBDBD"),
    right =Side(style="thin", color="BDBDBD"),
    top   =Side(style="thin", color="BDBDBD"),
    bottom=Side(style="thin", color="BDBDBD"),
)

def _style_header_row(ws, row_num: int, fill=None):
    fill = fill or _HEADER_FILL
    for cell in ws[row_num]:
        cell.font      = Font(bold=True, color="FFFFFF", size=10)
        cell.fill      = fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border    = _THIN_BORDER


def _autofit_columns(ws, min_width=8, max_width=40):
    for col in ws.columns:
        col_letter = get_column_letter(col[0].column)
        max_len = max_width
        for cell in col:
            try:
                cell_len = len(str(cell.value)) if cell.value else 0
                if cell_len > max_len:
                    max_len = cell_len
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_len + 2, min_width), max_width)


def _write_transactions_sheet(wb, df: pd.DataFrame, sheet_name: str = "Transactions"):
    # Build column list for this sheet:
    # all original + derived useful columns
    display_cols_original = list(HEADER_NAMES)
    derived_cols = {
        "month"           : "Month",
        "vendor_type"     : "Vendor Type",
        "state_normalized": "State (Normalised)",
        "trans_date_only" : "Trans Date (Date)",
        "trans_time"      : "Trans Time",
        "gst_rate_label"  : "GST Rate",
    }

    ws = wb.create_sheet(sheet_name)

    # Build full header
    all_headers = display_cols_original + list(derived_cols.values())
    all_field_keys = FIELD_NAMES + list(derived_cols.keys())

    ws.append(all_headers)
    _style_header_row(ws, 1)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # Write data rows
    INR_FMT  = '#,##0.00'
    QTY_FMT  = '#,##0.000'
    PERC_FMT = '0.00'

    numeric_fmt_map = {
        "primary_quantity"      : QTY_FMT,
        "secondary_quantity"    : QTY_FMT,
        "purchase_value"        : INR_FMT,
        "gold_rate"             : INR_FMT,
        "gold_net_weight"       : QTY_FMT,
        "gold_converted_weight" : QTY_FMT,
        "gold_value"            : INR_FMT,
        "stone_value"           : INR_FMT,
        "loss_value"            : INR_FMT,
        "labour_charges"        : INR_FMT,
        "total_conv_weight"     : QTY_FMT,
        "total_gold_value"      : INR_FMT,
        "tax_perc_1"            : PERC_FMT,
        "tax_perc_2"            : PERC_FMT,
        "tax_amount"            : INR_FMT,
        "purity"                : PERC_FMT,
        "other_mat_wt"          : QTY_FMT,
    }

    records = df[all_field_keys].to_dict("records")

    for row_idx, record in enumerate(records, start=2):
        row_data = [record.get(k, "") for k in all_field_keys]
        ws.append(row_data)
        excel_row = ws[row_idx]

        for cell, field_key in zip(excel_row, all_field_keys):
            cell.border    = _THIN_BORDER
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 0:
                cell.fill = _ALT_FILL
            if field_key in numeric_fmt_map:
                cell.number_format = numeric_fmt_map[field_key]

    # Add totals row
    total_row = ["TOTAL"] + [""] * (len(all_headers) - 1)
    ws.append(total_row)
    tr = ws[ws.max_row]
    tr[0].value = "TOTAL"
    tr[0].font  = Font(bold=True)
    tr[0].fill  = _TOTAL_FILL

    qty_col_idx    = all_field_keys.index("primary_quantity") + 1
    purch_col_idx  = all_field_keys.index("purchase_value")   + 1
    tax_col_idx    = all_field_keys.index("tax_amount")       + 1

    for col_idx in [qty_col_idx, purch_col_idx, tax_col_idx]:
        col_letter = get_column_letter(col_idx)
        cell = tr[col_idx - 1]
        cell.value         = f"=SUM({col_letter}2:{col_letter}{ws.max_row - 1})"
        cell.font          = Font(bold=True)
        cell.fill          = _TOTAL_FILL
        cell.number_format = INR_FMT if col_idx != qty_col_idx else QTY_FMT
        cell.border        = _THIN_BORDER

    _autofit_columns(ws)
    ws.row_dimensions[1].height = 30


def _write_summary_sheet(wb, df: pd.DataFrame, group_cols: list, agg_label: str, sheet_name: str):
    agg = (
        df.groupby(group_cols, dropna=False)
          .agg(
              Transaction_Count  =("po_number",        "count"),
              Total_Qty_GMS      =("primary_quantity",  "sum"),
              Total_Purchase_Value=("purchase_value",   "sum"),
              Total_Tax_Amount   =("tax_amount",        "sum"),
          )
          .reset_index()
          .sort_values("Total_Purchase_Value", ascending=False)
    )

    ws = wb.create_sheet(sheet_name)
    headers = [c.replace("_", " ") for c in agg.columns.tolist()]
    ws.append(headers)
    _style_header_row(ws, 1, fill=PatternFill("solid", fgColor="1B5E20"))
    ws.freeze_panes = "A2"

    agg_records = agg.to_dict("records")
    agg_col_list = agg.columns.tolist()

    for row_idx, record in enumerate(agg_records, start=2):
        ws.append([record.get(c) for c in agg_col_list])
        excel_row = ws[row_idx]
        if row_idx % 2 == 0:
            for cell in excel_row:
                cell.fill = _ALT_FILL
        for cell in excel_row:
            cell.border = _THIN_BORDER
            cell.alignment = Alignment(vertical="center")

        # Format numeric cols
        n = len(group_cols)
        for i, col_name in enumerate(agg_col_list[n:], start=n+1):
            cell = excel_row[i - 1]
            if "Qty" in col_name or "Weight" in col_name:
                cell.number_format = '#,##0.000'
            elif "Count" in col_name:
                cell.number_format = '#,##0'
            else:
                cell.number_format = '#,##0.00'

    # Grand total row
    ws.append([])  # blank
    grand = ["GRAND TOTAL"] + [""] * (len(group_cols) - 1)
    ws.append(grand)
    tr = ws[ws.max_row]
    tr[0].font = Font(bold=True, size=11)
    tr[0].fill = _TOTAL_FILL

    n = len(group_cols)
    for i in range(n, len(agg_col_list)):
        col_letter = get_column_letter(i + 1)
        cell = tr[i]
        cell.value         = f"=SUM({col_letter}2:{col_letter}{ws.max_row - 2})"
        cell.font          = Font(bold=True)
        cell.fill          = _TOTAL_FILL
        cell.number_format = '#,##0.000' if "Qty" in agg_col_list[i] else '#,##0.00'
        cell.border        = _THIN_BORDER

    _autofit_columns(ws)
    ws.row_dimensions[1].height = 30

    print(f"  → Sheet '{sheet_name}': {len(agg)} groups")


def _write_filtered_sheet(wb, df: pd.DataFrame, mask: pd.Series, sheet_name: str):
    subset = df[mask].copy()
    if subset.empty:
        return
    _write_transactions_sheet(wb, subset, sheet_name)
    print(f"  → Sheet '{sheet_name}': {len(subset)} rows")


# ─────────────────────────────────────────────────────────────────────────────
# Main: write Excel workbook from DataFrame
# ─────────────────────────────────────────────────────────────────────────────
def write_excel_to_buffer(df: pd.DataFrame) -> tuple:
    """
    Write Excel to an in-memory buffer and return (bytes, stats_dict).
    Used by the Flask web API.
    """
    import io
    stats = {
        "rows":            len(df),
        "purchase_value":  round(df["purchase_value"].sum(), 2),
        "tax_amount":      round(df["tax_amount"].sum(), 2),
        "qty_gms":         round(df["primary_quantity"].sum(), 3),
        "unique_vendors":  int(df["vendor_name"].nunique()),
        "month":           df["month"].iloc[0] if len(df) else "",
        "vendor_types":    df["vendor_type"].value_counts().to_dict(),
    }
    buf = io.BytesIO()
    _build_workbook(df).save(buf)
    buf.seek(0)
    return buf.read(), stats


def _build_workbook(df: pd.DataFrame):
    from openpyxl import Workbook
    wb = Workbook()
    del wb[wb.sheetnames[0]]

    _write_transactions_sheet(wb, df, "Transactions")
    _write_summary_sheet(wb, df,
        group_cols=["vendor_code", "vendor_name", "state_normalized",
                    "gstin", "vendor_type", "month"],
        agg_label="vendor", sheet_name="By_Vendor")
    _write_summary_sheet(wb, df,
        group_cols=["material_code", "material_description", "purity", "glitem_class", "month"],
        agg_label="material", sheet_name="By_Material")
    _write_summary_sheet(wb, df,
        group_cols=["receipt_date", "month"],
        agg_label="date", sheet_name="By_Date")
    _write_filtered_sheet(wb, df, df["vendor_type"] == "JEWELLER",    "Jeweller_Exchange")
    _write_filtered_sheet(wb, df, df["vendor_type"] == "BANK",        "Bank_Purchases")
    _write_filtered_sheet(wb, df, df["vendor_type"] == "RAW_MATERIAL","Raw_Materials")
    _write_filtered_sheet(wb, df, df["vendor_type"] == "WATCH_DIV",   "Watch_Division")
    return wb


def write_excel(df: pd.DataFrame, out_path: str):
    wb = _build_workbook(df)
    wb.save(out_path)
    print(f"\n  ✓ Saved: {out_path}")



# ─────────────────────────────────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────────────────────────────────
def _infer_month_tag(filepath: str) -> str:
    """Extract month tag from filename like FEB'25.txt → FEB_25"""
    stem = Path(filepath).stem
    tag  = re.sub(r"['\s]+", "_", stem).upper()
    return tag


def process_file(filepath: str):
    if not os.path.exists(filepath):
        print(f"  ✗ File not found: {filepath}")
        return

    df = parse_txt(filepath)

    if df.empty:
        print(f"  ✗ No data rows parsed from {filepath}")
        return

    tag      = _infer_month_tag(filepath)
    out_dir  = Path(filepath).parent
    out_path = str(out_dir / f"GoldProcurement_{tag}.xlsx")

    write_excel(df, out_path)


def main():
    if len(sys.argv) > 1:
        files = sys.argv[1:]
    else:
        # Auto-discover all .txt files in current directory
        files = glob.glob("*.txt")
        if not files:
            print("No .txt files found in current directory.")
            print("Usage: python parse_gold_report.py <file1.txt> [file2.txt ...]")
            sys.exit(1)
        print(f"Auto-discovered {len(files)} file(s): {', '.join(files)}")

    for f in files:
        try:
            process_file(f)
        except Exception as exc:
            print(f"\n  ✗ Error processing {f}: {exc}")
            import traceback
            traceback.print_exc()

    print("\nDone.")


if __name__ == "__main__":
    main()
