import json
import re
from copy import copy
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="Inbound Pallets", layout="wide")

st.title("📦 Pallets per Trailer")

st.write(
    "Upload your inbound report throughout the day. The app keeps the **highest pallet count "
    "seen by the system** for each trailer last 3. At the end of the day, upload the manual "
    "unloader log and the app will match the manual unloaded pallets against the system highest."
)

# -------------------------------------------------------------------
# Transfer Log template
# Put this Excel file in your GitHub repository root.
#
# Your repo can look like this:
# your-repo/
# ├── Home.py
# ├── Transfer Log New 8-2025.xlsx
# ├── requirements.txt
# └── pages/
#     └── Update 7_Inbound_pallets.py
# -------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent if "__file__" in globals() else Path.cwd()

# If this Streamlit page is inside /pages, go one folder up to the repo root.
ROOT_DIR = APP_DIR.parent if APP_DIR.name == "pages" else APP_DIR

TRANSFER_LOG_TEMPLATE = ROOT_DIR / "Transfer Log New 8-2025.xlsx"

# Runtime storage.
# This remembers uploaded loads by selected day while the Streamlit app instance is running.
# It may reset if Streamlit Cloud restarts or redeploys the app.
HISTORY_FILE = Path("/tmp/inbound_pallets_daily_history.json")
MANUAL_HISTORY_FILE = Path("/tmp/inbound_pallets_manual_history.json")

# Transfer Log columns based on the template:
# C = Trailer Num.
# K = Total Pallets UNLOADED      -> Manual unloader count
# L = Total Pallets in this Trailer -> Highest system-recorded count
TRANSFER_LOG_START_ROW = 8
TRANSFER_LOG_TRAILER_COL = "C"
TRANSFER_LOG_MANUAL_UNLOADED_COL = "K"
TRANSFER_LOG_SYSTEM_HIGHEST_COL = "L"

# Clear body columns A:N so everything else stays blank.
TRANSFER_LOG_FIRST_COL = 1
TRANSFER_LOG_LAST_COL = 14

# Inbound report column positions, 0-indexed:
# C,D,E,F,G = 2,3,4,5,6
# LPN # = 7, column H
TRAILER_COLS = [2, 3, 4, 5, 6]
LPN_COL = 7

# Manual log column positions, 0-indexed, based on the Transfer Log template:
# C = Trailer Num.
# K = Total Pallets UNLOADED
MANUAL_LOG_START_ROW = 8
MANUAL_LOG_TRAILER_COL = 2
MANUAL_LOG_UNLOADED_COL = 10

HEADER_ROWS = 3       # first 3 rows are headers in the inbound report
THRESHOLD = 9         # system loads with 9 or fewer pallets get flagged for research


# -------------------------------------------------------------------
# Storage helpers
# -------------------------------------------------------------------
def load_json_file(path):
    """Load saved JSON history from the app runtime."""
    if path.exists():
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    return {}


def save_json_file(path, data):
    """Save JSON history to the app runtime."""
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2)
    except Exception:
        pass


# -------------------------------------------------------------------
# Cleaning helpers
# -------------------------------------------------------------------
def clean_excel_value(value):
    """Clean Excel values so numbers like 123.0 become 123."""
    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()

    return str(value).strip()


def parse_pallet_count(value):
    """Return an integer pallet count, or None if the value is blank/not usable."""
    if pd.isna(value):
        return None

    if isinstance(value, int):
        return int(value)

    if isinstance(value, float):
        if pd.isna(value):
            return None
        return int(round(value))

    text = str(value).strip()

    if text == "":
        return None

    # Accept values like "24", "24 pallets", "24.0".
    match = re.search(r"-?\d+(?:\.\d+)?", text)
    if not match:
        return None

    return int(round(float(match.group())))


def build_trailer(row):
    """Concatenate trailer parts from columns C through G in the inbound report."""
    return "".join(clean_excel_value(row[c]) for c in TRAILER_COLS)


def get_last_3_numbers(trailer):
    """Return the last 3 numeric digits from the trailer value."""
    digits = "".join(re.findall(r"\d", str(trailer)))

    if not digits:
        return ""

    return digits[-3:].zfill(3)


def flag_red(row):
    return ["background-color: #ffb3b3; color: #800000; font-weight: bold"] * len(row)


def flag_variance(row):
    """Highlight variance rows in the comparison table."""
    status = str(row.get("Match Status", ""))

    if "Mismatch" in status or "Missing" in status or "Manual only" in status or "System only" in status:
        return ["background-color: #fff2cc; color: #7a4f00; font-weight: bold"] * len(row)

    return [""] * len(row)


# -------------------------------------------------------------------
# Excel output helpers
# -------------------------------------------------------------------
def copy_row_style(ws, source_row, target_row):
    """Copy formatting from one row to another row."""
    for col_idx in range(TRANSFER_LOG_FIRST_COL, TRANSFER_LOG_LAST_COL + 1):
        source_cell = ws.cell(row=source_row, column=col_idx)
        target_cell = ws.cell(row=target_row, column=col_idx)

        if source_cell.has_style:
            target_cell._style = copy(source_cell._style)

        if source_cell.number_format:
            target_cell.number_format = source_cell.number_format

        if source_cell.alignment:
            target_cell.alignment = copy(source_cell.alignment)

        if source_cell.border:
            target_cell.border = copy(source_cell.border)

        if source_cell.fill:
            target_cell.fill = copy(source_cell.fill)

        if source_cell.font:
            target_cell.font = copy(source_cell.font)

        if source_cell.protection:
            target_cell.protection = copy(source_cell.protection)

    ws.row_dimensions[target_row].height = ws.row_dimensions[source_row].height


def create_transfer_log_excel(transfer_output_df):
    """
    Create a completed Transfer Log from the repository template.

    Fills only:
    - C: Trailer Num.
    - K: Total Pallets UNLOADED / manual unloader count
    - L: Highest system-recorded pallet count

    Everything else in the body rows is blank.
    """
    if not TRANSFER_LOG_TEMPLATE.exists():
        raise FileNotFoundError(
            f"Missing template file: {TRANSFER_LOG_TEMPLATE.name}. "
            "Upload it to your GitHub repository root folder."
        )

    wb = load_workbook(TRANSFER_LOG_TEMPLATE)
    ws = wb.active

    required_rows = len(transfer_output_df)
    required_end_row = TRANSFER_LOG_START_ROW + max(required_rows, 1) - 1

    # If the template does not have enough rows, extend it and copy row formatting.
    if required_end_row > ws.max_row:
        style_source_row = min(TRANSFER_LOG_START_ROW, ws.max_row)

        for row_idx in range(ws.max_row + 1, required_end_row + 1):
            copy_row_style(ws, style_source_row, row_idx)

    # Clear all body values in A:N so everything else is blank.
    clear_end_row = max(ws.max_row, required_end_row)

    for row in ws.iter_rows(
        min_row=TRANSFER_LOG_START_ROW,
        max_row=clear_end_row,
        min_col=TRANSFER_LOG_FIRST_COL,
        max_col=TRANSFER_LOG_LAST_COL,
    ):
        for cell in row:
            cell.value = None

    # Fill only Trailer Num., manual unloaded pallets, and system highest pallets.
    for offset, (_, row_data) in enumerate(transfer_output_df.iterrows()):
        excel_row = TRANSFER_LOG_START_ROW + offset

        ws[f"{TRANSFER_LOG_TRAILER_COL}{excel_row}"] = row_data["Trailer Last 3"]

        manual_value = row_data.get("Manual Unloaded", None)
        system_value = row_data.get("System Highest", None)

        if pd.notna(manual_value):
            ws[f"{TRANSFER_LOG_MANUAL_UNLOADED_COL}{excel_row}"] = int(manual_value)

        if pd.notna(system_value):
            ws[f"{TRANSFER_LOG_SYSTEM_HIGHEST_COL}{excel_row}"] = int(system_value)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


# -------------------------------------------------------------------
# Inbound system report logic
# -------------------------------------------------------------------
def build_current_report_result(uploaded_file):
    """
    Read the uploaded inbound report and return one row per trailer.

    If the same trailer appears more than once in the report, the LPN count is computed
    from unique LPNs.
    """
    df = pd.read_excel(uploaded_file, header=None, skiprows=HEADER_ROWS)

    # Drop blank/footer rows.
    df = df[df[LPN_COL].notna()].copy()

    df["Trailer"] = df.apply(build_trailer, axis=1)
    df = df[df["Trailer"] != ""].copy()

    # Each LPN = one pallet. nunique guards against duplicate LPN rows.
    result = (
        df.groupby("Trailer")[LPN_COL]
        .nunique()
        .reset_index()
        .rename(columns={LPN_COL: "Pallets"})
        .sort_values("Pallets", ascending=False)
        .reset_index(drop=True)
    )

    result["Trailer Last 3"] = result["Trailer"].apply(get_last_3_numbers)
    result = result[result["Trailer Last 3"] != ""].copy()

    return result


def next_order_for_day(day_records):
    """Return the next order number for the selected day."""
    if not day_records:
        return 1

    return max(int(record.get("Order", 0)) for record in day_records.values()) + 1


def merge_report_into_day_history(day_key, current_result):
    """
    Save current upload into the selected day.

    Logic:
    - Key = Trailer Last 3.
    - If trailer already exists for that selected day, keep only one row.
    - Pallets = the bigger pallet count found so far that day.
    """
    history = st.session_state["daily_history"]
    day_records = history.setdefault(day_key, {})

    # If the same Trailer Last 3 appears multiple times in the same upload,
    # keep the largest pallet count from that upload.
    current_grouped = (
        current_result
        .sort_values("Pallets", ascending=False)
        .groupby("Trailer Last 3", as_index=False)
        .agg(
            Pallets=("Pallets", "max"),
            Full_Trailers=("Trailer", lambda x: sorted(set(str(v) for v in x))),
        )
    )

    for _, row_data in current_grouped.iterrows():
        trailer_last3 = str(row_data["Trailer Last 3"])
        new_pallets = int(row_data["Pallets"])
        new_full_trailers = row_data["Full_Trailers"]

        if trailer_last3 not in day_records:
            day_records[trailer_last3] = {
                "Trailer Last 3": trailer_last3,
                "Pallets": new_pallets,
                "Full Trailers": new_full_trailers,
                "Order": next_order_for_day(day_records),
            }
        else:
            existing = day_records[trailer_last3]

            existing["Pallets"] = max(int(existing.get("Pallets", 0)), new_pallets)

            existing_full_trailers = set(existing.get("Full Trailers", []))
            existing_full_trailers.update(new_full_trailers)
            existing["Full Trailers"] = sorted(existing_full_trailers)

    save_json_file(HISTORY_FILE, history)


def get_day_df(day_key):
    """Return saved daily system trailer data as a DataFrame."""
    records = st.session_state["daily_history"].get(day_key, {})

    rows = []
    for record in records.values():
        rows.append(
            {
                "Order": int(record.get("Order", 0)),
                "Trailer Last 3": str(record.get("Trailer Last 3", "")),
                "System Highest": int(record.get("Pallets", 0)),
                "Full Trailers": ", ".join(record.get("Full Trailers", [])),
            }
        )

    if not rows:
        return pd.DataFrame(columns=["Order", "Trailer Last 3", "System Highest", "Full Trailers"])

    return (
        pd.DataFrame(rows)
        .sort_values("Order")
        .reset_index(drop=True)
    )


# -------------------------------------------------------------------
# Manual unloader log logic
# -------------------------------------------------------------------
def empty_manual_df():
    """Return an empty manual log DataFrame with the expected columns."""
    return pd.DataFrame(
        columns=[
            "Trailer Last 3",
            "Manual Unloaded",
            "Manual Trailer Raw",
            "Manual Duplicate Count",
        ]
    )


def finalize_manual_rows(rows):
    """
    Convert extracted manual rows into the grouped manual-log DataFrame.

    The app groups by trailer last 3 so it can match the system logic.
    If the same last 3 appears more than once in the manual log, the app SUMS
    the manual pallets and shows a warning.
    """
    if not rows:
        return empty_manual_df()

    manual_df = pd.DataFrame(rows)

    grouped = (
        manual_df
        .groupby("Trailer Last 3", as_index=False)
        .agg(
            **{
                "Manual Unloaded": ("Manual Unloaded", "sum"),
                "Manual Trailer Raw": ("Manual Trailer Raw", lambda x: ", ".join(sorted(set(str(v) for v in x)))),
                "Manual Duplicate Count": ("Trailer Last 3", "size"),
            }
        )
        .sort_values("Trailer Last 3")
        .reset_index(drop=True)
    )

    return grouped


def build_manual_log_result_from_spreadsheet(uploaded_manual_file):
    """
    Read manual log from Excel or CSV.

    Expected format:
    - C = Trailer Num.
    - K = Total Pallets UNLOADED

    The app starts reading at row 8, matching the Transfer Log template.
    """
    file_name = uploaded_manual_file.name.lower()

    if file_name.endswith(".csv"):
        raw = pd.read_csv(uploaded_manual_file, header=None)
    else:
        raw = pd.read_excel(uploaded_manual_file, header=None)

    if raw.shape[1] <= max(MANUAL_LOG_TRAILER_COL, MANUAL_LOG_UNLOADED_COL):
        raise ValueError(
            "Manual log does not have enough columns. Expected Trailer Num. in column C "
            "and Total Pallets UNLOADED in column K."
        )

    # Excel row 8 is pandas index 7.
    body = raw.iloc[MANUAL_LOG_START_ROW - 1:].copy()

    rows = []
    for _, row in body.iterrows():
        trailer_value = clean_excel_value(row[MANUAL_LOG_TRAILER_COL])
        manual_unloaded = parse_pallet_count(row[MANUAL_LOG_UNLOADED_COL])

        trailer_last3 = get_last_3_numbers(trailer_value)

        if trailer_last3 == "" or manual_unloaded is None:
            continue

        rows.append(
            {
                "Trailer Last 3": trailer_last3,
                "Manual Unloaded": int(manual_unloaded),
                "Manual Trailer Raw": trailer_value,
            }
        )

    return finalize_manual_rows(rows)


def normalize_pdf_cell(value):
    """Clean a PDF table cell."""
    if value is None:
        return ""

    return str(value).replace("\n", " ").strip()


def find_pdf_table_columns(table):
    """
    Try to find the trailer and manual-unloaded columns from a PDF-extracted table.

    First choice:
    - Detect headers containing Trailer and Unloaded/Pallets.

    Fallback:
    - If the table has enough columns, use the same Excel positions:
      C = index 2, K = index 10.
    """
    for row_idx, row in enumerate(table[:15]):
        cells = [normalize_pdf_cell(cell).lower() for cell in row]
        joined = " ".join(cells)

        trailer_idx = None
        manual_idx = None

        for idx, cell in enumerate(cells):
            if "trailer" in cell and ("num" in cell or "number" in cell or cell.strip() == "trailer"):
                trailer_idx = idx
                break

        for idx, cell in enumerate(cells):
            if "unloaded" in cell:
                manual_idx = idx
                break

            if "pallet" in cell and "unload" in joined:
                manual_idx = idx
                break

        if trailer_idx is not None and manual_idx is not None:
            return trailer_idx, manual_idx, row_idx + 1

    max_cols = max(len(row) for row in table) if table else 0

    if max_cols > max(MANUAL_LOG_TRAILER_COL, MANUAL_LOG_UNLOADED_COL):
        return MANUAL_LOG_TRAILER_COL, MANUAL_LOG_UNLOADED_COL, 0

    return None, None, 0


def extract_manual_rows_from_pdf_table(table):
    """Extract manual trailer rows from one PDF table."""
    if not table:
        return []

    trailer_idx, manual_idx, start_idx = find_pdf_table_columns(table)

    if trailer_idx is None or manual_idx is None:
        return []

    rows = []
    for row in table[start_idx:]:
        if len(row) <= max(trailer_idx, manual_idx):
            continue

        trailer_value = normalize_pdf_cell(row[trailer_idx])
        manual_unloaded = parse_pallet_count(normalize_pdf_cell(row[manual_idx]))
        trailer_last3 = get_last_3_numbers(trailer_value)

        if trailer_last3 == "" or manual_unloaded is None:
            continue

        rows.append(
            {
                "Trailer Last 3": trailer_last3,
                "Manual Unloaded": int(manual_unloaded),
                "Manual Trailer Raw": trailer_value,
            }
        )

    return rows


def extract_manual_rows_from_pdf_text(text):
    """
    Fallback parser for PDFs where table extraction fails.

    This is less reliable than a true table. It looks for lines that contain:
    - a trailer number or trailer last 3
    - a pallet count on the same line

    Best results come from a PDF exported from Excel, not a scanned picture PDF.
    """
    rows = []

    for line in str(text).splitlines():
        line = line.strip()

        if not line:
            continue

        lower_line = line.lower()
        if "trailer" in lower_line or "pallet" in lower_line or "unloaded" in lower_line:
            continue

        # Capture numeric tokens such as 175, 24, 24.0.
        number_matches = list(re.finditer(r"\b\d+(?:\.\d+)?\b", line))

        if len(number_matches) < 2:
            continue

        number_texts = [match.group() for match in number_matches]

        # Trailer candidate: first number with at least 3 digits.
        trailer_candidate = None
        for num in number_texts:
            integer_part = num.split(".")[0]
            if len(integer_part) >= 3:
                trailer_candidate = integer_part
                break

        if trailer_candidate is None:
            continue

        # Pallet candidate: use the last numeric value in the line that is not the trailer.
        pallet_candidate = None
        for num in reversed(number_texts):
            if num == trailer_candidate:
                continue

            parsed = parse_pallet_count(num)

            # Most inbound pallet counts should be a reasonable warehouse pallet count.
            # This filter avoids grabbing dates/page numbers when possible.
            if parsed is not None and 0 <= parsed <= 200:
                pallet_candidate = parsed
                break

        if pallet_candidate is None:
            continue

        trailer_last3 = get_last_3_numbers(trailer_candidate)

        if trailer_last3 == "":
            continue

        rows.append(
            {
                "Trailer Last 3": trailer_last3,
                "Manual Unloaded": int(pallet_candidate),
                "Manual Trailer Raw": trailer_candidate,
            }
        )

    return rows


def build_manual_log_result_from_pdf(uploaded_manual_file):
    """
    Read manual unloader log from a PDF.

    Best PDF format:
    - PDF exported from the same Transfer Log Excel template.
    - Trailer Num. in column C.
    - Total Pallets UNLOADED in column K.

    Note:
    - Text-based/exported PDFs work best.
    - Scanned image PDFs may not work because this app does not perform OCR.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ValueError(
            "PDF support requires pdfplumber. Add 'pdfplumber' to requirements.txt, "
            "commit the change, and redeploy the Streamlit app."
        )

    pdf_bytes = BytesIO(uploaded_manual_file.getvalue())
    rows = []

    with pdfplumber.open(pdf_bytes) as pdf:
        # First try true table extraction.
        for page in pdf.pages:
            tables = page.extract_tables() or []

            for table in tables:
                rows.extend(extract_manual_rows_from_pdf_table(table))

        # If no rows were found, try a text-line fallback.
        if not rows:
            for page in pdf.pages:
                text = page.extract_text() or ""
                rows.extend(extract_manual_rows_from_pdf_text(text))

    if not rows:
        raise ValueError(
            "Could not read trailer/pallet rows from the PDF. Use a text-based PDF exported "
            "from Excel, or upload the manual log as Excel/CSV. Scanned picture PDFs are not supported."
        )

    return finalize_manual_rows(rows)


def build_manual_log_result(uploaded_manual_file):
    """
    Read the manual unloader log.

    Supported formats:
    - Excel: .xlsx, .xlsm
    - CSV: .csv
    - PDF: .pdf

    For Excel/CSV, the app reads:
    - C = Trailer Num.
    - K = Total Pallets UNLOADED
    - starting row = 8

    For PDF, the app tries to extract the same table from a text-based/exported PDF.
    """
    file_name = uploaded_manual_file.name.lower()

    if file_name.endswith(".pdf"):
        return build_manual_log_result_from_pdf(uploaded_manual_file)

    return build_manual_log_result_from_spreadsheet(uploaded_manual_file)


def merge_manual_log_into_history(day_key, manual_result):
    """
    Save the manual unloader log for the selected day.

    Logic:
    - Key = Trailer Last 3.
    - Manual unloaded pallets are replaced by the latest manual upload.
    """
    manual_history = st.session_state["manual_history"]
    day_records = manual_history.setdefault(day_key, {})

    for _, row_data in manual_result.iterrows():
        trailer_last3 = str(row_data["Trailer Last 3"])

        day_records[trailer_last3] = {
            "Trailer Last 3": trailer_last3,
            "Manual Unloaded": int(row_data["Manual Unloaded"]),
            "Manual Trailer Raw": str(row_data.get("Manual Trailer Raw", "")),
            "Manual Duplicate Count": int(row_data.get("Manual Duplicate Count", 1)),
        }

    save_json_file(MANUAL_HISTORY_FILE, manual_history)


def get_manual_df(day_key):
    """Return saved manual unloader data as a DataFrame."""
    records = st.session_state["manual_history"].get(day_key, {})

    rows = []
    for record in records.values():
        rows.append(
            {
                "Trailer Last 3": str(record.get("Trailer Last 3", "")),
                "Manual Unloaded": int(record.get("Manual Unloaded", 0)),
                "Manual Trailer Raw": str(record.get("Manual Trailer Raw", "")),
                "Manual Duplicate Count": int(record.get("Manual Duplicate Count", 1)),
            }
        )

    if not rows:
        return pd.DataFrame(
            columns=["Trailer Last 3", "Manual Unloaded", "Manual Trailer Raw", "Manual Duplicate Count"]
        )

    return (
        pd.DataFrame(rows)
        .sort_values("Trailer Last 3")
        .reset_index(drop=True)
    )


# -------------------------------------------------------------------
# Comparison logic
# -------------------------------------------------------------------
def build_comparison_df(system_df, manual_df):
    """
    Build one comparison table:
    - System Highest = highest count recorded by snapshots
    - Manual Unloaded = number the unloader manually logged
    - Difference = Manual Unloaded - System Highest
    """
    if system_df.empty and manual_df.empty:
        return pd.DataFrame(
            columns=[
                "Order",
                "Trailer Last 3",
                "Manual Unloaded",
                "System Highest",
                "Difference",
                "Match Status",
                "Full Trailers",
                "Manual Trailer Raw",
            ]
        )

    system_small = system_df.copy()
    manual_small = manual_df.copy()

    if system_small.empty:
        system_small = pd.DataFrame(columns=["Order", "Trailer Last 3", "System Highest", "Full Trailers"])

    if manual_small.empty:
        manual_small = pd.DataFrame(columns=["Trailer Last 3", "Manual Unloaded", "Manual Trailer Raw", "Manual Duplicate Count"])

    comparison = pd.merge(
        system_small,
        manual_small,
        on="Trailer Last 3",
        how="outer",
    )

    # Keep manual-log rows first when manual exists, otherwise system order.
    comparison["Sort Order"] = comparison["Order"].fillna(999999).astype(int)
    comparison = comparison.sort_values(["Sort Order", "Trailer Last 3"]).reset_index(drop=True)

    def calc_difference(row):
        manual = row.get("Manual Unloaded")
        system = row.get("System Highest")

        if pd.isna(manual) or pd.isna(system):
            return None

        return int(manual) - int(system)

    def calc_status(row):
        manual = row.get("Manual Unloaded")
        system = row.get("System Highest")

        if pd.isna(manual) and pd.isna(system):
            return "No data"

        if pd.isna(manual):
            return "Missing manual count"

        if pd.isna(system):
            return "Manual only - not seen by system"

        diff = int(manual) - int(system)

        if diff == 0:
            return "Matched"

        return "Mismatch"

    comparison["Difference"] = comparison.apply(calc_difference, axis=1)
    comparison["Match Status"] = comparison.apply(calc_status, axis=1)

    wanted_cols = [
        "Order",
        "Trailer Last 3",
        "Manual Unloaded",
        "System Highest",
        "Difference",
        "Match Status",
        "Full Trailers",
        "Manual Trailer Raw",
        "Manual Duplicate Count",
    ]

    for col in wanted_cols:
        if col not in comparison.columns:
            comparison[col] = None

    return comparison[wanted_cols]


def build_transfer_output(comparison_df):
    """
    Build rows for the Excel Transfer Log.

    If a manual log has been uploaded, include:
    - all manual rows
    - plus any system non-research rows missing from the manual log

    If no manual log exists yet, include only system non-research rows.
    """
    if comparison_df.empty:
        return pd.DataFrame(columns=["Trailer Last 3", "Manual Unloaded", "System Highest"])

    has_any_manual = comparison_df["Manual Unloaded"].notna().any()

    if has_any_manual:
        output = comparison_df[
            comparison_df["Manual Unloaded"].notna()
            | (comparison_df["System Highest"].fillna(0) > THRESHOLD)
        ].copy()
    else:
        output = comparison_df[
            comparison_df["System Highest"].fillna(0) > THRESHOLD
        ].copy()

    output = output[["Trailer Last 3", "Manual Unloaded", "System Highest"]].copy()

    return output.reset_index(drop=True)


# -------------------------------------------------------------------
# App state
# -------------------------------------------------------------------
if "daily_history" not in st.session_state:
    st.session_state["daily_history"] = load_json_file(HISTORY_FILE)

if "manual_history" not in st.session_state:
    st.session_state["manual_history"] = load_json_file(MANUAL_HISTORY_FILE)


selected_day = st.date_input("Select day", value=date.today())
day_key = selected_day.strftime("%Y-%m-%d")

col_a, col_b, col_c = st.columns([3, 3, 1.3])

with col_a:
    uploaded = st.file_uploader(
        "Upload system inbound report",
        type=["xlsx", "xlsm"],
        key="system_inbound_report",
    )

with col_b:
    manual_uploaded = st.file_uploader(
        "Upload manual unloader log at end of day",
        type=["xlsx", "xlsm", "csv", "pdf"],
        key="manual_unloader_log",
    )

with col_c:
    st.write("")
    st.write("")
    clear_day = st.button("Clear selected day")

if clear_day:
    st.session_state["daily_history"].pop(day_key, None)
    st.session_state["manual_history"].pop(day_key, None)
    save_json_file(HISTORY_FILE, st.session_state["daily_history"])
    save_json_file(MANUAL_HISTORY_FILE, st.session_state["manual_history"])
    st.success(f"Cleared saved system and manual loads for {day_key}.")
    st.rerun()


if uploaded is not None:
    try:
        current_result = build_current_report_result(uploaded)
        merge_report_into_day_history(day_key, current_result)

        st.success(
            f"System report processed for {day_key}. "
            "If a trailer was already saved for this day, the app kept the bigger pallet count."
        )

    except Exception as e:
        st.error(f"Could not process the uploaded system report: {e}")


if manual_uploaded is not None:
    try:
        manual_result = build_manual_log_result(manual_uploaded)
        merge_manual_log_into_history(day_key, manual_result)

        st.success(
            f"Manual unloader log processed for {day_key}. "
            "Manual counts were matched by trailer last 3."
        )

        duplicate_manual = manual_result[manual_result["Manual Duplicate Count"] > 1]
        if not duplicate_manual.empty:
            st.warning(
                "The manual log has duplicate trailer last-3 numbers. The app summed the manual pallets "
                "for those duplicate last-3 numbers. Please review before using the final log."
            )

    except Exception as e:
        st.error(f"Could not process the manual unloader log: {e}")


system_df = get_day_df(day_key)
manual_df = get_manual_df(day_key)
comparison_df = build_comparison_df(system_df, manual_df)
transfer_output = build_transfer_output(comparison_df)

# Research rows based on system highest <= threshold.
research_display = comparison_df[
    comparison_df["System Highest"].notna()
    & (comparison_df["System Highest"] <= THRESHOLD)
].copy()

if not research_display.empty:
    research_display = research_display[
        [
            "Trailer Last 3",
            "Manual Unloaded",
            "System Highest",
            "Difference",
            "Match Status",
            "Full Trailers",
        ]
    ].copy()
    research_display["Status"] = "research"


# -------------------------------------------------------------------
# Metrics
# -------------------------------------------------------------------
matched_count = len(comparison_df[comparison_df["Match Status"] == "Matched"]) if not comparison_df.empty else 0
mismatch_count = len(comparison_df[comparison_df["Match Status"] == "Mismatch"]) if not comparison_df.empty else 0

c1, c2, c3, c4 = st.columns(4)
c1.metric("System trailers saved", len(system_df))
c2.metric("Manual trailers saved", len(manual_df))
c3.metric("Matched exactly", matched_count)
c4.metric("Mismatches", mismatch_count)


# -------------------------------------------------------------------
# Transfer Log / Comparison
# -------------------------------------------------------------------
st.subheader("📋 Transfer Log comparison")

if transfer_output.empty:
    st.info("No Transfer Log rows for this selected day yet.")
else:
    st.write(
        "The Excel log fills only **Trailer Num.**, **Total Pallets UNLOADED** "
        "and **Total Pallets in this Trailer**. In this version, **Total Pallets UNLOADED** "
        "comes from the manual unloader log, and **Total Pallets in this Trailer** comes from "
        "the highest system count recorded during the day."
    )

    display_cols = [
        "Trailer Last 3",
        "Manual Unloaded",
        "System Highest",
        "Difference",
        "Match Status",
        "Full Trailers",
        "Manual Trailer Raw",
    ]

    display_df = comparison_df[
        comparison_df["Trailer Last 3"].isin(transfer_output["Trailer Last 3"])
    ][display_cols].copy()

    styled_comparison = display_df.style.apply(flag_variance, axis=1).hide(axis="index")

    st.dataframe(styled_comparison, use_container_width=True)

    try:
        transfer_log_file = create_transfer_log_excel(transfer_output)

        st.download_button(
            "⬇️ Download completed Transfer Log comparison",
            data=transfer_log_file,
            file_name=f"completed_transfer_log_comparison_{day_key}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    except FileNotFoundError as e:
        st.error(str(e))
        st.info(
            "Fix: upload 'Transfer Log New 8-2025.xlsx' to your GitHub repo root folder. "
            "If this page is inside /pages, the code already looks one folder up."
        )


# -------------------------------------------------------------------
# Research Loads
# -------------------------------------------------------------------
st.subheader(f"🚨 Research loads with system highest of {THRESHOLD} or fewer pallets")

if research_display.empty:
    st.success("No short system loads saved for this selected day.")
else:
    styled_research = research_display.style.apply(flag_red, axis=1).hide(axis="index")

    st.dataframe(styled_research, use_container_width=True)

    st.download_button(
        "⬇️ Download research list (CSV)",
        data=research_display.to_csv(index=False).encode("utf-8"),
        file_name=f"loads_to_research_{day_key}.csv",
        mime="text/csv",
    )
