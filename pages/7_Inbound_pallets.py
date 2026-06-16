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
    "Upload your inbound report throughout the day. Each **LPN** counts as one pallet. "
    "The app saves one row per trailer number for the selected day and keeps the "
    "**largest pallet count found** for that trailer."
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

# Transfer Log columns based on the template:
# C = Trailer Num.
# L = Total Pallets in this Trailer
TRANSFER_LOG_START_ROW = 8
TRANSFER_LOG_TRAILER_COL = "C"
TRANSFER_LOG_PALLETS_COL = "L"

# Clear body columns A:N so everything else stays blank.
TRANSFER_LOG_FIRST_COL = 1
TRANSFER_LOG_LAST_COL = 14

# Inbound report column positions, 0-indexed:
# C,D,E,F,G = 2,3,4,5,6
# LPN # = 7, column H
TRAILER_COLS = [2, 3, 4, 5, 6]
LPN_COL = 7

HEADER_ROWS = 3       # first 3 rows are headers
THRESHOLD = 9         # loads with 9 or fewer pallets get flagged for research


def load_history():
    """Load saved daily trailer history from the app runtime."""
    if HISTORY_FILE.exists():
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    return {}


def save_history(history):
    """Save daily trailer history to the app runtime."""
    try:
        with open(HISTORY_FILE, "w", encoding="utf-8") as f:
            json.dump(history, f, indent=2)
    except Exception:
        pass


def clean_excel_value(value):
    """Clean Excel values so numbers like 123.0 become 123."""
    if pd.isna(value):
        return ""

    if isinstance(value, float) and value.is_integer():
        return str(int(value)).strip()

    return str(value).strip()


def build_trailer(row):
    """Concatenate trailer parts from columns C through G."""
    return "".join(clean_excel_value(row[c]) for c in TRAILER_COLS)


def get_last_3_numbers(trailer):
    """Return the last 3 numeric digits from the trailer value."""
    digits = "".join(re.findall(r"\d", str(trailer)))

    if not digits:
        return ""

    return digits[-3:].zfill(3)


def flag_red(row):
    return ["background-color: #ffb3b3; color: #800000; font-weight: bold"] * len(row)


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


def create_transfer_log_excel(transfer_ready_df):
    """
    Create a completed Transfer Log from the repository template.

    Fills only:
    - Trailer Num. column
    - Total Pallets in this Trailer column

    Everything else in the body rows is blank.
    """
    if not TRANSFER_LOG_TEMPLATE.exists():
        raise FileNotFoundError(
            f"Missing template file: {TRANSFER_LOG_TEMPLATE.name}. "
            "Upload it to your GitHub repository root folder."
        )

    wb = load_workbook(TRANSFER_LOG_TEMPLATE)
    ws = wb.active

    required_rows = len(transfer_ready_df)
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

    # Fill only Trailer Num. and Total Pallets in this Trailer.
    for offset, (_, row_data) in enumerate(transfer_ready_df.iterrows()):
        excel_row = TRANSFER_LOG_START_ROW + offset
        ws[f"{TRANSFER_LOG_TRAILER_COL}{excel_row}"] = row_data["Trailer Last 3"]
        ws[f"{TRANSFER_LOG_PALLETS_COL}{excel_row}"] = int(row_data["Pallets"])

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    return output.getvalue()


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

    save_history(history)


def get_day_df(day_key):
    """Return saved daily trailer data as a DataFrame."""
    records = st.session_state["daily_history"].get(day_key, {})

    rows = []
    for record in records.values():
        rows.append(
            {
                "Order": int(record.get("Order", 0)),
                "Trailer Last 3": str(record.get("Trailer Last 3", "")),
                "Pallets": int(record.get("Pallets", 0)),
                "Full Trailers": ", ".join(record.get("Full Trailers", [])),
            }
        )

    if not rows:
        return pd.DataFrame(columns=["Order", "Trailer Last 3", "Pallets", "Full Trailers"])

    return (
        pd.DataFrame(rows)
        .sort_values("Order")
        .reset_index(drop=True)
    )


# -------------------------------------------------------------------
# App state
# -------------------------------------------------------------------
if "daily_history" not in st.session_state:
    st.session_state["daily_history"] = load_history()


selected_day = st.date_input("Select day", value=date.today())
day_key = selected_day.strftime("%Y-%m-%d")

col_a, col_b = st.columns([3, 1])

with col_a:
    uploaded = st.file_uploader("Upload your inbound report", type=["xlsx", "xlsm"])

with col_b:
    st.write("")
    st.write("")
    clear_day = st.button("Clear selected day")

if clear_day:
    st.session_state["daily_history"].pop(day_key, None)
    save_history(st.session_state["daily_history"])
    st.success(f"Cleared saved loads for {day_key}.")
    st.rerun()


if uploaded is not None:
    try:
        current_result = build_current_report_result(uploaded)
        merge_report_into_day_history(day_key, current_result)

        st.success(
            f"Upload processed for {day_key}. "
            "If a trailer was already saved for this day, the app kept the bigger pallet count."
        )

    except Exception as e:
        st.error(f"Could not process the uploaded report: {e}")


day_df = get_day_df(day_key)

transfer_ready = day_df[day_df["Pallets"] > THRESHOLD].copy()
research_ready = day_df[day_df["Pallets"] <= THRESHOLD].copy()

# Only the Transfer Log output needs these two columns.
transfer_log_output = transfer_ready[["Trailer Last 3", "Pallets"]].copy()

# Research display keeps the full trailer info so you can investigate.
research_display = research_ready[["Trailer Last 3", "Pallets", "Full Trailers"]].copy()
if not research_display.empty:
    research_display["Status"] = "research"

c1, c2, c3 = st.columns(3)
c1.metric("Saved trailers for selected day", len(day_df))
c2.metric(f"Transfer Log rows over {THRESHOLD}", len(transfer_ready))
c3.metric(f"Research rows {THRESHOLD} or less", len(research_ready))

# -------------------------------------------------------------------
# Transfer Log
# -------------------------------------------------------------------
st.subheader("📋 Transfer Log")

if transfer_log_output.empty:
    st.info("No non-research loads saved for this selected day yet.")
else:
    duplicate_last3 = transfer_log_output[
        transfer_log_output["Trailer Last 3"].duplicated(keep=False)
    ]

    if not duplicate_last3.empty:
        st.warning(
            "Warning: at least two saved trailers have the same last 3 digits. "
            "Review the Transfer Log before using it."
        )

    st.write(
        "This is the saved daily log. It fills only **Trailer Num.** and "
        "**Total Pallets in this Trailer** in the Excel template."
    )

    st.dataframe(transfer_log_output, use_container_width=True, hide_index=True)

    try:
        transfer_log_file = create_transfer_log_excel(transfer_log_output)

        st.download_button(
            "⬇️ Download completed Transfer Log",
            data=transfer_log_file,
            file_name=f"completed_transfer_log_{day_key}.xlsx",
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
st.subheader(f"🚨 Research loads with {THRESHOLD} or fewer pallets")

if research_display.empty:
    st.success("No short loads saved for this selected day.")
else:
    styled = research_display.style.apply(flag_red, axis=1).hide(axis="index")

    st.dataframe(styled, use_container_width=True)

    st.download_button(
        "⬇️ Download research list (CSV)",
        data=research_display.to_csv(index=False).encode("utf-8"),
        file_name=f"loads_to_research_{day_key}.csv",
        mime="text/csv",
    )
