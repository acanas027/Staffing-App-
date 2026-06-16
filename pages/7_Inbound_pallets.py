import re
from copy import copy
from datetime import date
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="Inbound Pallets", layout="wide")

st.title("Pallets per Trailer")

st.write(
    "Upload your inbound report. Each LPN counts as one pallet. "
    "This version does not save previous uploads. Every upload is treated as a fresh run."
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


def create_transfer_log_excel(transfer_ready_df):
    """
    Create a completed Transfer Log from the repository template.

    Fills only:
    - C: Trailer Num.
    - L: Total Pallets in this Trailer

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


def create_research_excel(research_df):
    """Create an Excel file for the research list."""
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        research_df.to_excel(writer, index=False, sheet_name="Research Loads")

    output.seek(0)
    return output.getvalue()


# -------------------------------------------------------------------
# Inbound report logic
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


def filter_by_trailer(df, search_text):
    """Filter preview table by trailer last 3 or full trailer number."""
    if df.empty:
        return df

    search_text = str(search_text).strip()

    if search_text == "":
        return df

    search_digits = "".join(re.findall(r"\d", search_text))
    search_value = search_digits if search_digits else search_text.lower()

    if "Full Trailers" in df.columns:
        mask = (
            df["Trailer Last 3"].astype(str).str.contains(search_value, case=False, na=False)
            | df["Full Trailers"].astype(str).str.contains(search_value, case=False, na=False)
        )
    else:
        mask = df["Trailer Last 3"].astype(str).str.contains(search_value, case=False, na=False)

    return df[mask].copy()


# -------------------------------------------------------------------
# App
# -------------------------------------------------------------------
selected_day = st.date_input("Select day", value=date.today())
day_key = selected_day.strftime("%Y-%m-%d")

uploaded = st.file_uploader("Upload your inbound report", type=["xlsx", "xlsm"])

if uploaded is not None:
    try:
        current_result = build_current_report_result(uploaded)

        # If the same Trailer Last 3 appears multiple times in the uploaded report,
        # keep the largest pallet count from that upload.
        day_df = (
            current_result
            .sort_values("Pallets", ascending=False)
            .groupby("Trailer Last 3", as_index=False)
            .agg(
                Pallets=("Pallets", "max"),
                Full_Trailers=("Trailer", lambda x: ", ".join(sorted(set(str(v) for v in x)))),
            )
            .sort_values("Pallets", ascending=False)
            .reset_index(drop=True)
        )

        transfer_ready = day_df[day_df["Pallets"] > THRESHOLD].copy()
        research_ready = day_df[day_df["Pallets"] <= THRESHOLD].copy()

        transfer_log_output = transfer_ready[["Trailer Last 3", "Pallets"]].copy()

        research_display = research_ready[["Trailer Last 3", "Pallets", "Full_Trailers"]].copy()
        research_display = research_display.rename(columns={"Full_Trailers": "Full Trailers"})

        if not research_display.empty:
            research_display["Status"] = "research"

        transfer_preview = transfer_ready[["Trailer Last 3", "Pallets", "Full_Trailers"]].copy()
        transfer_preview = transfer_preview.rename(columns={"Full_Trailers": "Full Trailers"})

        c1, c2, c3 = st.columns(3)
        c1.metric("Trailers in current upload", len(day_df))
        c2.metric(f"Transfer Log rows over {THRESHOLD}", len(transfer_ready))
        c3.metric(f"Research rows {THRESHOLD} or less", len(research_ready))

        # -------------------------------------------------------------------
        # Transfer Log
        # -------------------------------------------------------------------
        st.subheader("Transfer Log")

        if transfer_log_output.empty:
            st.info("No non-research loads found in this upload.")
        else:
            duplicate_last3 = transfer_log_output[
                transfer_log_output["Trailer Last 3"].duplicated(keep=False)
            ]

            if not duplicate_last3.empty:
                st.warning(
                    "Warning: at least two trailers have the same last 3 digits. "
                    "Review the Transfer Log before using it."
                )

            st.write(
                "This fills only Trailer Num. and Total Pallets in this Trailer "
                "in the Excel template, using the highest pallet count found in the uploaded report."
            )

            search_col, download_col = st.columns([2, 1])

            with search_col:
                transfer_search = st.text_input(
                    "Search trailer in preview",
                    placeholder="Type trailer last 3 or full trailer number",
                    key="transfer_search",
                )

            with download_col:
                try:
                    transfer_log_file = create_transfer_log_excel(transfer_log_output)

                    st.write("")
                    st.download_button(
                        "Download completed Transfer Log",
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

            filtered_transfer_preview = filter_by_trailer(transfer_preview, transfer_search)

            st.dataframe(
                filtered_transfer_preview,
                use_container_width=True,
                hide_index=True,
            )

        # -------------------------------------------------------------------
        # Research Loads
        # -------------------------------------------------------------------
        st.subheader(f"Research loads with {THRESHOLD} or fewer pallets")

        if research_display.empty:
            st.success("No short loads found in this upload.")
        else:
            research_search_col, research_download_col = st.columns([2, 1])

            with research_search_col:
                research_search = st.text_input(
                    "Search research trailer",
                    placeholder="Type trailer last 3 or full trailer number",
                    key="research_search",
                )

            with research_download_col:
                research_excel_file = create_research_excel(research_display)

                st.write("")
                st.download_button(
                    "Download research list",
                    data=research_excel_file,
                    file_name=f"loads_to_research_{day_key}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            filtered_research_display = filter_by_trailer(research_display, research_search)
            styled = filtered_research_display.style.apply(flag_red, axis=1).hide(axis="index")

            st.dataframe(styled, use_container_width=True)

    except Exception as e:
        st.error(f"Could not process the uploaded report: {e}")

else:
    st.info("Waiting for an inbound report.")
