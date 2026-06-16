import re
from copy import copy
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from openpyxl import load_workbook


st.set_page_config(page_title="Inbound Pallets", layout="wide")

st.title("📦 Pallets per Trailer")

st.write(
    "Upload your inbound report. Each **LPN** counts as one pallet, and the "
    "**trailer number** is columns C, D, E, F and G combined."
)

# -------------------------------------------------------------------
# Transfer Log template
# Put this Excel file in your GitHub repository, in the same folder as app.py.
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
            "Upload it to your GitHub repository in the same folder as app.py."
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


uploaded = st.file_uploader("Upload your inbound report", type=["xlsx", "xlsm"])

if uploaded is not None:
    # Headers span the first 3 rows, so read with no header and skip them.
    df = pd.read_excel(uploaded, header=None, skiprows=HEADER_ROWS)

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

    high = result[result["Pallets"] > THRESHOLD].reset_index(drop=True)
    low = result[result["Pallets"] <= THRESHOLD].reset_index(drop=True)

    c1, c2, c3 = st.columns(3)

    c1.metric("Total trailers", len(result))
    c2.metric(f"Loads over {THRESHOLD}", len(high))
    c3.metric(f"Loads {THRESHOLD} or less", len(low))

    # ---- List 1: more than 9 pallets ----
    st.subheader(f"✅ Loads with more than {THRESHOLD} pallets")

    if high.empty:
        st.warning("No loads are over the research threshold.")
    else:
        st.dataframe(high, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Download full loads (CSV)",
            data=high.to_csv(index=False).encode("utf-8"),
            file_name="loads_over_9.csv",
            mime="text/csv",
        )

    # ---- Transfer Log output ----
    st.subheader("📋 Transfer Log Excel")

    transfer_ready = high[["Trailer Last 3", "Pallets"]].copy()

    if transfer_ready.empty:
        st.info("No non-research loads to send to the Transfer Log.")
    else:
        duplicate_last3 = transfer_ready[
            transfer_ready["Trailer Last 3"].duplicated(keep=False)
        ]

        if not duplicate_last3.empty:
            st.warning(
                "Warning: at least two trailers have the same last 3 digits. "
                "Review the Transfer Log before using it."
            )

        st.write(
            "The Transfer Log will fill only **Trailer Num.** and "
            "**Total Pallets in this Trailer**. Everything else in the body rows stays blank."
        )

        st.dataframe(transfer_ready, use_container_width=True, hide_index=True)

        try:
            transfer_log_file = create_transfer_log_excel(transfer_ready)

            st.download_button(
                "⬇️ Download completed Transfer Log",
                data=transfer_log_file,
                file_name="completed_transfer_log.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except FileNotFoundError as e:
            st.error(str(e))
            st.info(
                "Fix: upload 'Transfer Log New 8-2025.xlsx' to your GitHub repo "
                "in the same folder as this app.py file."
            )

    # ---- List 2: 9 or fewer pallets, flagged for research ----
    st.subheader(f"🚨 Loads with {THRESHOLD} or fewer pallets")

    if low.empty:
        st.success("No short loads — nothing to research.")
    else:
        low_display = low.copy()
        low_display["Status"] = "research"

        styled = low_display.style.apply(flag_red, axis=1).hide(axis="index")

        st.dataframe(styled, use_container_width=True)

        st.download_button(
            "⬇️ Download research list (CSV)",
            data=low_display.to_csv(index=False).encode("utf-8"),
            file_name="loads_to_research.csv",
            mime="text/csv",
        )

else:
    st.info("Waiting for a file...")
