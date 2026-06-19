 -*- coding: utf-8 -*-
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st


st.set_page_config(page_title="SKU Data + WMS Matcher", layout="wide")
st.title("SKU Data + WMS Matcher")
st.caption(
    "Upload the same two Excel files: the Data/orders file and the big qPORT/WMS file. "
    "The app returns one Excel showing only SKUs from the Data file, matched to all available WMS rows."
)

with st.expander("What this new app does"):
    st.markdown(
        "- Normalizes SKUs before matching, so values like `06795.48940`, `6795.48940`, and `6795.4894` match.\n"
        "- Uses **only the SKUs from the Data/orders file** as the driver list.\n"
        "- Pulls matching rows from the big qPORT/WMS file.\n"
        "- Normalizes current and previous locations into compact readable location fields.\n"
        "- Does **not** allocate inventory and does **not** create shortages. This is only a data visibility/matching tool.\n"
        "- The downloaded Excel includes summary, matched WMS detail, original Data lines, and unmatched Data SKUs."
    )


with st.sidebar:
    st.header("Upload Files")
    data_file = st.file_uploader("Data / orders file", type=["xlsx"], key="data_file")
    wms_file = st.file_uploader("Big qPORT / WMS file", type=["xlsx"], key="wms_file")

    st.divider()
    st.subheader("Optional filters")
    exclude_s_items = st.checkbox("Exclude Items starting with S", value=True)
    decimal_only = st.checkbox("Only include decimal SKUs", value=True)

    st.divider()
    run_btn = st.button("Build Matched Excel", type="primary", use_container_width=True)


# ---------- basic cleaners ----------

def clean_str(x):
    if pd.isna(x):
        return ""
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x).strip()


def clean_order_value(x):
    """Display orders/statuses cleanly instead of 3000129000.0."""
    if pd.isna(x):
        return ""
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x).strip()


def h_to_str(x):
    if pd.isna(x):
        return None
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x).strip()


def i_to_str(x):
    if pd.isna(x):
        return "000"
    if isinstance(x, float) and x == int(x):
        return str(int(x)).zfill(3)
    return str(x).strip().zfill(3)


# ---------- SKU normalization ----------

def normalize_sku(x):
    """Return a consistent SKU string for matching.

    Examples:
    - 06795.48940 -> 6795.4894
    - 6795.48940  -> 6795.4894
    - 6795.4894   -> 6795.4894
    - 79341.0     -> 79341
    """
    if pd.isna(x):
        return ""

    s = str(x).strip()
    if not s or s.lower() == "nan":
        return ""

    # Remove commas that sometimes appear in exported numbers.
    s = s.replace(",", "")

    # Handle whole-number float artifact.
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]

    if "." in s:
        whole, dec = s.split(".", 1)
        whole = whole.strip()
        dec = dec.strip()

        if whole.isdigit():
            whole = str(int(whole))

        # Excel often drops trailing decimal zeros on one file but not the other.
        dec = dec.rstrip("0")
        return f"{whole}.{dec}" if dec else whole

    return s


def has_decimal_sku(x):
    return "." in normalize_sku(x)


def build_full_sku(h_val, i_val):
    """Build full SKU from qPORT/WMS columns H and I, same logic as your shortage app."""
    if h_val is None:
        return None

    if "." in h_val:
        whole, dec = h_val.split(".", 1)
        whole = whole.zfill(5)
        dec = dec.ljust(2, "0")[:2]
        return f"{whole}.{dec}{i_val}"

    return h_val.zfill(5)


# ---------- location normalization ----------

def normalize_location_text(x):
    """Compact location text with spaces removed and uppercase applied."""
    s = clean_str(x).upper().replace(" ", "")
    if not s or s == "NAN":
        return ""

    # qPORT often gives BLACKHOL because the location is split across six columns.
    # Keep it readable in the output.
    if s in {"BLACKHOL", "BLACKHO", "BLACKH", "BLACK"}:
        return "BLACKHOLE"

    return s


def build_location_from_parts(df, start_col, end_col):
    """Join qPORT location columns into one compact location string."""
    joined = df.iloc[:, start_col:end_col].apply(
        lambda col: col.map(clean_str)
    ).agg("".join, axis=1)
    return joined.map(normalize_location_text)


def parse_location(loc):
    """Split compact location into simple useful pieces.

    Example: RC2A30X8 -> zone RC2, aisle A, bay 30, position X8, normalized RC2-A-30-X8.
    If the structure is unusual, the compact value is still preserved.
    """
    compact = normalize_location_text(loc)
    if not compact:
        return pd.Series({
            "Location Normalized": "",
            "Location Zone": "",
            "Location Aisle": "",
            "Location Bay": "",
            "Location Position": "",
        })

    if compact == "BLACKHOLE":
        return pd.Series({
            "Location Normalized": "BLACKHOLE",
            "Location Zone": "BLACKHOLE",
            "Location Aisle": "",
            "Location Bay": "",
            "Location Position": "",
        })

    zone = compact[:3] if len(compact) >= 3 else compact
    remainder = compact[3:] if len(compact) > 3 else ""

    aisle = ""
    if remainder and remainder[0].isalpha():
        aisle = remainder[0]
        remainder = remainder[1:]

    bay = ""
    position = ""
    m = re.match(r"^(\d+)(.*)$", remainder)
    if m:
        bay = m.group(1)
        position = m.group(2)
    else:
        position = remainder

    pieces = [p for p in [zone, aisle, bay, position] if p]
    normalized = "-".join(pieces) if pieces else compact

    return pd.Series({
        "Location Normalized": normalized,
        "Location Zone": zone,
        "Location Aisle": aisle,
        "Location Bay": bay,
        "Location Position": position,
    })


def add_location_fields(df, source_col, prefix):
    parsed = df[source_col].apply(parse_location)
    parsed = parsed.rename(columns={
        "Location Normalized": f"{prefix} Location Normalized",
        "Location Zone": f"{prefix} Location Zone",
        "Location Aisle": f"{prefix} Location Aisle",
        "Location Bay": f"{prefix} Location Bay",
        "Location Position": f"{prefix} Location Position",
    })
    return pd.concat([df, parsed], axis=1)


# ---------- loaders ----------

def load_data_file(file, exclude_s=True, decimals_only=True):
    df = pd.read_excel(file, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["Item", "Target Date", "Delivery Date", "Quantity Ordered"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(
            "Missing expected column(s) in Data/orders file: " + ", ".join(missing)
        )

    df = df[df["Item"].notna()].copy()
    df["Data Item"] = df["Item"].map(clean_str)
    df["Normalized SKU"] = df["Data Item"].map(normalize_sku)

    if exclude_s:
        df = df[~df["Data Item"].str.upper().str.startswith("S")].copy()

    if decimals_only:
        df = df[df["Normalized SKU"].str.contains(r"\.", regex=True, na=False)].copy()

    df = df[df["Normalized SKU"] != ""].copy()

    df["Target Date"] = pd.to_datetime(df["Target Date"], errors="coerce")
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"], errors="coerce")
    df["Quantity Ordered"] = pd.to_numeric(
        df["Quantity Ordered"], errors="coerce"
    ).fillna(0)

    if "Order" in df.columns:
        df["Order"] = df["Order"].map(clean_order_value)
    else:
        df["Order"] = ""

    if "Order Status" in df.columns:
        df["Order Status"] = df["Order Status"].map(clean_order_value)
    else:
        df["Order Status"] = ""

    df = df.reset_index(drop=True)
    df["Data Row #"] = df.index + 2  # approximate Excel row because row 1 is header

    return df


def load_wms_file(file):
    # qPORT/WMS file has useful rows starting after the first 3 header rows.
    raw = pd.read_excel(file, sheet_name=0, header=None, skiprows=3)

    current_location = build_location_from_parts(raw, 0, 6)
    previous_location = build_location_from_parts(raw, 10, 16)

    lpn = raw.iloc[:, 6].map(clean_order_value)
    sku_h = raw.iloc[:, 7].apply(h_to_str)
    sku_i = raw.iloc[:, 8].apply(i_to_str)
    full_sku = [build_full_sku(h, i) for h, i in zip(sku_h, sku_i)]
    quantity = pd.to_numeric(raw.iloc[:, 9], errors="coerce").fillna(0)
    priority_date_raw = raw.iloc[:, 27]

    out = pd.DataFrame({
        "WMS Row #": raw.index + 4,  # because we skipped 3 rows; Excel rows start at 1
        "WMS SKU Number": full_sku,
        "Normalized SKU": [normalize_sku(x) for x in full_sku],
        "LPN #": lpn,
        "WMS Quantity": quantity,
        "Current Location Compact": current_location,
        "Previous Location Compact": previous_location,
        "Consumer Priority Date": pd.to_datetime(
            priority_date_raw.astype(str), format="%Y%m%d", errors="coerce"
        ),
    })

    out = out[out["Normalized SKU"] != ""].copy()
    out = add_location_fields(out, "Current Location Compact", "Current")
    out = add_location_fields(out, "Previous Location Compact", "Previous")
    out = out.reset_index(drop=True)
    return out


# ---------- grouping helpers ----------

def unique_join(values, limit=50):
    vals = []
    for v in values:
        if pd.isna(v):
            continue
        s = str(v).strip()
        if not s or s.lower() in {"nan", "nat"}:
            continue
        vals.append(s)

    unique_vals = sorted(set(vals))
    if len(unique_vals) > limit:
        shown = unique_vals[:limit]
        return ", ".join(shown) + f" ... (+{len(unique_vals) - limit} more)"
    return ", ".join(unique_vals)


def unique_date_join(values, limit=30):
    dates = pd.to_datetime(pd.Series(values), errors="coerce").dropna().dt.date
    return unique_join(dates.astype(str), limit=limit)


def build_data_summary(data_df):
    rows = []
    for sku, g in data_df.groupby("Normalized SKU", dropna=False):
        rows.append({
            "Normalized SKU": sku,
            "Data Item Values": unique_join(g["Data Item"]),
            "Data Lines": len(g),
            "Order Count": g["Order"].replace("", pd.NA).dropna().nunique(),
            "Orders": unique_join(g["Order"]),
            "Order Statuses": unique_join(g["Order Status"]),
            "Total Quantity Ordered": g["Quantity Ordered"].sum(),
            "Earliest Customer Target Date": pd.to_datetime(g["Target Date"], errors="coerce").min(),
            "Latest Customer Target Date": pd.to_datetime(g["Target Date"], errors="coerce").max(),
            "Customer Target Dates": unique_date_join(g["Target Date"]),
            "Earliest Delivery Date": pd.to_datetime(g["Delivery Date"], errors="coerce").min(),
            "Latest Delivery Date": pd.to_datetime(g["Delivery Date"], errors="coerce").max(),
            "Delivery Dates": unique_date_join(g["Delivery Date"]),
        })
    return pd.DataFrame(rows).sort_values("Normalized SKU").reset_index(drop=True)


def build_wms_summary(wms_filtered_df):
    columns = [
        "Normalized SKU", "Matched WMS Rows", "Total WMS Quantity",
        "Earliest Consumer Priority Date", "Latest Consumer Priority Date",
        "Consumer Priority Dates", "Current Location Count", "Current Locations",
        "Current Zones", "Previous Locations", "LPN Count"
    ]
    if wms_filtered_df.empty:
        return pd.DataFrame(columns=columns)

    rows = []
    for sku, g in wms_filtered_df.groupby("Normalized SKU", dropna=False):
        rows.append({
            "Normalized SKU": sku,
            "Matched WMS Rows": len(g),
            "Total WMS Quantity": g["WMS Quantity"].sum(),
            "Earliest Consumer Priority Date": pd.to_datetime(g["Consumer Priority Date"], errors="coerce").min(),
            "Latest Consumer Priority Date": pd.to_datetime(g["Consumer Priority Date"], errors="coerce").max(),
            "Consumer Priority Dates": unique_date_join(g["Consumer Priority Date"]),
            "Current Location Count": g["Current Location Normalized"].replace("", pd.NA).dropna().nunique(),
            "Current Locations": unique_join(g["Current Location Normalized"]),
            "Current Zones": unique_join(g["Current Location Zone"]),
            "Previous Locations": unique_join(g["Previous Location Normalized"]),
            "LPN Count": g["LPN #"].replace("", pd.NA).dropna().nunique(),
        })
    return pd.DataFrame(rows, columns=columns).sort_values("Normalized SKU").reset_index(drop=True)


def build_outputs(data_df, wms_df):
    data_summary = build_data_summary(data_df)
    data_skus = set(data_summary["Normalized SKU"])

    # Only keep WMS rows whose SKU appears in the Data file.
    wms_filtered = wms_df[wms_df["Normalized SKU"].isin(data_skus)].copy()

    wms_summary = build_wms_summary(wms_filtered)
    summary = data_summary.merge(wms_summary, on="Normalized SKU", how="left")

    fill_zero_cols = [
        "Matched WMS Rows", "Total WMS Quantity", "Current Location Count", "LPN Count"
    ]
    for c in fill_zero_cols:
        if c in summary.columns:
            summary[c] = summary[c].fillna(0)

    summary["Match Status"] = summary["Matched WMS Rows"].apply(
        lambda x: "MATCHED" if float(x) > 0 else "NO WMS MATCH"
    )
    summary["WMS Qty Minus Ordered Qty"] = (
        summary["Total WMS Quantity"].fillna(0) - summary["Total Quantity Ordered"].fillna(0)
    )

    # Days after target: positive means Consumer Priority Date is after the earliest customer target date.
    summary["Days: Latest CPD vs Earliest Target"] = (
        pd.to_datetime(summary["Latest Consumer Priority Date"], errors="coerce")
        - pd.to_datetime(summary["Earliest Customer Target Date"], errors="coerce")
    ).dt.days

    # Detail is one row per matching WMS pallet/row, enriched with Data summary fields.
    detail = wms_filtered.merge(data_summary, on="Normalized SKU", how="left")
    detail["Days: CPD vs Earliest Target"] = (
        pd.to_datetime(detail["Consumer Priority Date"], errors="coerce")
        - pd.to_datetime(detail["Earliest Customer Target Date"], errors="coerce")
    ).dt.days
    detail["Date Match Check"] = detail["Days: CPD vs Earliest Target"].apply(
        lambda x: "NO DATE" if pd.isna(x) else ("CPD ON/AFTER TARGET" if x >= 0 else "CPD BEFORE TARGET")
    )

    detail_columns = [
        "Normalized SKU", "Data Item Values", "Total Quantity Ordered", "Order Count", "Orders",
        "Earliest Customer Target Date", "Latest Customer Target Date", "Customer Target Dates",
        "Earliest Delivery Date", "Latest Delivery Date", "Delivery Dates",
        "WMS SKU Number", "WMS Quantity", "Consumer Priority Date", "Days: CPD vs Earliest Target",
        "Date Match Check", "Current Location Compact", "Current Location Normalized", "Current Location Zone",
        "Current Location Aisle", "Current Location Bay", "Current Location Position",
        "Previous Location Compact", "Previous Location Normalized", "Previous Location Zone",
        "Previous Location Aisle", "Previous Location Bay", "Previous Location Position",
        "LPN #", "WMS Row #",
    ]
    detail = detail[[c for c in detail_columns if c in detail.columns]].sort_values(
        ["Normalized SKU", "Consumer Priority Date", "Current Location Normalized"],
        na_position="last"
    ).reset_index(drop=True)

    data_lines = data_df.copy()
    line_match_info = summary[[
        "Normalized SKU", "Match Status", "Matched WMS Rows", "Total WMS Quantity",
        "Earliest Consumer Priority Date", "Latest Consumer Priority Date", "Current Locations"
    ]]
    data_lines = data_lines.merge(line_match_info, on="Normalized SKU", how="left")
    data_lines["Days: Latest CPD vs Target Date"] = (
        pd.to_datetime(data_lines["Latest Consumer Priority Date"], errors="coerce")
        - pd.to_datetime(data_lines["Target Date"], errors="coerce")
    ).dt.days

    data_lines_columns = [
        "Data Row #", "Data Item", "Normalized SKU", "Order Status", "Order",
        "Quantity Ordered", "Target Date", "Delivery Date", "Match Status",
        "Matched WMS Rows", "Total WMS Quantity", "Earliest Consumer Priority Date",
        "Latest Consumer Priority Date", "Days: Latest CPD vs Target Date", "Current Locations"
    ]
    data_lines = data_lines[[c for c in data_lines_columns if c in data_lines.columns]].sort_values(
        ["Normalized SKU", "Target Date", "Delivery Date"], na_position="last"
    ).reset_index(drop=True)

    unmatched = summary[summary["Match Status"] == "NO WMS MATCH"].copy()
    unmatched_columns = [
        "Normalized SKU", "Data Item Values", "Data Lines", "Order Count", "Orders",
        "Total Quantity Ordered", "Earliest Customer Target Date", "Latest Customer Target Date",
        "Customer Target Dates", "Earliest Delivery Date", "Latest Delivery Date", "Delivery Dates",
        "Match Status"
    ]
    unmatched = unmatched[[c for c in unmatched_columns if c in unmatched.columns]].reset_index(drop=True)

    summary_columns = [
        "Normalized SKU", "Data Item Values", "Match Status", "Total Quantity Ordered",
        "Total WMS Quantity", "WMS Qty Minus Ordered Qty", "Matched WMS Rows", "LPN Count",
        "Data Lines", "Order Count", "Orders", "Order Statuses",
        "Earliest Customer Target Date", "Latest Customer Target Date", "Customer Target Dates",
        "Earliest Delivery Date", "Latest Delivery Date", "Delivery Dates",
        "Earliest Consumer Priority Date", "Latest Consumer Priority Date", "Consumer Priority Dates",
        "Days: Latest CPD vs Earliest Target", "Current Location Count", "Current Zones",
        "Current Locations", "Previous Locations"
    ]
    summary = summary[[c for c in summary_columns if c in summary.columns]].sort_values(
        ["Match Status", "Normalized SKU"]
    ).reset_index(drop=True)

    return summary, detail, data_lines, unmatched


# ---------- Excel output ----------

DATE_COLUMNS = {
    "Target Date",
    "Delivery Date",
    "Earliest Customer Target Date",
    "Latest Customer Target Date",
    "Earliest Delivery Date",
    "Latest Delivery Date",
    "Consumer Priority Date",
    "Earliest Consumer Priority Date",
    "Latest Consumer Priority Date",
}


def convert_dates_for_excel(df):
    out = df.copy()
    for col in out.columns:
        if col in DATE_COLUMNS:
            converted = pd.to_datetime(out[col], errors="coerce")
            if converted.notna().any():
                out[col] = converted.dt.date
    return out


def write_df(writer, df, sheet_name):
    safe_df = convert_dates_for_excel(df)
    safe_df.to_excel(writer, sheet_name=sheet_name, index=False)


def format_workbook(writer):
    workbook = writer.book

    for ws in workbook.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions

        # Header style
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
            cell.fill = cell.fill.copy(fill_type="solid", fgColor="D9EAF7")
            cell.alignment = cell.alignment.copy(wrap_text=True, vertical="top")

        # Reasonable column widths
        for col_cells in ws.columns:
            column_letter = col_cells[0].column_letter
            header = str(col_cells[0].value or "")
            max_len = len(header)
            for cell in col_cells[1:2000]:
                value = cell.value
                if value is None:
                    continue
                max_len = max(max_len, len(str(value)))
            width = min(max(max_len + 2, 10), 45)
            ws.column_dimensions[column_letter].width = width

        # Number/date formats
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                header = ws.cell(row=1, column=cell.column).value
                header = str(header or "")
                if "Date" in header or "CPD" in header:
                    cell.number_format = "yyyy-mm-dd"
                elif "Quantity" in header or "Qty" in header or "Rows" in header or "Count" in header or "Days" in header:
                    cell.number_format = "#,##0.##"


def to_excel_bytes(summary_df, detail_df, data_lines_df, unmatched_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        write_df(writer, summary_df, "SKU SUMMARY")
        write_df(writer, detail_df, "MATCHED WMS DETAIL")
        write_df(writer, data_lines_df, "DATA LINES")
        write_df(writer, unmatched_df, "UNMATCHED DATA SKUS")
        format_workbook(writer)

    output.seek(0)
    return output


def date_only_for_display(df):
    out = df.copy()
    for col in out.columns:
        if col in DATE_COLUMNS:
            converted = pd.to_datetime(out[col], errors="coerce")
            if converted.notna().any():
                out[col] = converted.dt.date
    return out


# ---------- main ----------

if run_btn:
    if not data_file or not wms_file:
        st.error("Please upload both Excel files before building the matched output.")
    else:
        try:
            with st.spinner("Matching SKUs and building Excel..."):
                data_df = load_data_file(
                    data_file,
                    exclude_s=exclude_s_items,
                    decimals_only=decimal_only,
                )
                wms_df = load_wms_file(wms_file)
                summary_df, detail_df, data_lines_df, unmatched_df = build_outputs(data_df, wms_df)
                excel_bytes = to_excel_bytes(summary_df, detail_df, data_lines_df, unmatched_df)

        except KeyError as e:
            st.error(str(e))
            st.stop()
        except Exception as e:
            st.error(f"Something went wrong while processing the files: {e}")
            st.stop()

        st.success("Done.")

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Data SKUs", summary_df["Normalized SKU"].nunique())
        col2.metric("Matched SKUs", (summary_df["Match Status"] == "MATCHED").sum())
        col3.metric("Unmatched SKUs", (summary_df["Match Status"] == "NO WMS MATCH").sum())
        col4.metric("Matched WMS Rows", len(detail_df))

        st.download_button(
            "⬇️ Download matched SKU Excel",
            data=excel_bytes,
            file_name=f"sku_data_wms_match_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        tab1, tab2, tab3, tab4 = st.tabs([
            "SKU SUMMARY", "MATCHED WMS DETAIL", "DATA LINES", "UNMATCHED DATA SKUS"
        ])

        with tab1:
            st.dataframe(date_only_for_display(summary_df), use_container_width=True)

        with tab2:
            st.dataframe(date_only_for_display(detail_df), use_container_width=True)

        with tab3:
            st.dataframe(date_only_for_display(data_lines_df), use_container_width=True)

        with tab4:
            if unmatched_df.empty:
                st.info("Every Data SKU has at least one WMS match.")
            else:
                st.dataframe(date_only_for_display(unmatched_df), use_container_width=True)
else:
    st.info("Upload the Data/orders file and the big qPORT/WMS file, then click **Build Matched Excel**.")
