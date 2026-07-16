"""
Order vs Inventory Matcher — single-file Streamlit app.

Upload a sample-orders workbook and a QPORT inventory report.
The app matches every ordered item to its available quantity and locations,
one row per item, with every stocking location and its quantity listed together.

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""

import io
import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Order vs Inventory Matcher", layout="wide")

COL_LOC = [0, 1, 2, 3, 4, 5]
COL_SKU_H = 7
COL_SKU_I = 8
COL_QTY = 9
COL_PREV = [10, 11, 12, 13, 14, 15]
COL_DATE2 = 24
COL_TIME2 = 25

HIDDEN_COLS = {"_OrderSortKey"}


def _clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def normalize_order_no(order_no_numeric, raw_text):
    """
    Display text for Order #.
    Uses the already-parsed numeric value when available (e.g. 3700007249.0 -> "3700007249"),
    otherwise falls back to the original raw text (trimmed).
    """
    if pd.notna(order_no_numeric):
        return str(int(order_no_numeric))
    return _clean(raw_text)


def force_order_sort(df: pd.DataFrame) -> pd.DataFrame:
    """
    Final forced sort.

    Sorts by the numeric "_OrderSortKey" only (computed once, at load time).
    Does NOT sort by Item No.
    Keeps the original row order for ties / rows within the same order.
    Rows with no parseable Order # sort to the end.
    """
    df = df.copy()
    df["_line_order"] = range(len(df))

    df = df.sort_values(
        by=["_OrderSortKey", "_line_order"],
        ascending=[True, True],
        na_position="last",
        kind="mergesort",  # stable sort so ties keep their original relative order
    )

    df = df.drop(columns=["_line_order"]).reset_index(drop=True)
    return df


def build_sku(h, i):
    if pd.isna(h) or pd.isna(i):
        return None

    try:
        h = float(h)
        i = float(i)
    except (ValueError, TypeError):
        return None

    int_part = int(h)
    frac = int(round((h - int_part) * 100))
    suffix = int(round(i))

    return f"{int_part}.{frac:02d}{suffix:03d}"


def load_tkreserve(file) -> pd.DataFrame:
    raw = pd.read_excel(file, sheet_name=0, header=0)

    loc = raw.iloc[:, COL_LOC].apply(
        lambda row: "".join([_clean(v) for v in row if _clean(v) != ""]),
        axis=1
    )

    prev_loc = raw.iloc[:, COL_PREV].apply(
        lambda row: "".join([_clean(v) for v in row if _clean(v) != ""]),
        axis=1
    )

    h = pd.to_numeric(raw.iloc[:, COL_SKU_H], errors="coerce")
    i = pd.to_numeric(raw.iloc[:, COL_SKU_I], errors="coerce")

    sku = [build_sku(a, b) for a, b in zip(h, i)]

    qty = pd.to_numeric(raw.iloc[:, COL_QTY], errors="coerce").fillna(0)

    date2 = pd.to_numeric(raw.iloc[:, COL_DATE2], errors="coerce").fillna(0)
    time2 = pd.to_numeric(raw.iloc[:, COL_TIME2], errors="coerce").fillna(0)

    tx_time = date2 * 1_000_000 + time2

    df = pd.DataFrame({
        "SKU": sku,
        "Location": loc,
        "Quantity": qty,
        "Previous Location": prev_loc,
        "TxTime": tx_time,
    })

    df = df.dropna(subset=["SKU"])

    return df[df["SKU"] != ""]


def aggregate_inventory(inv: pd.DataFrame) -> pd.DataFrame:
    def summarize(g):
        loc_qty = (
            g.groupby("Location")["Quantity"]
            .sum()
            .sort_values(ascending=False)
        )

        loc_str = "; ".join(
            f"{loc} ({int(q)})"
            for loc, q in loc_qty.items()
            if loc
        )

        latest = g.loc[g["TxTime"].idxmax(), "Previous Location"]

        return pd.Series({
            "Quantity Available": g["Quantity"].sum(),
            "Current Location": loc_str,
            "Previous Location": latest,
        })

    return inv.groupby("SKU").apply(summarize, include_groups=False).reset_index()


def clean_qty(qty):
    """Numeric order quantity, as an int when it has no decimal part."""
    if pd.isna(qty):
        return 0

    try:
        qty = float(qty)
    except (ValueError, TypeError):
        return 0

    return int(qty) if qty.is_integer() else qty


def clean_um(um):
    """Unit of measure, upper-cased and trimmed (CA, EA, ...)."""
    return _clean(um).upper()


def load_sample_orders(file) -> pd.DataFrame:
    xl = pd.ExcelFile(file)

    frames = []

    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=0)

        if df.empty or "Item no" not in df.columns:
            continue

        df = df.dropna(subset=["Item no"]).copy()

        if df.empty:
            continue

        if "CO no" in df.columns:
            order_no_raw = df["CO no"]
        else:
            order_no_raw = pd.Series([""] * len(df), index=df.index)

        # Parse the numeric sort key ONCE here. Everything downstream
        # (display text and sort order) is derived from this same value,
        # so they can never disagree.
        order_no_numeric = pd.to_numeric(order_no_raw, errors="coerce")

        if "Order qty" in df.columns:
            qty_num = pd.to_numeric(df["Order qty"], errors="coerce").fillna(0)
        else:
            qty_num = pd.Series([0] * len(df), index=df.index)

        if "U/M" in df.columns:
            um = df["U/M"]
        else:
            um = pd.Series([""] * len(df), index=df.index)

        if "Name" in df.columns:
            item_name = df["Name"]
        else:
            item_name = pd.Series([""] * len(df), index=df.index)

        frames.append(pd.DataFrame({
            "Order Date": sheet,
            "Order #": [normalize_order_no(n, r) for n, r in zip(order_no_numeric, order_no_raw)],
            "_OrderSortKey": order_no_numeric,
            "Item No": df["Item no"].astype(str).str.strip(),
            "Item Name": item_name,
            # Quantity is now split: the number in one column, the unit in the next.
            "Order Qty": [clean_qty(q) for q in qty_num],
            "U/M": [clean_um(u) for u in um],
        }))

    if not frames:
        return pd.DataFrame(columns=[
            "Order Date",
            "Order #",
            "_OrderSortKey",
            "Item No",
            "Item Name",
            "Order Qty",
            "U/M",
        ])

    orders = pd.concat(frames, ignore_index=True)

    return orders


def is_sample_item(item_no: str) -> bool:
    item_no = str(item_no).strip()
    return not item_no[:1].isdigit()


def match_orders_to_inventory(orders: pd.DataFrame, tk_file) -> pd.DataFrame:
    inv_raw = load_tkreserve(tk_file)
    inv = aggregate_inventory(inv_raw)

    result = orders.merge(
        inv,
        how="left",
        left_on="Item No",
        right_on="SKU"
    ).drop(columns=["SKU"])

    result["Quantity Available"] = result["Quantity Available"].fillna(0)
    result["Current Location"] = result["Current Location"].fillna("Not found in inventory")
    result["Previous Location"] = result["Previous Location"].fillna("")

    cols = [
        "Order Date",
        "Order #",
        "_OrderSortKey",
        "Item No",
        "Item Name",
        "Order Qty",
        "U/M",
        "Current Location",
        "Previous Location",
        "Quantity Available",
    ]

    result = result[cols]

    # IMPORTANT:
    # This is the only sort.
    # It groups everything by Order #.
    # It does not sort by Item No.
    result = force_order_sort(result)

    return result


# ---------------------------------------------------------------------------
# Summary / dashboard
# ---------------------------------------------------------------------------

def build_summary(result: pd.DataFrame) -> dict:
    """
    Headline numbers plus two breakdown tables:
      by_um   — total quantity per unit of measure (CA, EA, ...)
      by_date — orders / lines / CA / EA per order date (one per source sheet)
    """
    df = result.copy()
    df["U/M"] = df["U/M"].astype(str).str.strip().str.upper()
    df["Order Qty"] = pd.to_numeric(df["Order Qty"], errors="coerce").fillna(0)

    um_totals = df.groupby("U/M")["Order Qty"].sum()

    headline = {
        "Order Dates": int(df["Order Date"].nunique()),
        "Orders": int(df["Order #"].nunique()),
        "Order Lines": int(len(df)),
        "Distinct Items": int(df["Item No"].nunique()),
        "Total Cases (CA)": int(um_totals.get("CA", 0)),
        "Total Each (EA)": int(um_totals.get("EA", 0)),
        "Zero Available": int((df["Quantity Available"] == 0).sum()),
        "Short on Stock": int((
            (df["Quantity Available"] > 0)
            & (df["Quantity Available"] < df["Order Qty"])
        ).sum()),
    }

    by_um = (
        df.groupby("U/M")
        .agg(Lines=("Order Qty", "size"), Total=("Order Qty", "sum"))
        .reset_index()
        .rename(columns={"Total": "Total Qty"})
        .sort_values("Total Qty", ascending=False)
        .reset_index(drop=True)
    )

    by_date = (
        df.groupby("Order Date")
        .agg(Orders=("Order #", "nunique"), Lines=("Item No", "size"))
        .reset_index()
    )

    for code, label in (("CA", "Cases (CA)"), ("EA", "Each (EA)")):
        sub = (
            df[df["U/M"] == code]
            .groupby("Order Date")["Order Qty"]
            .sum()
            .reindex(by_date["Order Date"])
            .fillna(0)
        )
        by_date[label] = sub.astype(int).values

    return {"headline": headline, "by_um": by_um, "by_date": by_date}


def _write_table(ws, start_row, title, df):
    """Write a titled table onto a worksheet, return the next free row."""
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", start_color="2F5496")

    cell = ws.cell(row=start_row, column=1, value=title)
    cell.font = Font(bold=True, size=12)

    r = start_row + 1

    for c, name in enumerate(df.columns, start=1):
        cell = ws.cell(row=r, column=c, value=str(name))
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    for _, row in df.iterrows():
        r += 1
        for c, name in enumerate(df.columns, start=1):
            ws.cell(row=r, column=c, value=row[name])

    return r + 2


def build_summary_sheet(wb, summary: dict):
    ws = wb.create_sheet("Summary")

    title = ws.cell(row=1, column=1, value="Order Summary")
    title.font = Font(bold=True, size=16)

    label_fill = PatternFill("solid", start_color="D9E2F3")

    r = 3
    for label, value in summary["headline"].items():
        lc = ws.cell(row=r, column=1, value=label)
        lc.font = Font(bold=True)
        lc.fill = label_fill

        vc = ws.cell(row=r, column=2, value=value)
        vc.font = Font(bold=True, size=12)
        vc.alignment = Alignment(horizontal="center")

        r += 1

    r += 1
    r = _write_table(ws, r, "Totals by Unit of Measure", summary["by_um"])
    r = _write_table(ws, r, "Breakdown by Order Date", summary["by_date"])

    ws.column_dimensions["A"].width = 20
    for col in "BCDE":
        ws.column_dimensions[col].width = 14

    return ws


def build_excel(result: pd.DataFrame) -> io.BytesIO:
    # Force the same order again before exporting.
    result = force_order_sort(result)

    wb = Workbook()
    ws = wb.active
    ws.title = "Order vs Inventory"

    display_cols = [c for c in result.columns if c not in HIDDEN_COLS]
    headers = display_cols

    ws.append(headers)

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", start_color="2F5496")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    zero_fill = PatternFill("solid", start_color="FFF2CC")

    for _, row in result.iterrows():
        ws.append([row[c] for c in display_cols])

    for offset, (_, row) in enumerate(result.iterrows()):
        r = offset + 2

        item_no = row["Item No"]
        qty_avail = row["Quantity Available"]

        is_sample = str(item_no)[:1].isalpha() if item_no is not None else False

        if qty_avail == 0 and not is_sample:
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).fill = zero_fill

    widths = {
        "Order Date": 11,
        "Order #": 13,
        "Item No": 13,
        "Item Name": 40,
        "Order Qty": 10,
        "U/M": 7,
        "Quantity Available": 16,
        "Current Location": 30,
        "Previous Location": 16
    }

    for idx, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(h, 14)

    ws.freeze_panes = "A2"

    # Summary goes last, so it's the final tab in the workbook.
    build_summary_sheet(wb, build_summary(result))

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    return buf


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

st.title("Sample Orders to Inventory Matcher")

st.write(
    "Upload the **sample orders** workbook and the **TKRESERVE_QPORT** inventory "
    "report. The app matches every ordered item to its available quantity and "
    "every location it's currently stocked in."
)

col1, col2 = st.columns(2)

with col1:
    orders_file = st.file_uploader(
        "Sample Orders (.xlsx)",
        type=["xlsx"],
        key="orders"
    )

with col2:
    tk_file = st.file_uploader(
        "QPORT Inventory Report (.xlsx)",
        type=["xlsx"],
        key="tk"
    )

if orders_file and tk_file:
    try:
        with st.spinner("Reading and matching..."):
            orders = load_sample_orders(orders_file)

            if orders.empty:
                st.error(
                    "No order rows found. Expected a column named 'Item no' "
                    "in the sample orders sheets."
                )
                st.stop()

            result = match_orders_to_inventory(orders, tk_file)

    except Exception as e:
        st.error(f"Something went wrong while processing the files: {e}")
        st.stop()

    summary = build_summary(result)
    h = summary["headline"]

    st.subheader("Summary")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Orders", h["Orders"])
    m2.metric("Order lines", h["Order Lines"])
    m3.metric("Total Cases (CA)", f"{h['Total Cases (CA)']:,}")
    m4.metric("Total Each (EA)", f"{h['Total Each (EA)']:,}")

    m5, m6, m7, m8 = st.columns(4)
    m5.metric("Distinct items", h["Distinct Items"])
    m6.metric("Order dates", h["Order Dates"])
    m7.metric("Zero available", h["Zero Available"])
    m8.metric("Short on stock", h["Short on Stock"])

    s1, s2 = st.columns(2)

    with s1:
        st.caption("Totals by unit of measure")
        st.dataframe(summary["by_um"], use_container_width=True, hide_index=True)
        st.bar_chart(summary["by_um"].set_index("U/M")["Total Qty"])

    with s2:
        st.caption("Breakdown by order date")
        st.dataframe(summary["by_date"], use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("Order lines")

    display_df = result.drop(columns=list(HIDDEN_COLS))

    n_cols = len(display_df.columns)
    row_styles = []

    for _, row in result.iterrows():
        is_sample = str(row["Item No"])[:1].isalpha()

        if row["Quantity Available"] == 0 and not is_sample:
            style = ["background-color: #fff2cc"] * n_cols
        else:
            style = [""] * n_cols

        row_styles.append(style)

    def highlight(row):
        return row_styles[row.name]

    st.dataframe(
        display_df.style.apply(highlight, axis=1),
        use_container_width=True,
        height=500
    )

    excel_buf = build_excel(result)

    st.download_button(
        "Download results as Excel",
        data=excel_buf,
        file_name="order_vs_inventory.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

else:
    st.info("Upload both files to run the match.")
