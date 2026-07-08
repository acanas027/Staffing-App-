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


def format_qty(qty, um):
    qty = 0 if pd.isna(qty) else qty

    try:
        qty_str = str(int(qty)) if float(qty).is_integer() else str(qty)
    except (ValueError, TypeError):
        qty_str = str(qty)

    um = _clean(um)

    return f"{qty_str} {um}".strip()


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
            "Order Qty Num": qty_num,
            "Order Qty": [format_qty(q, u) for q, u in zip(qty_num, um)],
        }))

    if not frames:
        return pd.DataFrame(columns=[
            "Order Date",
            "Order #",
            "_OrderSortKey",
            "Item No",
            "Item Name",
            "Order Qty Num",
            "Order Qty"
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
        "Current Location",
        "Previous Location",
        "Quantity Available",
        "Order Qty Num"
    ]

    result = result[cols]

    # IMPORTANT:
    # This is the only sort.
    # It groups everything by Order #.
    # It does not sort by Item No.
    result = force_order_sort(result)

    return result


def build_excel(result: pd.DataFrame) -> io.BytesIO:
    # Force the same order again before exporting.
    result = force_order_sort(result)

    wb = Workbook()
    ws = wb.active
    ws.title = "Order vs Inventory"

    hidden_cols = {"Order Qty Num", "_OrderSortKey"}
    display_cols = [c for c in result.columns if c not in hidden_cols]
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
        "Quantity Available": 16,
        "Current Location": 30,
        "Previous Location": 16
    }

    for idx, h in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = widths.get(h, 14)

    ws.freeze_panes = "A2"

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

    zero_avail = (result["Quantity Available"] == 0).sum()

    short = (
        (result["Quantity Available"] > 0)
        & (result["Quantity Available"] < result["Order Qty Num"])
    ).sum()

    m1, m2, m3 = st.columns(3)

    m1.metric("Order lines", len(result))
    m2.metric("Zero available", int(zero_avail))
    m3.metric("Short on stock", int(short))

    display_df = result.drop(columns=["Order Qty Num", "_OrderSortKey"])

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
