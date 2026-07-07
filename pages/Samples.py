"""
Order vs Inventory Matcher — single-file Streamlit app.

Upload a sample-orders workbook and a TKRESERVE_QPORT inventory report.
The app matches every ordered item to its available quantity and locations,
one row per item, with every stocking location and its quantity listed together.

Run with:
    pip install -r requirements.txt
    streamlit run app.py
"""
import io
import streamlit as st
import pandas as pd

st.set_page_config(page_title="Order vs Inventory Matcher", layout="wide")

# ---------------------------------------------------------------------------
# TKRESERVE_QPORT layout (0-based column positions, header row = row 1)
#   A-F : Location hierarchy (zone, aisle, rack, shelf, position, level)
#         -> combined into one "Location" string, e.g. "BL-A-C-KH-O-L"
#   H   : Sku Number prefix + 2-digit fraction (e.g. 71117.00)
#   I   : Sku Number last 3 digits of the suffix (e.g. 225)
#         Full SKU = f"{int(H)}.{frac:02d}{int(I):03d}" -> "71117.00225"
#   J   : QUANTITY (units available for that location/lot)
#   K   : Previous  (previous location/zone code, e.g. RC, TK, TR, BL, RE, RF)
# ---------------------------------------------------------------------------
COL_LOC = [0, 1, 2, 3, 4, 5]
COL_SKU_H = 7
COL_SKU_I = 8
COL_QTY = 9
COL_PREV = 10


def _clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def build_sku(h, i):
    """Reconstruct the full SKU number from the H (prefix.fraction) and I (suffix) columns."""
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
    """Read TKRESERVE_QPORT.xlsx -> ['SKU', 'Location', 'Quantity', 'Previous Location']"""
    raw = pd.read_excel(file, sheet_name=0, header=0)

    loc = raw.iloc[:, COL_LOC].apply(
        lambda row: "-".join([_clean(v) for v in row if _clean(v) != ""]), axis=1
    )

    h = pd.to_numeric(raw.iloc[:, COL_SKU_H], errors="coerce")
    i = pd.to_numeric(raw.iloc[:, COL_SKU_I], errors="coerce")
    sku = [build_sku(a, b) for a, b in zip(h, i)]

    qty = pd.to_numeric(raw.iloc[:, COL_QTY], errors="coerce").fillna(0)
    prev = raw.iloc[:, COL_PREV].apply(_clean)

    df = pd.DataFrame({"SKU": sku, "Location": loc, "Quantity": qty, "Previous Location": prev})
    df = df.dropna(subset=["SKU"])
    return df[df["SKU"] != ""]


def aggregate_inventory(inv: pd.DataFrame) -> pd.DataFrame:
    """One row per SKU:
    - Quantity Available = sum of all quantities for that SKU
    - Current Location   = every distinct location with its quantity, e.g. 'BL-A-C-KH-O-L (67); RC-2-A-1-X-2 (91)'
    - Previous Location  = every distinct previous-location code seen for that SKU
    """
    def summarize(g):
        loc_qty = g.groupby("Location")["Quantity"].sum().sort_values(ascending=False)
        loc_str = "; ".join(f"{loc} ({int(q)})" for loc, q in loc_qty.items() if loc)
        prev_vals = sorted(set(v for v in g["Previous Location"] if v))
        return pd.Series({
            "Quantity Available": g["Quantity"].sum(),
            "Current Location": loc_str,
            "Previous Location": "; ".join(prev_vals),
        })

    return inv.groupby("SKU").apply(summarize, include_groups=False).reset_index()


def load_sample_orders(file) -> pd.DataFrame:
    """Stack every non-empty sheet -> Order Date, CO No, Item No, Item Name, Order Qty"""
    xl = pd.ExcelFile(file)
    frames = []
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=0)
        if df.empty or "Item no" not in df.columns:
            continue
        df = df.dropna(subset=["Item no"])
        if df.empty:
            continue
        co = pd.to_numeric(df.get("CO no"), errors="coerce")
        frames.append(pd.DataFrame({
            "Order Date": sheet,
            "CO No": co.apply(lambda v: str(int(v)) if pd.notna(v) else ""),
            "Item No": df["Item no"].astype(str).str.strip(),
            "Item Name": df.get("Name"),
            "Order Qty": pd.to_numeric(df.get("Order qty"), errors="coerce").fillna(0),
        }))
    if not frames:
        return pd.DataFrame(columns=["Order Date", "CO No", "Item No", "Item Name", "Order Qty"])
    return pd.concat(frames, ignore_index=True)


def match_orders_to_inventory(orders: pd.DataFrame, tk_file) -> pd.DataFrame:
    inv_raw = load_tkreserve(tk_file)
    inv = aggregate_inventory(inv_raw)

    result = orders.merge(inv, how="left", left_on="Item No", right_on="SKU").drop(columns=["SKU"])
    result["Quantity Available"] = result["Quantity Available"].fillna(0)
    result["Current Location"] = result["Current Location"].fillna("Not found in inventory")
    result["Previous Location"] = result["Previous Location"].fillna("")

    cols = ["Order Date", "CO No", "Item No", "Item Name", "Order Qty",
            "Quantity Available", "Current Location", "Previous Location"]
    return result[cols].sort_values(["Order Date", "Item No"]).reset_index(drop=True)


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📦 Sample Orders → TKRESERVE Inventory Matcher")
st.write(
    "Upload the **sample orders** workbook and the **TKRESERVE_QPORT** inventory "
    "report. The app matches every ordered item to its available quantity and "
    "every location it's currently stocked in."
)

col1, col2 = st.columns(2)
with col1:
    orders_file = st.file_uploader("Sample Orders (.xlsx)", type=["xlsx"], key="orders")
with col2:
    tk_file = st.file_uploader("TKRESERVE_QPORT (.xlsx)", type=["xlsx"], key="tk")

if orders_file and tk_file:
    try:
        with st.spinner("Reading and matching..."):
            orders = load_sample_orders(orders_file)
            if orders.empty:
                st.error("No order rows found (expected a column named 'Item no' in the sample orders sheets).")
                st.stop()
            result = match_orders_to_inventory(orders, tk_file)
    except Exception as e:
        st.error(f"Something went wrong while processing the files: {e}")
        st.stop()

    not_found = (result["Current Location"] == "Not found in inventory").sum()
    short = (result["Quantity Available"] < result["Order Qty"]).sum()

    m1, m2, m3 = st.columns(3)
    m1.metric("Order lines", len(result))
    m2.metric("Not found in inventory", int(not_found))
    m3.metric("Short on stock", int(short))

    def highlight(row):
        if row["Current Location"] == "Not found in inventory":
            return ["background-color: #fff2cc"] * len(row)
        if row["Quantity Available"] < row["Order Qty"]:
            return ["background-color: #ffe0e0"] * len(row)
        return [""] * len(row)

    st.dataframe(result.style.apply(highlight, axis=1), use_container_width=True, height=500)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        result.to_excel(writer, index=False, sheet_name="Order vs Inventory")
    buf.seek(0)

    st.download_button(
        "⬇️ Download results as Excel",
        data=buf,
        file_name="order_vs_inventory.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Upload both files to run the match.")
