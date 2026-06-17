import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Inventory Shortage Checker", layout="wide")
st.title("Inventory Shortage Checker")
st.caption("Upload the orders file and the warehouse short code file to find shortages, expiration-date research items, and aged inventory.")

with st.expander("How this works"):
    st.markdown(
        "- Items starting with **S** are removed from the orders file.\n"
        "- Item codes are matched to SKU Numbers using the first 7 characters "
        "of the zero-padded code (e.g. `06795.48940` matches `6795.48`).\n"
        "- Order lines are processed most-urgent-Target-Date first, allocating "
        "inventory whose Consumer Priority Date is on or after the Target Date, "
        "soonest-qualifying-date first. Inventory is not double-counted across orders.\n"
        "- **SHORT SHEET**: order lines that don't have enough qualifying inventory.\n"
        "- **EMAIL / RESEARCH**: one row per Item + Location Zone (first 3 "
        "characters of the location, e.g. RC3, RF2) for fully-covered "
        "order lines where that zone does not start with RC2 — shows total "
        "quantity needed from that zone, earliest delivery date, and the "
        "previous location(s) involved, so you can send one email per product.\n"
        "- **SKU TO CODE**: all warehouse rows with a Consumer Priority Date on "
        "or before the cutoff date (default today)."
    )

with st.sidebar:
    st.header("Upload Files")
    data_file = st.file_uploader("Data (orders) file", type=["xlsx"], key="data_file")
    short_file = st.file_uploader("Short code (WMS) file", type=["xlsx"], key="short_file")
    st.divider()
    cutoff_date = st.date_input("SKU TO CODE cutoff date", value=datetime.now().date(),
                                 help="SKU TO CODE will include rows with a Consumer Priority Date on or before this date.")
    st.divider()
    run_btn = st.button("Run Analysis", type="primary", use_container_width=True)


# ---------- Normalization helpers ----------

def cell_to_str(x):
    if pd.isna(x):
        return ""
    if isinstance(x, float) and x == int(x):
        return str(int(x))
    return str(x).strip()


def sku_norm(x):
    if pd.isna(x):
        return None
    if isinstance(x, float):
        if x == int(x):
            return str(int(x))
        return str(x)
    return str(x).strip()


def make_key(s):
    """Build a matching key between Item codes (data file) and SKU Number
    (short code file). Two cases:
    - Decimal codes (e.g. 06795.48940): zero-pad the whole-number part to
      5 digits, keep the decimal part as-is, then take the first 7 characters.
      This intentionally matches on the first 2 decimal digits only, since
      the short code file truncates decimals to 2 places.
    - Whole-number-only codes (e.g. 28005, no decimal at all): zero-pad to
      5 digits and match directly against bare whole-number SKUs.
    A whole-number SKU never matches a decimal Item code and vice versa."""
    if s is None:
        return None
    s = str(s)
    if "." in s:
        whole, dec = s.split(".", 1)
        whole = whole.zfill(5)
        key = f"{whole}.{dec}"
        return key[:7]
    else:
        return s.zfill(5)


def load_data_file(file):
    df = pd.read_excel(file, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df["Item"].notna()].copy()
    df["Item"] = df["Item"].astype(str).str.strip()
    df = df[~df["Item"].str.upper().str.startswith("S")].copy()
    df["Target Date"] = pd.to_datetime(df["Target Date"])
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"])
    df["Quantity Ordered"] = pd.to_numeric(df["Quantity Ordered"], errors="coerce").fillna(0)
    df["match_key"] = df["Item"].apply(make_key)
    df = df.reset_index(drop=True)
    df["order_line_id"] = df.index
    return df


def load_short_code_file(file):
    raw = pd.read_excel(file, sheet_name=0, header=None, skiprows=3)

    location = raw.iloc[:, 0:6].map(cell_to_str).agg("".join, axis=1)
    previous_location = raw.iloc[:, 10:16].map(cell_to_str).agg("".join, axis=1)
    lpn = raw.iloc[:, 6]
    sku_raw = raw.iloc[:, 7]
    quantity = pd.to_numeric(raw.iloc[:, 9], errors="coerce").fillna(0)
    priority_date_raw = raw.iloc[:, 27]

    out = pd.DataFrame({
        "Location": location,
        "Previous Location": previous_location,
        "LPN #": lpn,
        "SKU Number": sku_raw.apply(sku_norm),
        "Quantity": quantity,
        "Consumer Priority Date": pd.to_datetime(priority_date_raw.astype(str), format="%Y%m%d", errors="coerce"),
    })
    out = out[out["SKU Number"].notna()].copy()
    out["match_key"] = out["SKU Number"].apply(make_key)
    out = out.reset_index(drop=True)
    out["wms_row_id"] = out.index
    return out


def run_allocation(orders_df, wms_df):
    """Allocate WMS inventory to order lines, most urgent Target Date first.
    Returns: short_df, email_df, unmatched_df (order lines whose Item has
    zero matching SKU rows in the WMS file at all)."""

    remaining = wms_df.set_index("wms_row_id")["Quantity"].astype(float).to_dict()

    by_key = {}
    for key, group in wms_df.groupby("match_key"):
        by_key[key] = group.sort_values("Consumer Priority Date").copy()

    short_rows = []
    email_rows = []
    unmatched_rows = []

    orders_sorted = orders_df.sort_values("Target Date", ascending=True)

    for _, order in orders_sorted.iterrows():
        key = order["match_key"]
        target_date = order["Target Date"]
        need = float(order["Quantity Ordered"])

        candidates = by_key.get(key)

        if candidates is None:
            unmatched_rows.append({
                "Item": order["Item"],
                "Quantity Needed": need,
                "Target Date": order["Target Date"],
                "Customer": order.get("Customer", ""),
                "Order": order.get("Order", ""),
            })
            short_rows.append({
                "Item": order["Item"],
                "Quantity Needed": need,
                "Quantity Available (qualifying)": 0,
                "Short By": need,
                "Target Date": order["Target Date"],
                "Customer": order.get("Customer", ""),
                "Order": order.get("Order", ""),
            })
            continue

        allocated_total = 0.0
        sources = []

        qualifying = candidates[candidates["Consumer Priority Date"] >= target_date]
        for _, wms_row in qualifying.iterrows():
            if need - allocated_total <= 0:
                break
            avail = remaining.get(wms_row["wms_row_id"], 0.0)
            if avail <= 0:
                continue
            take = min(avail, need - allocated_total)
            if take <= 0:
                continue
            remaining[wms_row["wms_row_id"]] = avail - take
            allocated_total += take
            sources.append({
                "wms_row_id": wms_row["wms_row_id"],
                "qty_taken": take,
                "Location": wms_row["Location"],
                "Previous Location": wms_row["Previous Location"],
                "Consumer Priority Date": wms_row["Consumer Priority Date"],
            })

        short_qty = need - allocated_total

        if short_qty > 0.0001:
            short_rows.append({
                "Item": order["Item"],
                "Quantity Needed": need,
                "Quantity Available (qualifying)": allocated_total,
                "Short By": short_qty,
                "Target Date": order["Target Date"],
                "Customer": order.get("Customer", ""),
                "Order": order.get("Order", ""),
            })
        else:
            non_rc2_sources = [s for s in sources if not str(s["Location"]).upper().startswith("RC2")]
            if non_rc2_sources:
                for s in non_rc2_sources:
                    email_rows.append({
                        "Item": order["Item"],
                        "Customer": order.get("Customer", ""),
                        "Order": order.get("Order", ""),
                        "Quantity Needed": need,
                        "Delivery Date": order["Delivery Date"],
                        "Location": s["Location"],
                        "Previous Location": s["Previous Location"],
                        "Quantity from this location": s["qty_taken"],
                    })

    short_df = pd.DataFrame(short_rows)
    email_df = pd.DataFrame(email_rows)
    unmatched_df = pd.DataFrame(unmatched_rows)
    return short_df, email_df, unmatched_df


def build_sku_to_code(wms_df, cutoff):
    cutoff_ts = pd.Timestamp(cutoff)
    aged = wms_df[wms_df["Consumer Priority Date"] <= cutoff_ts].copy()
    aged = aged[["SKU Number", "Quantity", "Consumer Priority Date"]].sort_values("Consumer Priority Date")
    return aged.reset_index(drop=True)


def build_email_summary(email_df):
    """Collapse the detailed EMAIL/RESEARCH rows into one row per
    Item + Location Zone (first 3 characters of Location, e.g. RC3, RF2),
    suitable for writing one email per product/zone."""
    if email_df.empty:
        return pd.DataFrame(columns=[
            "Item", "Location Zone", "Total Quantity Needed", "Earliest Delivery Date",
            "Previous Locations", "Orders Affected"
        ])

    email_df = email_df.copy()
    email_df["Location Zone"] = email_df["Location"].astype(str).str[:3]

    def agg_group(g):
        return pd.Series({
            "Total Quantity Needed": g["Quantity from this location"].sum(),
            "Earliest Delivery Date": g["Delivery Date"].min(),
            "Previous Locations": ", ".join(sorted(set(g["Previous Location"].astype(str)))),
            "Orders Affected": g["Order"].nunique(),
        })

    summary = email_df.groupby(["Item", "Location Zone"], as_index=False).apply(agg_group, include_groups=False)
    summary = summary.sort_values(["Item", "Earliest Delivery Date"]).reset_index(drop=True)
    return summary


def to_excel_bytes(short_df, email_summary_df, sku_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        short_out = short_df.copy()
        if not short_out.empty:
            short_out["Target Date"] = pd.to_datetime(short_out["Target Date"]).dt.date
        (short_out if not short_out.empty else pd.DataFrame(columns=[
            "Item", "Quantity Needed", "Quantity Available (qualifying)", "Short By",
            "Target Date", "Customer", "Order"
        ])).to_excel(writer, sheet_name="SHORT SHEET", index=False)

        summary_out = email_summary_df.copy()
        if not summary_out.empty:
            summary_out["Earliest Delivery Date"] = pd.to_datetime(summary_out["Earliest Delivery Date"]).dt.date
        (summary_out if not summary_out.empty else pd.DataFrame(columns=[
            "Item", "Location Zone", "Total Quantity Needed", "Earliest Delivery Date",
            "Previous Locations", "Orders Affected"
        ])).to_excel(writer, sheet_name="EMAIL_RESEARCH", index=False)

        sku_out = sku_df.copy()
        if not sku_out.empty:
            sku_out["Consumer Priority Date"] = pd.to_datetime(sku_out["Consumer Priority Date"]).dt.date
        (sku_out if not sku_out.empty else pd.DataFrame(columns=[
            "SKU Number", "Quantity", "Consumer Priority Date"
        ])).to_excel(writer, sheet_name="SKU TO CODE", index=False)

    output.seek(0)
    return output


if run_btn:
    if not data_file or not short_file:
        st.error("Please upload both files before running the analysis.")
    else:
        try:
            with st.spinner("Processing..."):
                orders_df = load_data_file(data_file)
                wms_df = load_short_code_file(short_file)
                short_df, email_df, unmatched_df = run_allocation(orders_df, wms_df)
                sku_df = build_sku_to_code(wms_df, cutoff_date)
                email_summary_df = build_email_summary(email_df)
        except KeyError as e:
            st.error(f"The orders file is missing an expected column: {e}. "
                     f"Expected columns: Item, Target Date, Delivery Date, Quantity Ordered.")
            st.stop()
        except Exception as e:
            st.error(f"Something went wrong while processing the files: {e}")
            st.stop()

        st.success("Done.")

        if not unmatched_df.empty:
            st.warning(f"{unmatched_df['Item'].nunique()} item(s) on order have no matching SKU at all "
                       f"in the short code file (counted as fully short). See the 'Unmatched Items' tab.")

        col1, col2, col3 = st.columns(3)
        col1.metric("Short order lines", len(short_df))
        col2.metric("Email/Research items+locations", len(email_summary_df))
        col3.metric("Aged SKU rows (cutoff or earlier)", len(sku_df))

        excel_bytes = to_excel_bytes(short_df, email_summary_df, sku_df)
        st.download_button(
            "Download results (Excel)",
            data=excel_bytes,
            file_name="shortage_analysis.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        def date_only(df, cols):
            df = df.copy()
            for c in cols:
                if c in df.columns:
                    df[c] = pd.to_datetime(df[c]).dt.date
            return df

        tab1, tab2, tab3, tab4 = st.tabs(["SHORT SHEET", "EMAIL / RESEARCH", "SKU TO CODE", "Unmatched Items"])
        with tab1:
            st.dataframe(date_only(short_df, ["Target Date"]), use_container_width=True)
        with tab2:
            st.dataframe(date_only(email_summary_df, ["Earliest Delivery Date"]), use_container_width=True)
        with tab3:
            st.dataframe(date_only(sku_df, ["Consumer Priority Date"]), use_container_width=True)
        with tab4:
            st.dataframe(date_only(unmatched_df, ["Target Date"]), use_container_width=True)
else:
    st.info("Upload both files in the sidebar, then click **Run Analysis**.")
