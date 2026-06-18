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
        "- The full SKU Number is built directly from the short code file's "
        "columns H and I (H gives the whole number plus the first 2 decimal "
        "digits, I gives the remaining 3 decimal digits), then matched "
        "exactly against the Item code.\n"
        "- Order lines are processed most-urgent-Target-Date first, allocating "
        "inventory whose Consumer Priority Date is on or after the Target Date, "
        "soonest-qualifying-date first. Inventory is not double-counted across orders.\n"
        "- **SHORT SHEET**: one row per Item/SKU that doesn't have enough "
        "qualifying inventory, with the total quantity short across all "
        "affected orders and the earliest Target Date among them.\n"
        "- **EMAIL / RESEARCH**: one row per Item + Location Zone for fully-covered "
        "order lines where that zone does not start with RC2 and the "
        "Delivery Date falls within the chosen departure window.\n"
        "- **SKU TO CODE**: aged inventory rows (Consumer Priority Date ≤ cutoff) "
        "that are **new** — not seen in any previous run. One row per pallet, "
        "showing Location and LPN #. Dedup key is SKU Number + Consumer Priority "
        "Date combined, so the same SKU can reappear legitimately if it has a "
        "different date.\n"
        "- **SKU HISTORY**: every SKU+date pair ever flagged, accumulated across "
        "all runs. Upload last run's Excel to carry the history forward."
    )

with st.sidebar:
    st.header("Upload Files")
    data_file = st.file_uploader("Data (orders) file", type=["xlsx"], key="data_file")
    short_file = st.file_uploader("Short code (WMS) file", type=["xlsx"], key="short_file")
    st.divider()
    st.subheader("History (optional)")
    history_file = st.file_uploader(
        "Previous report (for dedup)",
        type=["xlsx"],
        key="history_file",
        help="Upload the Excel downloaded from your last run. The app reads its "
             "'SKU HISTORY' sheet to filter out SKU+date pairs already seen."
    )
    st.divider()
    cutoff_date = st.date_input(
        "SKU TO CODE cutoff date",
        value=datetime.now().date(),
        help="SKU TO CODE includes rows with a Consumer Priority Date on or before this date."
    )
    st.divider()
    today = datetime.now().date()
    email_date_range = st.date_input(
        "EMAIL / RESEARCH departure window",
        value=(today, today + pd.Timedelta(days=3)),
        help="Only include order lines whose Delivery Date falls in this range (inclusive)."
    )
    st.divider()
    run_btn = st.button("Run Analysis", type="primary", use_container_width=True)


# ---------- helpers ----------

def cell_to_str(x):
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


def build_full_sku(h_val, i_val):
    if h_val is None:
        return None
    if "." in h_val:
        whole, dec = h_val.split(".", 1)
        whole = whole.zfill(5)
        dec = dec.ljust(2, "0")[:2]
        return f"{whole}.{dec}{i_val}"
    else:
        return h_val.zfill(5)


def load_data_file(file):
    df = pd.read_excel(file, sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]
    df = df[df["Item"].notna()].copy()
    df["Item"] = df["Item"].astype(str).str.strip()
    df = df[~df["Item"].str.upper().str.startswith("S")].copy()
    df["Target Date"] = pd.to_datetime(df["Target Date"])
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"])
    df["Quantity Ordered"] = pd.to_numeric(df["Quantity Ordered"], errors="coerce").fillna(0)
    df["match_key"] = df["Item"]
    df = df.reset_index(drop=True)
    df["order_line_id"] = df.index
    return df


def load_short_code_file(file):
    raw = pd.read_excel(file, sheet_name=0, header=None, skiprows=3)

    location = raw.iloc[:, 0:6].map(cell_to_str).agg("".join, axis=1)
    previous_location = raw.iloc[:, 10:16].map(cell_to_str).agg("".join, axis=1)
    lpn = raw.iloc[:, 6]
    sku_h = raw.iloc[:, 7].apply(h_to_str)
    sku_i = raw.iloc[:, 8].apply(i_to_str)
    full_sku = [build_full_sku(h, i) for h, i in zip(sku_h, sku_i)]
    quantity = pd.to_numeric(raw.iloc[:, 9], errors="coerce").fillna(0)
    priority_date_raw = raw.iloc[:, 27]

    out = pd.DataFrame({
        "Location": location,
        "Previous Location": previous_location,
        "LPN #": lpn,
        "SKU Number": full_sku,
        "Quantity": quantity,
        "Consumer Priority Date": pd.to_datetime(
            priority_date_raw.astype(str), format="%Y%m%d", errors="coerce"
        ),
    })
    out = out[out["SKU Number"].notna()].copy()
    out["match_key"] = out["SKU Number"]
    out = out.reset_index(drop=True)
    out["wms_row_id"] = out.index
    return out


def load_history(file):
    """Read the SKU HISTORY sheet from a previous report.
    Returns a set of (sku_number, date_str) tuples."""
    if file is None:
        return set()
    try:
        hist_df = pd.read_excel(file, sheet_name="SKU HISTORY")
        hist_df["Consumer Priority Date"] = pd.to_datetime(
            hist_df["Consumer Priority Date"], errors="coerce"
        )
        seen = set()
        for _, row in hist_df.iterrows():
            sku = str(row["SKU Number"]).strip()
            dt = row["Consumer Priority Date"]
            if pd.notna(dt):
                seen.add((sku, str(dt.date())))
        return seen
    except Exception:
        return set()


def run_allocation(orders_df, wms_df):
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


def build_sku_to_code(wms_df, cutoff, seen_pairs):
    """Return aged WMS rows whose (SKU, date) pair has NOT been seen before.
    One row per pallet, including Location and LPN #.
    Also returns the set of new (sku, date_str) pairs to add to history."""
    cutoff_ts = pd.Timestamp(cutoff)
    aged = wms_df[wms_df["Consumer Priority Date"] <= cutoff_ts].copy()
    aged = aged[
        ["SKU Number", "LPN #", "Location", "Quantity", "Consumer Priority Date"]
    ].sort_values(["Consumer Priority Date", "SKU Number", "Location"])

    def is_new(row):
        sku = str(row["SKU Number"]).strip()
        dt = row["Consumer Priority Date"]
        if pd.isna(dt):
            return False
        return (sku, str(dt.date())) not in seen_pairs

    aged["_is_new"] = aged.apply(is_new, axis=1)
    new_rows = aged[aged["_is_new"]].drop(columns=["_is_new"]).reset_index(drop=True)

    # Collect new unique SKU+date pairs to add to history
    new_pairs = set()
    for _, row in new_rows.iterrows():
        dt = row["Consumer Priority Date"]
        if pd.notna(dt):
            new_pairs.add((str(row["SKU Number"]).strip(), str(dt.date())))

    return new_rows, new_pairs


def build_updated_history(seen_pairs, new_pairs):
    """Merge old history + new pairs into a single history DataFrame."""
    all_pairs = seen_pairs | new_pairs
    rows = []
    for sku, date_str in sorted(all_pairs):
        rows.append({"SKU Number": sku, "Consumer Priority Date": date_str})
    hist_df = pd.DataFrame(rows)
    if not hist_df.empty:
        hist_df["Consumer Priority Date"] = pd.to_datetime(
            hist_df["Consumer Priority Date"], errors="coerce"
        )
        hist_df = hist_df.sort_values(["Consumer Priority Date", "SKU Number"]).reset_index(drop=True)
    return hist_df


def build_short_summary(short_df):
    if short_df.empty:
        return pd.DataFrame(columns=["Item", "Total Short By", "Earliest Target Date"])

    def agg_group(g):
        return pd.Series({
            "Total Short By": g["Short By"].sum(),
            "Earliest Target Date": g["Target Date"].min(),
        })

    summary = short_df.groupby("Item", as_index=False).apply(agg_group, include_groups=False)
    summary = summary.sort_values("Earliest Target Date").reset_index(drop=True)
    return summary


def build_email_summary(email_df, date_range=None):
    if email_df.empty:
        return pd.DataFrame(columns=[
            "Item", "Location Zone", "Total Quantity Needed", "Earliest Delivery Date",
            "Previous Locations", "Orders Affected"
        ])

    email_df = email_df.copy()

    if date_range is not None and len(date_range) == 2:
        start_ts = pd.Timestamp(date_range[0])
        end_ts = pd.Timestamp(date_range[1])
        email_df = email_df[
            (email_df["Delivery Date"] >= start_ts) & (email_df["Delivery Date"] <= end_ts)
        ]

    if email_df.empty:
        return pd.DataFrame(columns=[
            "Item", "Location Zone", "Total Quantity Needed", "Earliest Delivery Date",
            "Previous Locations", "Orders Affected"
        ])

    email_df["Location Zone"] = email_df["Location"].astype(str).str[:3]

    def agg_group(g):
        return pd.Series({
            "Total Quantity Needed": g["Quantity from this location"].sum(),
            "Earliest Delivery Date": g["Delivery Date"].min(),
            "Previous Locations": ", ".join(sorted(set(g["Previous Location"].astype(str)))),
            "Orders Affected": g["Order"].nunique(),
        })

    summary = email_df.groupby(["Item", "Location Zone"], as_index=False).apply(
        agg_group, include_groups=False
    )
    summary = summary.sort_values(["Item", "Earliest Delivery Date"]).reset_index(drop=True)
    return summary


def to_excel_bytes(short_summary_df, email_summary_df, sku_df, history_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # SHORT SHEET
        short_out = short_summary_df.copy()
        if not short_out.empty:
            short_out["Earliest Target Date"] = pd.to_datetime(short_out["Earliest Target Date"]).dt.date
        (short_out if not short_out.empty else pd.DataFrame(
            columns=["Item", "Total Short By", "Earliest Target Date"]
        )).to_excel(writer, sheet_name="SHORT SHEET", index=False)

        # EMAIL_RESEARCH
        summary_out = email_summary_df.copy()
        if not summary_out.empty:
            summary_out["Earliest Delivery Date"] = pd.to_datetime(
                summary_out["Earliest Delivery Date"]
            ).dt.date
        (summary_out if not summary_out.empty else pd.DataFrame(columns=[
            "Item", "Location Zone", "Total Quantity Needed", "Earliest Delivery Date",
            "Previous Locations", "Orders Affected"
        ])).to_excel(writer, sheet_name="EMAIL_RESEARCH", index=False)

        # SKU TO CODE (new only, one row per pallet)
        sku_out = sku_df.copy()
        if not sku_out.empty:
            sku_out["Consumer Priority Date"] = pd.to_datetime(
                sku_out["Consumer Priority Date"]
            ).dt.date
        (sku_out if not sku_out.empty else pd.DataFrame(
            columns=["SKU Number", "LPN #", "Location", "Quantity", "Consumer Priority Date"]
        )).to_excel(writer, sheet_name="SKU TO CODE", index=False)

        # SKU HISTORY (cumulative)
        hist_out = history_df.copy()
        if not hist_out.empty:
            hist_out["Consumer Priority Date"] = pd.to_datetime(
                hist_out["Consumer Priority Date"]
            ).dt.date
        (hist_out if not hist_out.empty else pd.DataFrame(
            columns=["SKU Number", "Consumer Priority Date"]
        )).to_excel(writer, sheet_name="SKU HISTORY", index=False)

    output.seek(0)
    return output


# ---------- main ----------

if run_btn:
    if not data_file or not short_file:
        st.error("Please upload both files before running the analysis.")
    else:
        try:
            with st.spinner("Processing..."):
                orders_df = load_data_file(data_file)
                wms_df = load_short_code_file(short_file)
                seen_pairs = load_history(history_file)
                short_df, email_df, unmatched_df = run_allocation(orders_df, wms_df)
                short_summary_df = build_short_summary(short_df)
                sku_df, new_pairs = build_sku_to_code(wms_df, cutoff_date, seen_pairs)
                history_df = build_updated_history(seen_pairs, new_pairs)

                if isinstance(email_date_range, (tuple, list)) and len(email_date_range) == 2:
                    email_summary_df = build_email_summary(email_df, email_date_range)
                else:
                    st.warning(
                        "Please select both a start and end date for the EMAIL / RESEARCH "
                        "departure window. Showing all dates for now."
                    )
                    email_summary_df = build_email_summary(email_df)

        except KeyError as e:
            st.error(
                f"The orders file is missing an expected column: {e}. "
                f"Expected columns: Item, Target Date, Delivery Date, Quantity Ordered."
            )
            st.stop()
        except Exception as e:
            st.error(f"Something went wrong while processing the files: {e}")
            st.stop()

        st.success("Done.")

        # History context banner
        if history_file:
            st.info(
                f"📋 History loaded: {len(seen_pairs):,} previously seen SKU+date pair(s). "
                f"**{len(new_pairs):,}** new pair(s) found this run and added to history."
            )
        else:
            st.warning(
                "⚠️ No previous report uploaded — treating all aged SKUs as new. "
                "Download this report and upload it next time to enable deduplication."
            )

        if not unmatched_df.empty:
            st.warning(
                f"{unmatched_df['Item'].nunique()} item(s) on order have no matching SKU at all "
                f"in the short code file (counted as fully short). See the 'Unmatched Items' tab."
            )

        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Short SKUs", len(short_summary_df))
        col2.metric("Email/Research items+locations", len(email_summary_df))
        col3.metric("New SKU TO CODE rows", len(sku_df))
        col4.metric("Total history pairs", len(history_df))

        excel_bytes = to_excel_bytes(short_summary_df, email_summary_df, sku_df, history_df)
        st.download_button(
            "⬇️ Download results (Excel)",
            data=excel_bytes,
            file_name=f"shortage_analysis_{datetime.now().strftime('%Y-%m-%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

        def date_only(df, cols):
            df = df.copy()
            for c in cols:
                if c in df.columns:
                    df[c] = pd.to_datetime(df[c], errors="coerce").dt.date
            return df

        tab1, tab2, tab3, tab4, tab5 = st.tabs(
            ["SHORT SHEET", "EMAIL / RESEARCH", "SKU TO CODE (new)", "SKU HISTORY", "Unmatched Items"]
        )
        with tab1:
            st.dataframe(date_only(short_summary_df, ["Earliest Target Date"]), use_container_width=True)
        with tab2:
            st.dataframe(date_only(email_summary_df, ["Earliest Delivery Date"]), use_container_width=True)
        with tab3:
            if sku_df.empty:
                st.info("No new SKU TO CODE rows this run — all aged SKUs were already in history.")
            else:
                st.dataframe(date_only(sku_df, ["Consumer Priority Date"]), use_container_width=True)
        with tab4:
            st.caption("Every SKU + Consumer Priority Date pair ever flagged across all runs.")
            st.dataframe(date_only(history_df, ["Consumer Priority Date"]), use_container_width=True)
        with tab5:
            st.dataframe(date_only(unmatched_df, ["Target Date"]), use_container_width=True)
else:
    st.info("Upload both files in the sidebar, then click **Run Analysis**.")
