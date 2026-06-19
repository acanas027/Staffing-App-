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
        "- **Whole-number SKUs (no decimal) are excluded from all output sheets.**\n"
        "- The full SKU Number is built from the short code file's columns H and I, "
        "then normalized and matched against the Item code so leading zeros/trailing decimal zeros do not create false shortages.\n"
        "- Order lines are processed most-urgent-Target-Date first, allocating "
        "inventory whose Consumer Priority Date is on or after the Target Date, "
        "soonest-qualifying-date first. Inventory is not double-counted across orders.\n"
        "- **SHORT SHEET**: new shortages only — one row per Item/SKU+Target Date not "
        "seen in a previous run. Dedup key is SKU + Earliest Target Date.\n"
        "- **EMAIL / RESEARCH**: one row per Item + Location Zone for fully-covered "
        "order lines where that zone does not start with RC2 and the "
        "Delivery Date falls within the chosen departure window.\n"
        "- **SKU TO CODE**: all aged inventory (Consumer Priority Date ≤ cutoff), "
        "one row per pallet with Location and LPN #. Always shows everything — "
        "no dedup, so missed items from a previous run will reappear.\n"
        "- **SHORT HISTORY**: cumulative record of every Item + Earliest Target Date "
        "pair ever flagged as short. Upload last run's Excel to carry it forward."
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
             "'SHORT HISTORY' sheet to skip already-reported shortages."
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


def normalize_sku(x):
    """Consistent string form for dedup comparisons.
    - Strips leading zeros from the whole part (Excel drops them on read-back)
    - Strips trailing zeros from the decimal part (Excel drops them on read-back)
    - Handles pandas trailing '.0' artifact for whole-number floats
    So '06795.48940', '6795.4894', and '6795.48940' all normalize to '6795.4894'."""
    s = str(x).strip()
    # Whole-number float artifact from pandas: '79341.0' -> '79341'
    if s.endswith(".0") and s[:-2].isdigit():
        return s[:-2]
    if "." in s:
        whole, dec = s.split(".", 1)
        # Strip leading zeros from whole part
        whole = str(int(whole)) if whole.isdigit() else whole
        # Strip trailing zeros from decimal part
        dec = dec.rstrip("0")
        return f"{whole}.{dec}" if dec else whole
    return s


def has_decimal(sku_str):
    """Return True only if the SKU contains a real decimal portion (not .0)."""
    s = normalize_sku(sku_str)
    return "." in s


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
    # Drop whole-number items (no decimal)
    df = df[df["Item"].str.contains(r"\.", regex=True)].copy()
    df["Target Date"] = pd.to_datetime(df["Target Date"])
    df["Delivery Date"] = pd.to_datetime(df["Delivery Date"])
    df["Quantity Ordered"] = pd.to_numeric(df["Quantity Ordered"], errors="coerce").fillna(0)
    df["match_key"] = df["Item"].apply(normalize_sku)
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
    # Drop whole-number SKUs (no decimal)
    out = out[out["SKU Number"].apply(has_decimal)].copy()
    out["match_key"] = out["SKU Number"].apply(normalize_sku)
    out = out.reset_index(drop=True)
    out["wms_row_id"] = out.index
    return out



def load_short_history(file):
    """Read SHORT HISTORY sheet. Returns set of (sku, target_date_str) tuples."""
    if file is None:
        return set()
    try:
        hist_df = pd.read_excel(file, sheet_name="SHORT HISTORY")
        hist_df["Earliest Target Date"] = pd.to_datetime(
            hist_df["Earliest Target Date"], errors="coerce"
        )
        seen = set()
        for _, row in hist_df.iterrows():
            sku = normalize_sku(row["Item"])
            dt = row["Earliest Target Date"]
            if pd.notna(dt):
                seen.add((sku, str(dt.date())))
        return seen
    except Exception:
        return set()

def days_until_or_past_target(target_date):
    """Positive = days remaining until target date. Negative = days past target date."""
    if pd.isna(target_date):
        return None
    target = pd.Timestamp(target_date).date()
    today = datetime.now().date()
    return (target - today).days


def short_by_days(customer_target_date, comparison_date):
    """How many days the compared WMS date misses the customer target date.
    Positive = WMS date is before target date. Zero = date meets/exceeds target or no date gap.
    Blank = no comparison date available.
    """
    if pd.isna(customer_target_date) or pd.isna(comparison_date):
        return None
    customer_target = pd.Timestamp(customer_target_date).date()
    compared = pd.Timestamp(comparison_date).date()
    return max((customer_target - compared).days, 0)


def prepare_short_sheet_output(short_summary_df):
    """Friendly SHORT SHEET view for Excel download and Streamlit display."""
    columns = [
        "Item", "Total Short By", "Customer Target Date",
        "Date Compared to Customer Target", "Short By Days", "Partial Locations"
    ]
    if short_summary_df.empty:
        return pd.DataFrame(columns=columns)

    out = short_summary_df.copy()

    # Keep the user-facing customer target date and the WMS date being compared to it.
    # Earliest Target Date stays internal for SHORT HISTORY deduplication only.
    if "Customer Target Date" not in out.columns and "Earliest Target Date" in out.columns:
        out["Customer Target Date"] = out["Earliest Target Date"]
    if "Date Compared to Customer Target" not in out.columns:
        out["Date Compared to Customer Target"] = pd.NaT
    if "Short By Days" not in out.columns:
        out["Short By Days"] = out.apply(
            lambda row: short_by_days(
                row.get("Customer Target Date"),
                row.get("Date Compared to Customer Target")
            ),
            axis=1
        )

    return out[columns]



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
                "Date Compared to Customer Target": pd.NaT,
                "Partial Locations": "",
            })
            continue

        comparison_date = candidates["Consumer Priority Date"].dropna().max()

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
            partial_locations = ", ".join(
                sorted(set(s["Location"] for s in sources))
            ) if sources else ""
            short_rows.append({
                "Item": order["Item"],
                "Quantity Needed": need,
                "Quantity Available (qualifying)": allocated_total,
                "Short By": short_qty,
                "Target Date": order["Target Date"],
                "Customer": order.get("Customer", ""),
                "Order": order.get("Order", ""),
                "Date Compared to Customer Target": comparison_date,
                "Partial Locations": partial_locations,
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


def build_short_summary(short_df, seen_short_pairs):
    """Collapse to one row per Item, dedup against SHORT HISTORY.
    Dedup key: SKU + Earliest Target Date.

    User-facing SHORT SHEET uses:
    - Customer Target Date = earliest customer target date for that short SKU
    - Date Compared to Customer Target = best/latest WMS Consumer Priority Date available for that SKU
    - Short By Days = how many days the WMS date is before the customer target date
    """
    if short_df.empty:
        return pd.DataFrame(columns=[
            "Item", "Total Short By", "Earliest Target Date",
            "Customer Target Date", "Date Compared to Customer Target",
            "Short By Days", "Partial Locations"
        ]), set()

    def agg_group(g):
        # Collect all non-empty locations across order lines for this SKU
        all_locs = set()
        for loc_str in g["Partial Locations"]:
            if loc_str:
                for loc in loc_str.split(", "):
                    if loc.strip():
                        all_locs.add(loc.strip())

        customer_target_date = g["Target Date"].min()
        comparison_date = pd.to_datetime(
            g.get("Date Compared to Customer Target"), errors="coerce"
        ).max()

        return pd.Series({
            "Total Short By": g["Short By"].sum(),
            "Earliest Target Date": customer_target_date,  # internal dedup/history key
            "Customer Target Date": customer_target_date,  # customer date you need to meet
            "Date Compared to Customer Target": comparison_date,  # WMS Consumer Priority Date used for visual comparison
            "Short By Days": short_by_days(customer_target_date, comparison_date),
            "Partial Locations": ", ".join(sorted(all_locs)) if all_locs else "None",
        })

    summary = short_df.groupby("Item", as_index=False).apply(agg_group, include_groups=False)
    summary = summary.sort_values("Earliest Target Date").reset_index(drop=True)

    def is_new(row):
        sku = normalize_sku(row["Item"])
        dt = row["Earliest Target Date"]
        if pd.isna(dt):
            return False
        return (sku, str(pd.Timestamp(dt).date())) not in seen_short_pairs

    summary["_is_new"] = summary.apply(is_new, axis=1)
    new_summary = summary[summary["_is_new"]].drop(columns=["_is_new"]).reset_index(drop=True)

    new_short_pairs = set()
    for _, row in new_summary.iterrows():
        dt = row["Earliest Target Date"]
        if pd.notna(dt):
            new_short_pairs.add((normalize_sku(row["Item"]), str(pd.Timestamp(dt).date())))

    return new_summary, new_short_pairs


def build_sku_to_code(wms_df, cutoff):
    """Return all aged WMS rows (Consumer Priority Date <= cutoff).
    One row per pallet, including Location and LPN #. No dedup — always shows everything."""
    cutoff_ts = pd.Timestamp(cutoff)
    aged = wms_df[wms_df["Consumer Priority Date"] <= cutoff_ts].copy()
    aged = aged[
        ["SKU Number", "LPN #", "Location", "Quantity", "Consumer Priority Date"]
    ].sort_values(["Consumer Priority Date", "SKU Number", "Location"]).reset_index(drop=True)
    return aged


def build_updated_history(seen_pairs, new_pairs, key_cols):
    """Merge old + new pairs into a history DataFrame with given column names."""
    all_pairs = seen_pairs | new_pairs
    rows = [{key_cols[0]: sku, key_cols[1]: date_str} for sku, date_str in sorted(all_pairs)]
    hist_df = pd.DataFrame(rows) if rows else pd.DataFrame(columns=key_cols)
    if not hist_df.empty:
        hist_df[key_cols[1]] = pd.to_datetime(hist_df[key_cols[1]], errors="coerce")
        hist_df = hist_df.sort_values(key_cols[::-1]).reset_index(drop=True)
    return hist_df


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


def to_excel_bytes(short_summary_df, email_summary_df, sku_df, short_history_df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # SHORT SHEET (new only)
        short_out = prepare_short_sheet_output(short_summary_df)
        if not short_out.empty:
            for date_col in ["Customer Target Date", "Date Compared to Customer Target"]:
                if date_col in short_out.columns:
                    short_out[date_col] = pd.to_datetime(
                        short_out[date_col], errors="coerce"
                    ).dt.date
        short_out.to_excel(writer, sheet_name="SHORT SHEET", index=False)

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

        # SHORT HISTORY (cumulative)
        short_hist_out = short_history_df.copy()
        if not short_hist_out.empty:
            short_hist_out["Earliest Target Date"] = pd.to_datetime(
                short_hist_out["Earliest Target Date"]
            ).dt.date
        (short_hist_out if not short_hist_out.empty else pd.DataFrame(
            columns=["Item", "Earliest Target Date"]
        )).to_excel(writer, sheet_name="SHORT HISTORY", index=False)

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

                seen_short_pairs = load_short_history(history_file)

                short_df, email_df, unmatched_df = run_allocation(orders_df, wms_df)

                short_summary_df, new_short_pairs = build_short_summary(short_df, seen_short_pairs)
                sku_df = build_sku_to_code(wms_df, cutoff_date)

                short_history_df = build_updated_history(
                    seen_short_pairs, new_short_pairs,
                    ["Item", "Earliest Target Date"]
                )

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

        if history_file:
            st.info(
                f"📋 History loaded — "
                f"SHORT SHEET: {len(seen_short_pairs):,} previously seen pair(s), "
                f"**{len(new_short_pairs):,}** new this run."
            )
        else:
            st.warning(
                "⚠️ No previous report uploaded — all short items treated as new. "
                "Download this report and upload it next time to enable SHORT SHEET deduplication."
            )

        if not unmatched_df.empty:
            st.warning(
                f"{unmatched_df['Item'].nunique()} item(s) on order have no matching SKU at all "
                f"in the short code file (counted as fully short). See the 'Unmatched Items' tab."
            )

        col1, col2, col3 = st.columns(3)
        col1.metric("New Short SKUs", len(short_summary_df))
        col2.metric("Email/Research items+locations", len(email_summary_df))
        col3.metric("SKU TO CODE rows", len(sku_df))

        excel_bytes = to_excel_bytes(
            short_summary_df, email_summary_df, sku_df, short_history_df
        )
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

        tab1, tab2, tab3, tab4, tab5 = st.tabs([
            "SHORT SHEET", "EMAIL / RESEARCH", "SKU TO CODE",
            "SHORT HISTORY", "Unmatched Items"
        ])
        with tab1:
            if short_summary_df.empty:
                st.info("No new shortages this run — all were already in SHORT HISTORY.")
            else:
                short_display_df = prepare_short_sheet_output(short_summary_df)
                st.dataframe(
                    date_only(short_display_df, ["Customer Target Date", "Date Compared to Customer Target"]),
                    use_container_width=True
                )
        with tab2:
            st.dataframe(date_only(email_summary_df, ["Earliest Delivery Date"]), use_container_width=True)
        with tab3:
            if sku_df.empty:
                st.info("No aged SKU TO CODE rows found for this cutoff date.")
            else:
                st.dataframe(date_only(sku_df, ["Consumer Priority Date"]), use_container_width=True)
        with tab4:
            st.caption("Every Item + Earliest Target Date pair ever flagged as short across all runs.")
            st.dataframe(date_only(short_history_df, ["Earliest Target Date"]), use_container_width=True)
        with tab5:
            st.dataframe(date_only(unmatched_df, ["Target Date"]), use_container_width=True)
else:
    st.info("Upload both files in the sidebar, then click **Run Analysis**.")
