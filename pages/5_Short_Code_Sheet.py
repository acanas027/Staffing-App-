# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="Inventory Shortage Checker", layout="wide")
st.title("Inventory Shortage Checker")
st.caption("Upload the orders file and the warehouse short code file to find shortages, expiration-date research items, and LPNs that need to be coded.")

with st.expander("How this works"):
    st.markdown(
        "- Items starting with **S** are removed from the orders file.\n"
        "- **Whole-number items in the orders file are excluded from shortage matching.**\n"
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
        "- **SKU TO CODE**: every LPN in the qPORT/WMS file with Consumer Priority Date on or before today. "
        "This sheet is LPN-level and is not limited to SKUs from the Data/orders file.\n"
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
    today = datetime.now().date()
    sku_to_code_cutoff_date = today
    st.caption(f"SKU TO CODE cutoff: Consumer Priority Date on or before {sku_to_code_cutoff_date}")
    st.divider()
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


def normalize_location(x):
    """Clean location text so the same location is shown consistently.
    Example: ' rc2 a30 x8 ' -> 'RC2A30X8'.
    """
    if pd.isna(x):
        return ""
    s = str(x).strip().upper()
    # remove common separators/spaces created by Excel exports or manual typing
    for ch in [" ", "-", "_", "."]:
        s = s.replace(ch, "")
    if s in {"", "NAN", "NONE", "NULL"}:
        return ""
    return s


def location_zone(normalized_location):
    loc = normalize_location(normalized_location)
    return loc[:3] if loc else ""


def location_detail(normalized_location):
    loc = normalize_location(normalized_location)
    return loc[3:] if len(loc) > 3 else ""


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

    location = raw.iloc[:, 0:6].apply(lambda col: col.map(cell_to_str)).agg("".join, axis=1)
    previous_location = raw.iloc[:, 10:16].apply(lambda col: col.map(cell_to_str)).agg("".join, axis=1)
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
    out["Normalized Location"] = out["Location"].apply(normalize_location)
    out["Location Zone"] = out["Normalized Location"].apply(location_zone)
    out["Location Detail"] = out["Normalized Location"].apply(location_detail)
    out["Previous Normalized Location"] = out["Previous Location"].apply(normalize_location)
    out["Previous Location Zone"] = out["Previous Normalized Location"].apply(location_zone)
    out["Previous Location Detail"] = out["Previous Normalized Location"].apply(location_detail)

    out = out[out["SKU Number"].notna()].copy()
    # Keep every WMS SKU here so SKU TO CODE can list every eligible LPN in qPORT.
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


def fmt_qty(x):
    """Format quantities cleanly for readable date breakdown text."""
    try:
        value = float(x)
    except Exception:
        return str(x)
    if value.is_integer():
        return str(int(value))
    return f"{value:g}"


def format_wms_date_qty_breakdown(candidate_remaining):
    """Return all remaining WMS dates and quantities for the SKU at the time of the order-line check."""
    if candidate_remaining.empty:
        return ""

    work = candidate_remaining[["Consumer Priority Date", "Remaining Quantity"]].copy()
    work["Sort Date"] = pd.to_datetime(work["Consumer Priority Date"], errors="coerce")
    work["Date Label"] = work["Sort Date"].apply(
        lambda d: "No WMS date" if pd.isna(d) else str(pd.Timestamp(d).date())
    )
    work = work.sort_values(["Sort Date", "Date Label"], na_position="last")

    rows = []
    for label, g in work.groupby("Date Label", sort=False):
        rows.append(f"{label}: {fmt_qty(g['Remaining Quantity'].sum())}")
    return "; ".join(rows)


WMS_COMPARE_COL = "WMS Date Compared to Customer Target Date"


def prepare_short_sheet_output(short_summary_df):
    """Friendly SHORT SHEET view for Excel download and Streamlit display."""
    columns = [
        "Item",
        "Total Short By",
        "Product Short Cases",
        "Date Short Cases",
        "Customer Target Date",
        WMS_COMPARE_COL,
        "Short By Days",
        "Earliest WMS Date",
        "Latest WMS Date",
        "Partial Locations",
    ]
    if short_summary_df.empty:
        return pd.DataFrame(columns=columns)

    out = short_summary_df.copy()

    # Customer Target Date stays user-facing. Earliest Target Date stays internal for history/dedup.
    if "Customer Target Date" not in out.columns and "Earliest Target Date" in out.columns:
        out["Customer Target Date"] = out["Earliest Target Date"]

    # Backward-compatible rename if an older history/run used the prior column name.
    old_col = "Date Compared to Customer Target"
    if WMS_COMPARE_COL not in out.columns and old_col in out.columns:
        out[WMS_COMPARE_COL] = out[old_col]
    if WMS_COMPARE_COL not in out.columns:
        out[WMS_COMPARE_COL] = pd.NaT

    for qty_col in ["Product Short Cases", "Date Short Cases"]:
        if qty_col not in out.columns:
            out[qty_col] = 0

    if "Earliest WMS Date" not in out.columns:
        out["Earliest WMS Date"] = pd.NaT
    if "Latest WMS Date" not in out.columns:
        out["Latest WMS Date"] = pd.NaT

    if "Short By Days" not in out.columns:
        out["Short By Days"] = out.apply(
            lambda row: short_by_days(
                row.get("Customer Target Date"),
                row.get(WMS_COMPARE_COL),
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
            # No SKU match in the WMS/short-code file.
            # Keep this OUT of SHORT SHEET because it is not a true date/quantity shortage yet;
            # it needs to be reviewed separately in the Unmatched Items tab.
            unmatched_rows.append({
                "Item": order["Item"],
                "Quantity Needed": need,
                "Target Date": order["Target Date"],
                "Customer": order.get("Customer", ""),
                "Order": order.get("Order", ""),
            })
            continue

        # Look at remaining inventory for this SKU before allocating this order line.
        # This lets us split the shortage into:
        # 1) product shortage = not enough cases exist at all, regardless of date
        # 2) date shortage = cases exist, but their WMS Consumer Priority Date is before the customer target date
        candidate_remaining = candidates.copy()
        candidate_remaining["Remaining Quantity"] = candidate_remaining["wms_row_id"].map(remaining).fillna(0).astype(float)
        candidate_remaining = candidate_remaining[candidate_remaining["Remaining Quantity"] > 0].copy()

        total_remaining_all_dates = candidate_remaining["Remaining Quantity"].sum()
        good_date_cases_available = candidate_remaining.loc[
            candidate_remaining["Consumer Priority Date"] >= target_date,
            "Remaining Quantity"
        ].sum()
        bad_date_cases_available = candidate_remaining.loc[
            candidate_remaining["Consumer Priority Date"] < target_date,
            "Remaining Quantity"
        ].sum()
        no_wms_date_cases_available = candidate_remaining.loc[
            candidate_remaining["Consumer Priority Date"].isna(),
            "Remaining Quantity"
        ].sum()
        qualifying_remaining = good_date_cases_available

        earliest_wms_date = candidate_remaining["Consumer Priority Date"].dropna().min()
        latest_wms_date = candidate_remaining["Consumer Priority Date"].dropna().max()

        failing_dates = candidate_remaining.loc[
            candidate_remaining["Consumer Priority Date"] < target_date,
            "Consumer Priority Date"
        ].dropna()
        closest_failing_wms_date = failing_dates.max() if not failing_dates.empty else pd.NaT
        wms_date_qty_breakdown = format_wms_date_qty_breakdown(candidate_remaining)

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
            product_short_cases = min(short_qty, max(need - total_remaining_all_dates, 0))
            date_short_cases = max(short_qty - product_short_cases, 0)

            # For the single visual date comparison:
            # - If there is a date shortage, show the closest WMS date that failed the target.
            # - Otherwise, show the latest WMS date available for that SKU.
            # The case split above still uses ALL WMS dates and quantities.
            if date_short_cases > 0 and pd.notna(closest_failing_wms_date):
                wms_comparison_date = closest_failing_wms_date
            else:
                wms_comparison_date = latest_wms_date

            partial_locations = ", ".join(
                sorted(set(s["Location"] for s in sources))
            ) if sources else ""
            short_rows.append({
                "Item": order["Item"],
                "Quantity Needed": need,
                "Quantity Available (qualifying)": allocated_total,
                "Short By": short_qty,
                "Product Short Cases": product_short_cases,
                "Date Short Cases": date_short_cases,
                "Target Date": order["Target Date"],
                "Customer": order.get("Customer", ""),
                "Order": order.get("Order", ""),
                WMS_COMPARE_COL: wms_comparison_date,
                "Short By Days": short_by_days(order["Target Date"], wms_comparison_date),
                "Earliest WMS Date": earliest_wms_date,
                "Latest WMS Date": latest_wms_date,
                "Total Remaining All WMS Dates": total_remaining_all_dates,
                "Good-Date Cases Available": good_date_cases_available,
                "Bad-Date Cases Available": bad_date_cases_available,
                "No WMS Date Cases Available": no_wms_date_cases_available,
                "Remaining Qty Meeting Customer Target Date": qualifying_remaining,
                "WMS Date Qty Breakdown": wms_date_qty_breakdown,
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
    - Customer Target Date = earliest customer target date for that matched short SKU
    - Product Short Cases = cases missing because not enough cases exist in WMS at all
    - Date Short Cases = cases existing in WMS but not qualifying because the WMS date is before the customer target date
    - WMS Date Compared to Customer Target Date = closest failing WMS date when date-short, otherwise latest WMS date
    - Short By Days = how many days that comparison date is before the customer target date

    Items with no WMS SKU match are excluded from this sheet and sent to Unmatched Items.
    """
    base_columns = [
        "Item", "Total Short By", "Product Short Cases", "Date Short Cases",
        "Earliest Target Date", "Customer Target Date", WMS_COMPARE_COL,
        "Short By Days", "Earliest WMS Date", "Latest WMS Date", "Partial Locations"
    ]
    if short_df.empty:
        return pd.DataFrame(columns=base_columns), set()

    rows = []
    for item, g in short_df.groupby("Item", dropna=False):
        # Collect all non-empty locations across order lines for this SKU
        all_locs = set()
        if "Partial Locations" in g.columns:
            for loc_str in g["Partial Locations"].fillna(""):
                if loc_str:
                    for loc in str(loc_str).split(", "):
                        if loc.strip():
                            all_locs.add(loc.strip())

        customer_target_date = pd.to_datetime(g["Target Date"], errors="coerce").min()
        comparison_series = pd.to_datetime(
            g[WMS_COMPARE_COL], errors="coerce"
        ) if WMS_COMPARE_COL in g.columns else pd.Series(dtype="datetime64[ns]")

        # Prefer the comparison date from the row with the largest day miss, so the summary shows
        # the most urgent date issue. If there is no day miss, use the latest WMS date for visual reference.
        if "Short By Days" in g.columns and g["Short By Days"].notna().any():
            idx = g["Short By Days"].fillna(-1).astype(float).idxmax()
            comparison_date = pd.to_datetime(g.loc[idx, WMS_COMPARE_COL], errors="coerce") if WMS_COMPARE_COL in g.columns else pd.NaT
            max_short_by_days = g["Short By Days"].max()
        else:
            comparison_date = comparison_series.max() if not comparison_series.empty else pd.NaT
            max_short_by_days = short_by_days(customer_target_date, comparison_date)

        earliest_wms_date = pd.to_datetime(g["Earliest WMS Date"], errors="coerce").min() if "Earliest WMS Date" in g.columns else pd.NaT
        latest_wms_date = pd.to_datetime(g["Latest WMS Date"], errors="coerce").max() if "Latest WMS Date" in g.columns else pd.NaT

        rows.append({
            "Item": item,
            "Total Short By": g["Short By"].sum(),
            "Product Short Cases": g["Product Short Cases"].sum() if "Product Short Cases" in g.columns else 0,
            "Date Short Cases": g["Date Short Cases"].sum() if "Date Short Cases" in g.columns else 0,
            "Earliest Target Date": customer_target_date,  # internal dedup/history key
            "Customer Target Date": customer_target_date,  # customer date you need to meet
            WMS_COMPARE_COL: comparison_date,
            "Short By Days": max_short_by_days,
            "Earliest WMS Date": earliest_wms_date,
            "Latest WMS Date": latest_wms_date,
            "Partial Locations": ", ".join(sorted(all_locs)) if all_locs else "None",
        })

    summary = pd.DataFrame(rows, columns=base_columns)
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


def prepare_short_detail_output(short_df):
    """Detailed shortage view: one row per short order line, including all WMS date/qty buckets."""
    columns = [
        "Item", "Order", "Customer", "Quantity Needed", "Quantity Available (qualifying)",
        "Short By", "Product Short Cases", "Date Short Cases", "Customer Target Date",
        WMS_COMPARE_COL, "Short By Days", "Earliest WMS Date", "Latest WMS Date",
        "Total Remaining All WMS Dates", "Good-Date Cases Available",
        "Bad-Date Cases Available", "No WMS Date Cases Available",
        "Remaining Qty Meeting Customer Target Date",
        "WMS Date Qty Breakdown", "Partial Locations",
    ]
    if short_df.empty:
        return pd.DataFrame(columns=columns)

    out = short_df.copy()
    if "Target Date" in out.columns:
        out["Customer Target Date"] = out["Target Date"]
    for col in columns:
        if col not in out.columns:
            out[col] = None
    return out[columns].sort_values(["Customer Target Date", "Item", "Order"]).reset_index(drop=True)


def build_sku_to_code(wms_df, cutoff):
    """Return every qPORT/WMS LPN with Consumer Priority Date on or before cutoff.

    Important logic:
    - This is LPN-level: one row per LPN/pallet/date.
    - This is NOT limited to SKUs from the Data/orders file.
    - A row qualifies only when Consumer Priority Date is today or earlier.
    """
    cutoff_ts = pd.Timestamp(cutoff)
    aged = wms_df[wms_df["Consumer Priority Date"] <= cutoff_ts].copy()

    columns = [
        "LPN #", "SKU Number", "Quantity on LPN",
        "Consumer Priority Date", "Today Cutoff Date", "Days Past Today",
        "Location", "Normalized Location", "Location Zone", "Location Detail",
        "Previous Location", "Previous Normalized Location", "Previous Location Zone",
        "Previous Location Detail", "Code Reason",
    ]

    if aged.empty:
        return pd.DataFrame(columns=columns)

    aged["Quantity on LPN"] = aged["Quantity"]
    aged["Today Cutoff Date"] = cutoff_ts
    aged["Days Past Today"] = (
        cutoff_ts.normalize() - pd.to_datetime(aged["Consumer Priority Date"], errors="coerce").dt.normalize()
    ).dt.days
    aged["Code Reason"] = "LPN Consumer Priority Date is today or earlier"

    # Ensure normalized location fields exist even if an older edited file is used.
    if "Normalized Location" not in aged.columns:
        aged["Normalized Location"] = aged["Location"].apply(normalize_location)
    if "Location Zone" not in aged.columns:
        aged["Location Zone"] = aged["Normalized Location"].apply(location_zone)
    if "Location Detail" not in aged.columns:
        aged["Location Detail"] = aged["Normalized Location"].apply(location_detail)
    if "Previous Normalized Location" not in aged.columns:
        aged["Previous Normalized Location"] = aged["Previous Location"].apply(normalize_location)
    if "Previous Location Zone" not in aged.columns:
        aged["Previous Location Zone"] = aged["Previous Normalized Location"].apply(location_zone)
    if "Previous Location Detail" not in aged.columns:
        aged["Previous Location Detail"] = aged["Previous Normalized Location"].apply(location_detail)

    return aged[columns].sort_values(
        ["Consumer Priority Date", "SKU Number", "LPN #"]
    ).reset_index(drop=True)

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
    columns = [
        "Item", "Location Zone", "Total Quantity Needed", "Earliest Delivery Date",
        "Previous Locations", "Orders Affected"
    ]
    if email_df.empty:
        return pd.DataFrame(columns=columns)

    email_df = email_df.copy()

    if date_range is not None and len(date_range) == 2:
        start_ts = pd.Timestamp(date_range[0])
        end_ts = pd.Timestamp(date_range[1])
        email_df = email_df[
            (email_df["Delivery Date"] >= start_ts) & (email_df["Delivery Date"] <= end_ts)
        ]

    if email_df.empty:
        return pd.DataFrame(columns=columns)

    email_df["Location Zone"] = email_df["Location"].astype(str).str[:3]

    rows = []
    for (item, zone), g in email_df.groupby(["Item", "Location Zone"], dropna=False):
        rows.append({
            "Item": item,
            "Location Zone": zone,
            "Total Quantity Needed": g["Quantity from this location"].sum(),
            "Earliest Delivery Date": pd.to_datetime(g["Delivery Date"], errors="coerce").min(),
            "Previous Locations": ", ".join(sorted(set(g["Previous Location"].astype(str)))),
            "Orders Affected": g["Order"].nunique(),
        })

    summary = pd.DataFrame(rows, columns=columns)
    summary = summary.sort_values(["Item", "Earliest Delivery Date"]).reset_index(drop=True)
    return summary

def to_excel_bytes(short_summary_df, email_summary_df, sku_df, short_history_df, short_detail_df=None):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # SHORT SHEET (new only)
        short_out = prepare_short_sheet_output(short_summary_df)
        if not short_out.empty:
            for date_col in ["Customer Target Date", WMS_COMPARE_COL, "Earliest WMS Date", "Latest WMS Date"]:
                if date_col in short_out.columns:
                    short_out[date_col] = pd.to_datetime(
                        short_out[date_col], errors="coerce"
                    ).dt.date
        short_out.to_excel(writer, sheet_name="SHORT SHEET", index=False)

        # SHORT DETAIL (one row per short order line, including all WMS date/qty buckets)
        detail_out = prepare_short_detail_output(short_detail_df if short_detail_df is not None else pd.DataFrame())
        if not detail_out.empty:
            for date_col in ["Customer Target Date", WMS_COMPARE_COL, "Earliest WMS Date", "Latest WMS Date"]:
                if date_col in detail_out.columns:
                    detail_out[date_col] = pd.to_datetime(detail_out[date_col], errors="coerce").dt.date
        detail_out.to_excel(writer, sheet_name="SHORT DETAIL", index=False)

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

        # SKU TO CODE (LPN-level: one row per LPN/pallet/date)
        sku_out = sku_df.copy()
        if not sku_out.empty:
            for date_col in ["Consumer Priority Date", "Today Cutoff Date"]:
                if date_col in sku_out.columns:
                    sku_out[date_col] = pd.to_datetime(
                        sku_out[date_col], errors="coerce"
                    ).dt.date
        (sku_out if not sku_out.empty else pd.DataFrame(columns=[
            "LPN #", "SKU Number", "Quantity on LPN",
            "Consumer Priority Date", "Today Cutoff Date", "Days Past Today",
            "Location", "Normalized Location", "Location Zone", "Location Detail",
            "Previous Location", "Previous Normalized Location", "Previous Location Zone",
            "Previous Location Detail", "Code Reason",
        ])).to_excel(writer, sheet_name="SKU TO CODE", index=False)

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
                sku_df = build_sku_to_code(wms_df, sku_to_code_cutoff_date)

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
                f"History loaded - "
                f"SHORT SHEET: {len(seen_short_pairs):,} previously seen pair(s), "
                f"**{len(new_short_pairs):,}** new this run."
            )
        else:
            st.warning(
                "No previous report uploaded - all short items treated as new. "
                "Download this report and upload it next time to enable SHORT SHEET deduplication."
            )

        if not unmatched_df.empty:
            unmatched_count = unmatched_df["Item"].nunique()
            st.warning(
                f"{unmatched_count} item(s) on order have no matching SKU at all "
                f"in the short code file. They are listed only in the 'Unmatched Items' tab, "
                f"not in SHORT SHEET."
            )

        col1, col2, col3 = st.columns(3)
        col1.metric("New Short SKUs", len(short_summary_df))
        col2.metric("Email/Research items+locations", len(email_summary_df))
        col3.metric("SKU TO CODE LPN rows", len(sku_df))

        excel_bytes = to_excel_bytes(
            short_summary_df, email_summary_df, sku_df, short_history_df, short_df
        )
        st.download_button(
            "Download results (Excel)",
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

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
            "SHORT SHEET", "SHORT DETAIL", "EMAIL / RESEARCH", "SKU TO CODE",
            "SHORT HISTORY", "Unmatched Items"
        ])
        with tab1:
            if short_summary_df.empty:
                st.info("No new shortages this run — all were already in SHORT HISTORY.")
            else:
                short_display_df = prepare_short_sheet_output(short_summary_df)
                st.dataframe(
                    date_only(short_display_df, ["Customer Target Date", WMS_COMPARE_COL, "Earliest WMS Date", "Latest WMS Date"]),
                    use_container_width=True
                )
        with tab2:
            detail_display_df = prepare_short_detail_output(short_df)
            st.dataframe(
                date_only(detail_display_df, ["Customer Target Date", WMS_COMPARE_COL, "Earliest WMS Date", "Latest WMS Date"]),
                use_container_width=True
            )
        with tab3:
            st.dataframe(date_only(email_summary_df, ["Earliest Delivery Date"]), use_container_width=True)
        with tab4:
            if sku_df.empty:
                st.info("No aged SKU TO CODE rows found for this cutoff date.")
            else:
                st.dataframe(date_only(sku_df, ["Consumer Priority Date", "Today Cutoff Date"]), use_container_width=True)
        with tab5:
            st.caption("Every Item + Earliest Target Date pair ever flagged as short across all runs.")
            st.dataframe(date_only(short_history_df, ["Earliest Target Date"]), use_container_width=True)
        with tab6:
            st.dataframe(date_only(unmatched_df, ["Target Date"]), use_container_width=True)
else:
    st.info("Upload both files in the sidebar, then click **Run Analysis**.")
