import re

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Expected Pallets by Trailer", layout="wide")

st.title("Expected Pallets by Trailer")

st.write(
    "Upload your inbound report. Each LPN counts as one pallet. "
)

# Inbound report column positions, 0-indexed:
# C,D,E,F,G = 2,3,4,5,6
# LPN # = 7, column H
TRAILER_COLS = [2, 3, 4, 5, 6]
LPN_COL = 7

HEADER_ROWS = 3       # first 3 rows are headers
THRESHOLD = 9         # loads with 9 or fewer pallets are highlighted red


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


def highlight_short_loads(row):
    """Highlight loads with 9 or fewer pallets."""
    if row["Expected Pallets"] <= THRESHOLD:
        return ["background-color: #ffb3b3; color: #800000; font-weight: bold"] * len(row)

    return [""] * len(row)


def filter_by_trailer(df, search_text):
    """Filter table by trailer last 3 or full trailer number."""
    if df.empty:
        return df

    search_text = str(search_text).strip()

    if search_text == "":
        return df

    search_digits = "".join(re.findall(r"\d", search_text))
    search_value = search_digits if search_digits else search_text.lower()

    mask = (
        df["Trailer Last 3"].astype(str).str.contains(search_value, case=False, na=False)
        | df["Full Trailer"].astype(str).str.contains(search_value, case=False, na=False)
    )

    return df[mask].copy()


def build_expected_pallets_table(uploaded_file):
    """
    Read the uploaded inbound report and return one row per trailer last 3.

    Each unique LPN counts as one pallet.
    If the same trailer last 3 appears more than once, the app keeps the highest pallet count.
    """
    df = pd.read_excel(uploaded_file, header=None, skiprows=HEADER_ROWS)

    # Drop blank/footer rows.
    df = df[df[LPN_COL].notna()].copy()

    df["Full Trailer"] = df.apply(build_trailer, axis=1)
    df = df[df["Full Trailer"] != ""].copy()

    # Each LPN = one pallet. nunique guards against duplicate LPN rows.
    trailer_result = (
        df.groupby("Full Trailer")[LPN_COL]
        .nunique()
        .reset_index()
        .rename(columns={LPN_COL: "Expected Pallets"})
    )

    trailer_result["Trailer Last 3"] = trailer_result["Full Trailer"].apply(get_last_3_numbers)
    trailer_result = trailer_result[trailer_result["Trailer Last 3"] != ""].copy()

    # If the same last 3 appears more than once, keep the highest pallet count.
    result = (
        trailer_result
        .sort_values("Expected Pallets", ascending=False)
        .groupby("Trailer Last 3", as_index=False)
        .agg(
            **{
                "Expected Pallets": ("Expected Pallets", "max"),
                "Full Trailer": ("Full Trailer", lambda x: ", ".join(sorted(set(str(v) for v in x)))),
            }
        )
        .sort_values("Expected Pallets", ascending=False)
        .reset_index(drop=True)
    )

    result["Status"] = result["Expected Pallets"].apply(
        lambda pallets: "Research" if pallets <= THRESHOLD else "OK"
    )

    return result[["Trailer Last 3", "Expected Pallets", "Full Trailer", "Status"]]


uploaded = st.file_uploader("Upload your inbound report", type=["xlsx", "xlsm"])

if uploaded is not None:
    try:
        result = build_expected_pallets_table(uploaded)

        c1, c2, c3 = st.columns(3)
        c1.metric("Total trailers", len(result))
        c2.metric(f"Loads over {THRESHOLD}", len(result[result["Expected Pallets"] > THRESHOLD]))
        c3.metric(f"Loads {THRESHOLD} or less", len(result[result["Expected Pallets"] <= THRESHOLD]))

        search_text = st.text_input(
            "Search trailer",
            placeholder="Type trailer last 3 or full trailer number",
        )

        filtered_result = filter_by_trailer(result, search_text)

        styled_result = filtered_result.style.apply(highlight_short_loads, axis=1).hide(axis="index")

        st.dataframe(styled_result, use_container_width=True)

    except Exception as e:
        st.error(f"Could not process the uploaded report: {e}")

else:
    st.info("Waiting for an inbound report.")
