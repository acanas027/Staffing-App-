import streamlit as st
import pandas as pd

st.set_page_config(page_title="Inbound Pallets", layout="wide")
st.title("📦 Pallets per Trailer")
st.write(
    "Upload your inbound report. Each **LPN** counts as one pallet, and the "
    "**trailer number** is columns C, D, E, F and G combined."
)

uploaded = st.file_uploader("Upload your Excel file", type=["xlsx", "xlsm"])

# Column positions (0-indexed): C,D,E,F,G = 2,3,4,5,6 ; LPN # = 7 (column H)
TRAILER_COLS = [2, 3, 4, 5, 6]
LPN_COL = 7
HEADER_ROWS = 3       # first 3 rows are headers
THRESHOLD = 9         # loads with 9 or fewer pallets get flagged for research


def build_trailer(row):
    # Concatenate the trailer parts with no separator
    return "".join(str(row[c]).strip() for c in TRAILER_COLS if pd.notna(row[c]))


def flag_red(row):
    return ["background-color: #ffb3b3; color: #800000; font-weight: bold"] * len(row)


if uploaded is not None:
    # Headers span the first 3 rows, so read with no header and skip them.
    df = pd.read_excel(uploaded, header=None, skiprows=HEADER_ROWS)
    df = df[df[LPN_COL].notna()]          # drop blank/footer rows
    df["Trailer"] = df.apply(build_trailer, axis=1)

    # Each LPN = one pallet (nunique guards against any duplicate LPN rows)
    result = (
        df.groupby("Trailer")[LPN_COL]
        .nunique()
        .reset_index()
        .rename(columns={LPN_COL: "Pallets"})
        .sort_values("Pallets", ascending=False)
        .reset_index(drop=True)
    )

    high = result[result["Pallets"] > THRESHOLD].reset_index(drop=True)
    low = result[result["Pallets"] <= THRESHOLD].reset_index(drop=True)

    c1, c2, c3 = st.columns(3)
    c1.metric("Total trailers", len(result))
    c2.metric(f"Loads over {THRESHOLD}", len(high))
    c3.metric(f"Loads {THRESHOLD} or less", len(low))

    # ---- List 1: more than 9 pallets ----
    st.subheader(f"✅ Loads with more than {THRESHOLD} pallets")
    st.dataframe(high, use_container_width=True, hide_index=True)
    st.download_button(
        "⬇️ Download full loads (CSV)",
        data=high.to_csv(index=False).encode("utf-8"),
        file_name="loads_over_9.csv",
        mime="text/csv",
    )

    # ---- List 2: 9 or fewer pallets (flagged for research) ----
    st.subheader(f"🚨 Loads with {THRESHOLD} or fewer pallets")
    if low.empty:
        st.success("No short loads — nothing to research.")
    else:
        low_display = low.copy()
        low_display["Status"] = "research"      # note in the cell next to it
        styled = low_display.style.apply(flag_red, axis=1).hide(axis="index")
        st.dataframe(styled, use_container_width=True)
        st.download_button(
            "⬇️ Download research list (CSV)",
            data=low_display.to_csv(index=False).encode("utf-8"),
            file_name="loads_to_research.csv",
            mime="text/csv",
        )
else:
    st.info("Waiting for a file…")
