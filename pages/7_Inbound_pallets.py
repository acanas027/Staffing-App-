import streamlit as st
import pandas as pd

st.set_page_config(page_title="Pallets per Trailer", layout="wide")
st.title("📦 Pallets per Trailer")
st.write(
    "Upload your inbound report. Each **LPN** counts as one pallet, and the "
    "**trailer number** is built from columns C, D, E, F and G combined."
)

uploaded = st.file_uploader("Upload your Excel file", type=["xlsx", "xlsm"])

# Column positions (0-indexed): C,D,E,F,G = 2,3,4,5,6 ; LPN # = 7 (column H)
TRAILER_COLS = [2, 3, 4, 5, 6]
LPN_COL = 7
HEADER_ROWS = 3  # first 3 rows are headers


def build_trailer(row):
    parts = [str(row[c]).strip() for c in TRAILER_COLS if pd.notna(row[c])]
    return "-".join(parts)


if uploaded is not None:
    # Headers span the first 3 rows, so read with no header and skip them.
    df = pd.read_excel(uploaded, header=None, skiprows=HEADER_ROWS)

    # Drop rows with no LPN (blank/footer rows)
    df = df[df[LPN_COL].notna()]

    df["Trailer"] = df.apply(build_trailer, axis=1)

    # Each LPN = one pallet. nunique guards against any duplicate LPN rows.
    result = (
        df.groupby("Trailer")[LPN_COL]
        .nunique()
        .reset_index()
        .rename(columns={LPN_COL: "Pallets"})
        .sort_values("Pallets", ascending=False)
        .reset_index(drop=True)
    )

    col1, col2 = st.columns(2)
    col1.metric("Total trailers", len(result))
    col2.metric("Total pallets", int(result["Pallets"].sum()))

    st.subheader("Pallets by trailer")
    st.dataframe(result, use_container_width=True, hide_index=True)

    st.bar_chart(result.set_index("Trailer")["Pallets"])

    csv = result.to_csv(index=False).encode("utf-8")
    st.download_button(
        "⬇️ Download results as CSV",
        data=csv,
        file_name="pallets_per_trailer.csv",
        mime="text/csv",
    )
else:
    st.info("Waiting for a file…")
