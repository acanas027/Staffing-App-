import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import Alignment
from io import BytesIO

# =====================================================================
# PAGE CONFIG
# =====================================================================
st.set_page_config(
    page_title="Shorts Analysis Tool",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =====================================================================
# DESIGN TOKENS + GLOBAL STYLE
# =====================================================================
NAVY = "#12233F"
STEEL = "#3E5C76"
AMBER = "#E8871E"
SUCCESS = "#2E9E5B"
WARNING = "#E8871E"
DANGER = "#D64545"
BG = "#F5F6F8"
CARD = "#FFFFFF"
TEXT_MUTED = "#6B7280"

STATUS_COLORS = {
    "Full ✅": SUCCESS,
    "Partial ⚠️": WARNING,
    "Short ❌": DANGER,
}


def inject_dashboard_style():
    """Professional styling — only applied to the results dashboard,
    never to the upload screen, so that one stays plain/default."""
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&family=Inter:wght@400;500;600&display=swap');

    /* Scope font + heading styling to the main results block only */
    div[data-testid="stAppViewContainer"] div[data-testid="block-container"] h1,
    div[data-testid="stAppViewContainer"] div[data-testid="block-container"] h2,
    div[data-testid="stAppViewContainer"] div[data-testid="block-container"] h3 {{
        font-family: 'Sora', sans-serif;
        color: {NAVY};
    }}

    /* Top banner */
    .dock-header {{
        background: linear-gradient(90deg, {NAVY} 0%, {STEEL} 100%);
        padding: 28px 32px;
        border-radius: 14px;
        margin-bottom: 24px;
        font-family: 'Inter', sans-serif;
    }}
    .dock-header h1 {{
        color: #FFFFFF !important;
        margin: 0;
        font-size: 1.7rem;
        font-family: 'Sora', sans-serif;
    }}
    .dock-header p {{
        color: #D7DEE8;
        margin: 6px 0 0 0;
        font-size: 0.95rem;
    }}

    /* KPI cards */
    .kpi-card {{
        background: {CARD};
        border-radius: 12px;
        padding: 18px 20px;
        box-shadow: 0 1px 3px rgba(18,35,63,0.08);
        border-left: 4px solid {AMBER};
        height: 100%;
        font-family: 'Inter', sans-serif;
    }}
    .kpi-label {{
        font-size: 0.78rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: {TEXT_MUTED};
        font-weight: 600;
        margin-bottom: 4px;
    }}
    .kpi-value {{
        font-family: 'Sora', sans-serif;
        font-size: 1.6rem;
        font-weight: 700;
        color: {NAVY};
    }}
    .kpi-sub {{
        font-size: 0.78rem;
        color: {TEXT_MUTED};
        margin-top: 2px;
    }}

    /* Section card wrapper */
    .section-card {{
        background: {CARD};
        border-radius: 12px;
        padding: 20px 22px;
        box-shadow: 0 1px 3px rgba(18,35,63,0.06);
        margin-bottom: 20px;
    }}
    </style>
    """, unsafe_allow_html=True)


PLOTLY_TEMPLATE = "plotly_white"
FONT = dict(family="Inter, sans-serif", color=NAVY)


def style_fig(fig, height=380):
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        font=FONT,
        title_font=dict(family="Sora, sans-serif", size=16, color=NAVY),
        margin=dict(l=10, r=10, t=50, b=10),
        height=height,
        plot_bgcolor="rgba(0,0,0,0)",
        paper_bgcolor="rgba(0,0,0,0)",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
    )
    return fig


def kpi_card(label, value, sub=""):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)


# =====================================================================
# HEADER (plain, matches the rest of the app's pages)
# =====================================================================
st.title("Shorts analysis Tool")
st.write("Turn incoming trailer inventory and outbound order shorts into a prioritized unloading plan.")

# =====================================================================
# INPUTS (plain, default Streamlit styling, on the page)
# =====================================================================
st.markdown("### Data Inputs")
col1, col2 = st.columns(2)
with col1:
    book_file = st.file_uploader("Trailer inventory", type=["xlsx"])
with col2:
    short_file = st.file_uploader("Order shorts (short sheet.xlsx)", type=["xlsx"])

run = st.button("Run Analysis")
st.caption(
    "Wave size and priority scoring are fixed at 4 trailers per wave, "
    "ranked by earliest dispatch time and demand efficiency."
)
st.divider()

if not book_file or not short_file:
    st.info("Upload both files above, then click **Run Analysis**.")
    st.stop()

if not run:
    st.info("Files loaded. Click **Run Analysis** to generate the dashboard.")
    st.stop()

# =====================================================================
# PROCESSING
# =====================================================================
try:
    # ---- LOAD BOOK1 ----
    df = pd.read_excel(book_file)
    df = df.iloc[2:].reset_index(drop=True)
    df = df.iloc[:, 4:]

    df.columns = [
        "ColE", "ColF", "ColG", "ColH", "ColI",
        "ColJ", "ColK", "ColL",
        "Date1", "Time1", "Date2", "Time2", "User", "ExtraDate"
    ]

    df["ColE"] = df["ColE"].astype(str).str.replace("L", "", regex=False)
    df["ColE"] = pd.to_numeric(df["ColE"], errors="coerce").fillna(0).astype(int)
    df["ColF"] = pd.to_numeric(df["ColF"], errors="coerce").fillna(0).astype(int)
    df["ColG"] = pd.to_numeric(df["ColG"], errors="coerce").fillna(0).astype(int)

    df["Trailer"] = df["ColE"].astype(str) + df["ColF"].astype(str) + df["ColG"].astype(str)
    df["SKU"] = df["ColJ"].astype(str).str.strip() + df["ColK"].astype(str).str.strip()

    clean_df = df[["Trailer", "SKU", "ColL"]].copy()
    clean_df.columns = ["Trailer", "SKU", "Quantity"]
    clean_df["Quantity"] = pd.to_numeric(clean_df["Quantity"], errors="coerce")
    clean_df = clean_df.dropna(subset=["SKU", "Quantity"])

    # ---- LOAD SHORT SHEET ----
    short_df = pd.read_excel(short_file)
    short_df = short_df.iloc[6:].reset_index(drop=True)

    short_df.columns = [
        "Trip", "Destination", "Dispatch", "Status", "Order",
        "Item", "Description", "Cases", "W", "ProdETA", "Comments"
    ]

    short_clean = short_df[["Trip", "Dispatch", "Item", "Cases"]].copy()
    short_clean["Item"] = short_clean["Item"].astype(str).str.strip()
    short_clean["Cases"] = pd.to_numeric(short_clean["Cases"], errors="coerce")
    short_clean["Dispatch"] = pd.to_numeric(short_clean["Dispatch"], errors="coerce")
    short_clean = short_clean.dropna(subset=["Item", "Cases"])

    # ---- MATCH ----
    # Full exact match: Book1's SKU (ColJ + ColK) reconstructs the same full decimal
    # item number used on the short sheet, e.g. "68820.13" + "081" = "68820.13081".
    match_df = short_clean.merge(clean_df, left_on="Item", right_on="SKU", how="left")

    # ---- TRAILER PRIORITY ----
    dispatch_lookup = short_clean.groupby("Item", as_index=False)["Dispatch"].min()
    dispatch_lookup = dispatch_lookup.rename(columns={"Dispatch": "Item_Dispatch"})
    match_with_dispatch = match_df.merge(dispatch_lookup, on="Item", how="left")

    trailer_priority = match_with_dispatch.groupby("Trailer").agg(
        Demand_Served=("Cases", "sum"),
        SKU_Count=("Item", "nunique"),
        Earliest_Dispatch=("Item_Dispatch", "min")
    ).reset_index()

    trailer_priority["Priority_Score"] = (
        trailer_priority["Demand_Served"] / trailer_priority["SKU_Count"]
    )

    trailer_priority = trailer_priority.sort_values(
        by=["Earliest_Dispatch", "Priority_Score"],
        ascending=[True, False]
    ).reset_index(drop=True)

    trailer_priority["Wave"] = (trailer_priority.index // 4) + 1

    # ---- LOAD COVERAGE + FILL (with real inventory allocation) ----
    # An item family's total inventory is shared across every load that needs it — two
    # loads can't both get credit for the same physical cases. Demand is allocated
    # family-by-family in dispatch order (earliest-shipping load first), so once a
    # family's cases run out, later loads correctly show as Partial/Short.
    item_totals = clean_df.groupby("SKU", as_index=False)["Quantity"].sum()
    item_totals = item_totals.rename(columns={"SKU": "Item", "Quantity": "Total_Item_Inventory"})

    alloc = short_clean.sort_values(["Item", "Dispatch", "Trip"]).copy()
    alloc = alloc.merge(item_totals, on="Item", how="left")
    alloc["Total_Item_Inventory"] = alloc["Total_Item_Inventory"].fillna(0)

    alloc["Cum_Demand"] = alloc.groupby("Item")["Cases"].cumsum()
    alloc["Cum_Allocated"] = alloc[["Cum_Demand", "Total_Item_Inventory"]].min(axis=1)
    alloc["Prev_Cum_Allocated"] = alloc.groupby("Item")["Cum_Allocated"].shift(fill_value=0)
    alloc["Allocated_Cases"] = alloc["Cum_Allocated"] - alloc["Prev_Cum_Allocated"]
    alloc["Shortage_Cases"] = alloc["Cases"] - alloc["Allocated_Cases"]

    alloc["Fill_Rate"] = (
        alloc["Allocated_Cases"] / alloc["Cases"]
    ).replace([np.inf, -np.inf], 0).fillna(0)

    def get_status(x):
        if x >= 1:
            return "Full ✅"
        elif x > 0:
            return "Partial ⚠️"
        else:
            return "Short ❌"

    alloc["Status"] = alloc["Fill_Rate"].apply(get_status)

    # Which trailer(s) actually carry this exact item, and the earliest wave any of
    # them unload in — this is "where to find it," shown for reference. Blank means no
    # trailer currently carries this item at all (a genuine inventory gap).
    trailer_lookup = clean_df.merge(trailer_priority[["Trailer", "Wave"]], on="Trailer", how="left")
    trailer_lookup = trailer_lookup.groupby("SKU").agg(
        Trailer=("Trailer", lambda s: ", ".join(sorted(set(s.astype(str))))),
        Wave=("Wave", "min")
    ).reset_index().rename(columns={"SKU": "Item"})

    load_trailer = alloc.merge(trailer_lookup, on="Item", how="left")
    load_trailer = load_trailer.rename(columns={"Cases": "Demand_Cases"})
    load_trailer = load_trailer[
        ["Wave", "Trailer", "Trip", "Item", "Dispatch", "Demand_Cases",
         "Allocated_Cases", "Total_Item_Inventory", "Fill_Rate", "Status", "Shortage_Cases"]
    ]

    # ---- EXCEPTIONS ----
    exceptions = load_trailer[load_trailer["Status"] != "Full ✅"].copy()
    exceptions = exceptions.sort_values(by=["Dispatch", "Status"])

    # ---- OPTIMIZED TRAILERS ----
    # For each item still short/partial, find the physical trailers that carry it and
    # how many cases each one holds (Fix_Cases — properly summed across every pallet/LPN
    # on that trailer, not just deduplicated rows), and how many distinct loads are
    # waiting on that item (Loads_Impacted).
    problem_items = exceptions["Item"].unique()

    trailer_item_qty = clean_df[clean_df["SKU"].isin(problem_items)].groupby(
        ["Trailer", "SKU"], as_index=False
    )["Quantity"].sum().rename(columns={"SKU": "Item"})

    loads_per_item = exceptions.groupby("Item")["Trip"].nunique()
    trailer_item_qty["Loads_Impacted"] = trailer_item_qty["Item"].map(loads_per_item).fillna(0).astype(int)

    optimized_trailers = trailer_item_qty.groupby("Trailer").agg(
        Fix_Cases=("Quantity", "sum"),
        Loads_Impacted=("Loads_Impacted", "sum")
    ).reset_index().sort_values(by=["Fix_Cases", "Loads_Impacted"], ascending=[False, False])

    top4_trailers = optimized_trailers.head(4).copy()
    top4_trailers.insert(0, "Rank", range(1, len(top4_trailers) + 1))

    # ---- FORMAT EXPORTS ----
    dock_plan_export = trailer_priority.copy()
    dock_plan_export["Priority_Score"] = dock_plan_export["Priority_Score"].round(0)
    cols = ["Wave"] + [c for c in dock_plan_export.columns if c != "Wave"]
    dock_plan_export = dock_plan_export[cols]

    load_export = load_trailer.copy()
    load_export["Fill_Rate"] = load_export["Fill_Rate"].round(2)

    exception_export = exceptions.copy()
    exception_export["Fill_Rate"] = exception_export["Fill_Rate"].round(2)

except Exception as e:
    st.error(f"Something went wrong while processing the files: {e}")
    st.stop()

# =====================================================================
# RESULTS DASHBOARD — professional styling starts here only
# =====================================================================
inject_dashboard_style()

st.markdown("""
<div class="dock-header">
    <h1>Dock Optimization Results</h1>
    <p>Prioritized unloading plan generated from your uploaded files.</p>
</div>
""", unsafe_allow_html=True)

# =====================================================================
# KPI ROW
# =====================================================================
total_trailers = dock_plan_export["Trailer"].nunique()
total_demand = int(load_export["Demand_Cases"].sum())
total_shortage = int(load_export["Shortage_Cases"].clip(lower=0).sum())
overall_fill = (
    load_export["Allocated_Cases"].sum() / load_export["Demand_Cases"].sum()
    if load_export["Demand_Cases"].sum() else 0
)
total_waves = dock_plan_export["Wave"].nunique()
loads_met_pct = (load_export["Status"].eq("Full ✅").mean() * 100) if len(load_export) else 0

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    kpi_card("Trailers", f"{total_trailers}", f"across {total_waves} waves")
with k2:
    kpi_card("Total Demand", f"{total_demand:,}", "cases ordered")
with k3:
    kpi_card("Overall Fill Rate", f"{overall_fill*100:.0f}%", "available vs. demand")
with k4:
    kpi_card("Shortage Cases", f"{total_shortage:,}", "cases still missing")
with k5:
    kpi_card("Loads Fully Met", f"{loads_met_pct:.0f}%", "of load lines fully covered")

st.write("")

# =====================================================================
# TABS
# =====================================================================
tab_overview, tab_wave, tab_load, tab_exceptions, tab_fix = st.tabs(
    ["Overview", "Wave Plan", "Load Coverage", "Exceptions", "Top Fixes"]
)

# ---------------- OVERVIEW ----------------
with tab_overview:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig1 = px.bar(
            dock_plan_export.head(15), x="Trailer", y="Demand_Served",
            color="Wave", title="Top Trailers by Demand", text="Demand_Served",
            color_continuous_scale=[STEEL, AMBER]
        )
        st.plotly_chart(style_fig(fig1), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        status_summary = load_export.groupby("Status").size().reset_index(name="Count")
        fig3 = px.pie(
            status_summary, names="Status", values="Count",
            title="Load Status Breakdown", color="Status",
            color_discrete_map=STATUS_COLORS, hole=0.55
        )
        st.plotly_chart(style_fig(fig3), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    c3, c4 = st.columns(2)
    with c3:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig2 = px.bar(
            top4_trailers, x="Trailer", y="Fix_Cases",
            title="Top 4 Trailers to Move to Door", text="Fix_Cases",
            color_discrete_sequence=[SUCCESS]
        )
        st.plotly_chart(style_fig(fig2), use_container_width=True)
        st.caption("Unload these trailers next — they resolve the most shortage cases across the most loads.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        short_items = load_export[load_export["Shortage_Cases"] > 0]
        top_skus = short_items.groupby("Item").agg(
            Shortage_Cases=("Shortage_Cases", "sum"),
            Trailer=("Trailer", lambda s: ", ".join(sorted(set(x for x in s if pd.notna(x) and x != ""))) or "No inventory found")
        ).reset_index()
        top_skus = top_skus.sort_values(by="Shortage_Cases", ascending=False).head(10)
        fig4 = px.bar(
            top_skus, x="Item", y="Shortage_Cases", title="Top Shortage Drivers (Cases Missing)",
            text="Shortage_Cases", color_discrete_sequence=[DANGER],
            hover_data={"Trailer": True, "Item": False}
        )
        st.plotly_chart(style_fig(fig4), use_container_width=True)
        st.caption("Hover a bar to see which trailer(s), if any, carry that item.")
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------- WAVE PLAN ----------------
with tab_wave:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Dock Plan — Trailers by Wave")
    st.caption("Wave 1 should be unloaded first. Sorted by earliest dispatch time, then priority score.")
    st.caption(
        "Priority Score = Demand Served ÷ SKU Count — the average cases served per distinct item on "
        "that trailer. A high score means the trailer covers a lot of demand with very few items "
        "(efficient to unload); a lower score means demand is spread across more items."
    )
    wave_filter = st.multiselect(
        "Filter by wave", sorted(dock_plan_export["Wave"].unique()),
        default=sorted(dock_plan_export["Wave"].unique())
    )
    st.dataframe(
        dock_plan_export[dock_plan_export["Wave"].isin(wave_filter)],
        use_container_width=True, hide_index=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- LOAD COVERAGE ----------------
with tab_load:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Load Coverage — Demand vs. Available by Trip")
    status_filter = st.multiselect(
        "Filter by status", list(STATUS_COLORS.keys()), default=list(STATUS_COLORS.keys())
    )
    st.dataframe(
        load_export[load_export["Status"].isin(status_filter)],
        use_container_width=True, hide_index=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- EXCEPTIONS ----------------
with tab_exceptions:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Exception Report — Partial & Short Loads")
    st.caption("Sorted by dispatch time. These are the loads that need attention before they ship.")
    st.caption(
        "Only load lines that did NOT get 100% of what they needed. Use this to see, at a glance, "
        "exactly which shipments are at risk before they go out the door."
    )
    st.dataframe(exception_export, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- TOP FIXES ----------------
with tab_fix:
    c1, c2 = st.columns([1, 1])
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        fig5 = px.bar(
            top4_trailers, x="Trailer", y="Fix_Cases",
            title="Top 4 Trailers to Fix Shorts", text="Fix_Cases",
            color_discrete_sequence=[SUCCESS]
        )
        st.plotly_chart(style_fig(fig5), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        st.subheader("Ranked Trailers")
        st.caption("Prioritize these trailers first — they resolve the most shortage cases across the most loads.")
        st.caption(
            "Fix_Cases = cases of shortage-causing items this trailer has on board. "
            "Loads_Impacted = how many different shipments that inventory could help complete."
        )
        st.dataframe(top4_trailers, use_container_width=True, hide_index=True)
        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("All Optimized Trailers")
    st.dataframe(optimized_trailers, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# =====================================================================
# EXPORT
# =====================================================================
def build_excel():
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        dock_plan_export.to_excel(writer, sheet_name="Dock Plan", index=False)
        load_export.to_excel(writer, sheet_name="Load Coverage", index=False)
        exception_export.to_excel(writer, sheet_name="Exception Report", index=False)
        optimized_trailers.to_excel(writer, sheet_name="Optimized Trailers", index=False)
        top4_trailers.to_excel(writer, sheet_name="Top 4 Trailers", index=False)

        for sheet in writer.sheets:
            ws = writer.sheets[sheet]
            for row in ws.iter_rows():
                for cell in row:
                    cell.alignment = Alignment(horizontal="center", vertical="center")
    buffer.seek(0)
    return buffer

st.markdown("### Export")
st.download_button(
    label="Download Full Report (.xlsx)",
    data=build_excel(),
    file_name="Short_Analysis.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
