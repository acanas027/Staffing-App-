import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import Alignment
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape, portrait
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak, KeepInFrame
from reportlab.lib.styles import getSampleStyleSheet

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
    "Full": SUCCESS,
    "Partial": WARNING,
    "Short": DANGER,
}


def inject_dashboard_style():
    st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Sora:wght@500;600;700&family=Inter:wght@400;500;600&display=swap');

    div[data-testid="stAppViewContainer"] div[data-testid="block-container"] h1,
    div[data-testid="stAppViewContainer"] div[data-testid="block-container"] h2,
    div[data-testid="stAppViewContainer"] div[data-testid="block-container"] h3 {{
        font-family: 'Sora', sans-serif;
        color: {NAVY};
    }}

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

    .kpi-card {{
        background: {CARD};
        border-radius: 10px;
        padding: 10px 14px;
        box-shadow: 0 1px 3px rgba(18,35,63,0.08);
        border-left: 4px solid {AMBER};
        height: 100%;
        font-family: 'Inter', sans-serif;
    }}
    .kpi-label {{
        font-size: 0.72rem;
        text-transform: uppercase;
        letter-spacing: 0.04em;
        color: {TEXT_MUTED};
        font-weight: 600;
        margin-bottom: 2px;
    }}
    .kpi-value {{
        font-family: 'Sora', sans-serif;
        font-size: 1.25rem;
        font-weight: 700;
        color: {NAVY};
    }}
    .kpi-sub {{
        font-size: 0.7rem;
        color: {TEXT_MUTED};
        margin-top: 1px;
    }}

    .section-card {{
        background: {CARD};
        border-radius: 10px;
        padding: 8px 12px;
        box-shadow: 0 1px 3px rgba(18,35,63,0.06);
        margin-bottom: 8px;
    }}
    </style>
    """, unsafe_allow_html=True)


PLOTLY_TEMPLATE = "plotly_white"
FONT = dict(family="Inter, sans-serif", color=NAVY)


def style_fig(fig, height=260):
    fig.update_layout(
        template=PLOTLY_TEMPLATE,
        font=FONT,
        title_font=dict(family="Sora, sans-serif", size=14, color=NAVY),
        margin=dict(l=10, r=10, t=40, b=10),
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


STATUS_PDF_COLORS = {
    "Full": colors.HexColor("#C7F0D8"),
    "Partial": colors.HexColor("#FFE8A3"),
    "Short": colors.HexColor("#F5C2C7"),
}


def build_status_table(df, status_col="Status"):
    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]
    if status_col in df.columns:
        col_idx = list(df.columns).index(status_col)
        for i, val in enumerate(df[status_col].astype(str).tolist()):
            color = STATUS_PDF_COLORS.get(val)
            if color:
                style_commands.append(("BACKGROUND", (col_idx, i + 1), (col_idx, i + 1), color))
    table.setStyle(TableStyle(style_commands))
    return table


def build_full_pdf(title, kpis, figs, wave_df, load_df):
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=portrait(letter),
        topMargin=30, bottomMargin=30, leftMargin=30, rightMargin=30
    )
    styles = getSampleStyleSheet()
    story = []

    overview_page = []
    overview_page.append(Paragraph(f"{title} — Overview", styles["Title"]))
    overview_page.append(Spacer(1, 14))

    header_row = [k[0] for k in kpis]
    value_row = [k[1] for k in kpis]
    sub_row = [k[2] for k in kpis]

    kpi_table = Table([header_row, value_row, sub_row], colWidths=[87] * len(kpis))
    kpi_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, 0), 8),
        ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 1), (-1, 1), 15),
        ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor(NAVY)),
        ("FONTSIZE", (0, 2), (-1, 2), 7),
        ("TEXTCOLOR", (0, 2), (-1, 2), colors.HexColor(TEXT_MUTED)),
        ("BOX", (0, 0), (-1, -1), 0.75, colors.HexColor(NAVY)),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#DDDDDD")),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))

    overview_page.append(kpi_table)
    overview_page.append(Spacer(1, 18))

    images = []
    for chart_title, fig in figs:
        png_bytes = fig.to_image(format="png", scale=2, width=520, height=316)
        images.append(Image(BytesIO(png_bytes), width=255, height=155))

    rows = []
    for i in range(0, len(images), 2):
        pair = images[i:i + 2]
        if len(pair) == 1:
            pair.append("")
        rows.append(pair)

    if rows:
        chart_table = Table(rows, colWidths=[265, 265])
        chart_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        overview_page.append(chart_table)

    story.append(KeepInFrame(doc.width, doc.height, overview_page, mode="shrink"))

    # Wave Plan — Priority column removed
    story.append(PageBreak())
    story.append(Paragraph(f"{title} — Wave Plan", styles["Title"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Dock Plan — Trailers in Priority Order", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(build_status_table(wave_df))

    story.append(PageBreak())
    story.append(Paragraph(f"{title} — Load Coverage", styles["Title"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Load Coverage — Demand vs. Available by Trip", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(build_status_table(load_df))

    doc.build(story)
    buffer.seek(0)
    return buffer


# =====================================================================
# HEADER
# =====================================================================
st.title("Shorts Analysis Tool")
st.write("Turn incoming trailer inventory and outbound order shorts into a prioritized unloading plan.")

# =====================================================================
# INPUTS
# =====================================================================
st.markdown("### Data Inputs")
col1, col2, col3 = st.columns(3)
with col1:
    book_file = st.file_uploader("Trailer inventory", type=["xlsx"])
with col2:
    short_file = st.file_uploader("Order shorts (short sheet.xlsx)", type=["xlsx"])
with col3:
    transfers_file = st.file_uploader("Transfers file (optional)", type=["xlsx"])

run = st.button("Run Analysis")
st.caption(
    "Wave size: 4 trailers per wave. "
    "Priority: earliest dispatch on a short load first; ties broken by cases solved."
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
    # ---- LOAD INVENTORY (BOOK) ----
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

    def _build_sku(j, k):
        if pd.isna(j):
            return None
        return f"{j:.2f}" + (f"{int(k):03d}" if pd.notna(k) else "000")

    df["SKU"] = [_build_sku(j, k) for j, k in zip(df["ColJ"], df["ColK"])]

    clean_df = df[["Trailer", "SKU", "ColL"]].copy()
    clean_df.columns = ["Trailer", "SKU", "Quantity"]
    clean_df["Quantity"] = pd.to_numeric(clean_df["Quantity"], errors="coerce")
    clean_df = clean_df.dropna(subset=["SKU", "Quantity"])

    # ---- LOAD TRANSFERS (optional) — count unique trailers ----
    total_transfer_trailers = None
    if transfers_file is not None:
        try:
            tr_raw = pd.read_excel(transfers_file)
            tr_data = tr_raw.iloc[2:].reset_index(drop=True)
            tr_data = tr_data.iloc[:, 4:]
            tr_data.columns = [
                "ColE", "ColF", "ColG", "ColH", "ColI",
                "ColJ", "ColK", "ColL",
                "Date1", "Time1", "Date2", "Time2", "User", "ExtraDate"
            ]
            tr_data["ColE"] = tr_data["ColE"].astype(str).str.replace("L", "", regex=False)
            tr_data["ColE"] = pd.to_numeric(tr_data["ColE"], errors="coerce").fillna(0).astype(int)
            tr_data["ColF"] = pd.to_numeric(tr_data["ColF"], errors="coerce").fillna(0).astype(int)
            tr_data["ColG"] = pd.to_numeric(tr_data["ColG"], errors="coerce").fillna(0).astype(int)
            tr_data["Trailer"] = tr_data["ColE"].astype(str) + tr_data["ColF"].astype(str) + tr_data["ColG"].astype(str)
            total_transfer_trailers = tr_data["Trailer"].nunique()
        except Exception:
            total_transfer_trailers = None

    # ---- LOAD SHORT SHEET ----
    short_df = pd.read_excel(short_file, header=2)
    short_df = short_df.iloc[:, :11].reset_index(drop=True)
    short_df.columns = [
        "Trip", "Destination", "Dispatch", "Status", "Order",
        "Item", "Description", "Cases", "W", "ProdETA", "Comments"
    ]

    short_clean = short_df[["Trip", "Dispatch", "Item", "Cases"]].copy()
    short_clean["Item"] = pd.to_numeric(short_clean["Item"], errors="coerce").map(
        lambda v: f"{v:.5f}" if pd.notna(v) else None
    )
    short_clean["Cases"] = pd.to_numeric(short_clean["Cases"], errors="coerce")
    short_clean["Dispatch"] = pd.to_numeric(short_clean["Dispatch"], errors="coerce")
    short_clean = short_clean.dropna(subset=["Item", "Cases"])

    # FIX: Drop rows where Trip is blank/empty string — these are sub-rows without a
    # real trip number and should not be counted as separate loads.
    short_clean["Trip"] = short_clean["Trip"].astype(str).str.strip()
    short_clean = short_clean[short_clean["Trip"] != ""].copy()

    # ---- MATCH ----
    match_df = short_clean.merge(clean_df, left_on="Item", right_on="SKU", how="left")

    _matched = int(match_df["SKU"].notna().sum())
    st.success(f"Matched {_matched} of {len(match_df)} short lines to trailer inventory.")

    # ---- LOAD COVERAGE + FILL ----
    item_totals = clean_df.groupby("SKU", as_index=False)["Quantity"].sum()
    item_totals = item_totals.rename(columns={"SKU": "Item", "Quantity": "Total_Item_Inventory"})

    alloc = short_clean.sort_values(["Item", "Dispatch", "Trip"]).copy()
    alloc = alloc.merge(item_totals, on="Item", how="left")
    alloc["Total_Item_Inventory"] = alloc["Total_Item_Inventory"].fillna(0)

    alloc["Cum_Demand"] = alloc.groupby("Item")["Cases"].cumsum()
    alloc["Cum_Allocated"] = alloc[["Cum_Demand", "Total_Item_Inventory"]].min(axis=1)
    alloc["Prev_Cum_Allocated"] = alloc.groupby("Item")["Cum_Allocated"].shift(fill_value=0)
    alloc["Allocated_Cases"] = alloc["Cum_Allocated"] - alloc["Prev_Cum_Allocated"]
    alloc["Actual_short_cases"] = alloc["Cases"] - alloc["Allocated_Cases"]

    alloc["Fill_Rate"] = (
        alloc["Allocated_Cases"] / alloc["Cases"]
    ).replace([np.inf, -np.inf], 0).fillna(0)

    def get_status(x):
        if x >= 1:
            return "Full"
        elif x > 0:
            return "Partial"
        else:
            return "Short"

    alloc["Status"] = alloc["Fill_Rate"].apply(get_status)

    # ---- EXCEPTIONS ----
    exceptions_raw = alloc[alloc["Status"] != "Full"].copy()

    # ---- OPTIMIZED TRAILERS ----
    problem_items = exceptions_raw["Item"].unique()

    trailer_item_qty = clean_df[clean_df["SKU"].isin(problem_items)].groupby(
        ["Trailer", "SKU"], as_index=False
    )["Quantity"].sum().rename(columns={"SKU": "Item"})

    loads_per_item = exceptions_raw.groupby("Item")["Trip"].nunique()
    trailer_item_qty["Loads_Impacted"] = trailer_item_qty["Item"].map(loads_per_item).fillna(0).astype(int)

    optimized_trailers = trailer_item_qty.groupby("Trailer").agg(
        Fix_Cases=("Quantity", "sum"),
        Loads_Impacted=("Loads_Impacted", "sum")
    ).reset_index().sort_values(by=["Fix_Cases", "Loads_Impacted"], ascending=[False, False])

    top4_trailers = optimized_trailers.head(4).copy()
    top4_trailers.insert(0, "Rank", range(1, len(top4_trailers) + 1))

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

    loads_impacted_lookup = optimized_trailers.set_index("Trailer")["Loads_Impacted"]
    trailer_priority["Loads_Impacted"] = trailer_priority["Trailer"].map(loads_impacted_lookup).fillna(0).astype(int)

    fix_cases_lookup = optimized_trailers.set_index("Trailer")["Fix_Cases"]
    trailer_priority["Fix_Cases"] = trailer_priority["Trailer"].map(fix_cases_lookup).fillna(0)

    trailer_priority["Fixes_Shortage"] = trailer_priority["Loads_Impacted"] > 0

    # Selection rank for Load Coverage (one best trailer per item)
    trailer_priority = trailer_priority.sort_values(
        by=["Fixes_Shortage", "Earliest_Dispatch", "Fix_Cases", "Priority_Score"],
        ascending=[False, True, False, False]
    ).reset_index(drop=True)
    trailer_priority["Selection_Rank"] = range(1, len(trailer_priority) + 1)

    trailer_lookup_base = clean_df[["SKU", "Trailer"]].drop_duplicates().merge(
        trailer_priority[["Trailer", "Selection_Rank"]],
        on="Trailer",
        how="left"
    )

    priority_trailer_lookup = trailer_lookup_base.sort_values(
        by=["SKU", "Selection_Rank"],
        na_position="last"
    ).groupby("SKU", as_index=False).first()
    priority_trailer_lookup = priority_trailer_lookup.rename(columns={"SKU": "Item"})

    load_trailer = alloc.merge(priority_trailer_lookup[["Item", "Trailer"]], on="Item", how="left")
    load_trailer = load_trailer.rename(columns={"Cases": "Demand_Cases"})

    solved_loads_lookup = load_trailer[load_trailer["Allocated_Cases"] > 0].groupby("Trailer")["Trip"].nunique()
    trailer_priority["Loads_Solved"] = trailer_priority["Trailer"].map(solved_loads_lookup).fillna(0).astype(int)
    trailer_priority["Solves_Load"] = trailer_priority["Loads_Solved"] > 0

    # =====================================================================
    # FINAL WAVE PLAN RANKING
    # Primary: earliest dispatch → more short cases breaks ties → loads impacted
    # FIX: only include trailers that have Fix_Cases > 0, i.e. they actually
    # carry shortage inventory. Trailers that only cover Full-status loads
    # (Fix_Cases = 0) are excluded from the Wave Plan entirely.
    # =====================================================================
    trailer_priority = trailer_priority.sort_values(
        by=["Solves_Load", "Fixes_Shortage", "Earliest_Dispatch", "Fix_Cases", "Loads_Impacted", "Priority_Score"],
        ascending=[False, False, True, False, False, False]
    ).reset_index(drop=True)

    # Only trailers with actual short/partial cases make the wave plan
    wave_mask = trailer_priority["Solves_Load"] & (trailer_priority["Fix_Cases"] > 0)

    trailer_priority["Wave"] = np.nan
    trailer_priority["Trailer_Priority"] = np.nan
    trailer_priority.loc[wave_mask, "Trailer_Priority"] = range(1, int(wave_mask.sum()) + 1)
    trailer_priority.loc[wave_mask, "Wave"] = (
        (trailer_priority.loc[wave_mask, "Trailer_Priority"] - 1) // 4
    ) + 1

    load_trailer = load_trailer.merge(
        trailer_priority[["Trailer", "Wave", "Trailer_Priority"]],
        on="Trailer",
        how="left"
    )

    load_trailer = load_trailer[
        ["Wave", "Trailer_Priority", "Trailer", "Trip", "Item", "Dispatch", "Demand_Cases",
         "Allocated_Cases", "Total_Item_Inventory", "Fill_Rate", "Status", "Actual_short_cases"]
    ]

    exceptions = load_trailer[load_trailer["Status"] != "Full"].copy()
    exceptions = exceptions.sort_values(by=["Dispatch", "Status"])

    # ---- FORMAT EXPORTS ----
    # Wave Plan: Priority column removed, only Fix_Cases > 0 trailers
    dock_plan_export = trailer_priority[wave_mask].drop(
        columns=[
            "Trailer_Priority", "Priority_Score", "Loads_Impacted",
            "Fixes_Shortage", "Selection_Rank", "Loads_Solved", "Solves_Load"
        ]
    ).copy()
    # No Priority column — Wave + order is self-evident
    cols = ["Wave"] + [c for c in dock_plan_export.columns if c != "Wave"]
    dock_plan_export = dock_plan_export[cols].reset_index(drop=True)

    load_export = load_trailer.copy()
    load_export["Fill_Rate"] = load_export["Fill_Rate"].round(2)

    load_export["_Wave_Sort"] = load_export["Wave"].fillna(9999)
    load_export["_Trailer_Priority_Sort"] = load_export["Trailer_Priority"].fillna(9999)
    load_export["_No_Value_Sort"] = load_export["Wave"].isna().astype(int)
    load_export = load_export.sort_values(
        by=["_No_Value_Sort", "_Wave_Sort", "_Trailer_Priority_Sort", "Dispatch", "Trip", "Status", "Actual_short_cases"],
        ascending=[True, True, True, True, True, True, False]
    ).drop(
        columns=["_Wave_Sort", "_Trailer_Priority_Sort", "_No_Value_Sort", "Trailer_Priority"]
    ).reset_index(drop=True)

    load_export["Wave"] = load_export["Wave"].apply(lambda x: "" if pd.isna(x) else int(x))
    load_export["Trailer"] = load_export["Trailer"].fillna("No trailer found")

    exception_export = exceptions.copy()
    exception_export["Fill_Rate"] = exception_export["Fill_Rate"].round(2)
    if "Trailer_Priority" in exception_export.columns:
        exception_export = exception_export.drop(columns=["Trailer_Priority"])
    exception_export["Wave"] = exception_export["Wave"].apply(lambda x: "" if pd.isna(x) else int(x))
    exception_export["Trailer"] = exception_export["Trailer"].fillna("No trailer found")

except Exception as e:
    st.error(f"Something went wrong while processing the files: {e}")
    st.stop()

# =====================================================================
# RESULTS DASHBOARD
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
total_cases_short = int(load_export["Demand_Cases"].sum())
total_shortage = int(load_export["Actual_short_cases"].clip(lower=0).sum())
total_waves = dock_plan_export["Wave"].nunique()

# FIX: only count trips that are real (non-blank) — already handled upstream,
# but this ensures the KPI reflects the corrected trip_status groupby.
trip_status = load_export.groupby("Trip")["Status"].apply(lambda s: (s == "Full").all())
loads_met_count = int(trip_status.sum())
total_loads = int(trip_status.shape[0])

if not top4_trailers.empty:
    top_trailer = top4_trailers.iloc[0]
    move_next_value = f"Trailer {top_trailer['Trailer']}"
    move_next_sub = f"fixes {int(top_trailer['Fix_Cases']):,} cases across {int(top_trailer['Loads_Impacted'])} load(s)"
else:
    move_next_value = "—"
    move_next_sub = "no shortages to fix"

# KPI columns: 6 if transfers loaded, 5 if not
if total_transfer_trailers is not None:
    k1, k2, k3, k4, k5, k6 = st.columns(6)
    with k6:
        kpi_card("Transfer Trailers", f"{total_transfer_trailers}", "unique trailers in today's transfers")
else:
    k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    kpi_card("Trailers Involved", f"{total_trailers}", f"carry shortage items, across {total_waves} waves")
with k2:
    kpi_card("Total Cases Ordered", f"{total_cases_short:,}", "cases on the short sheet")
with k3:
    kpi_card("Actual Short Cases", f"{total_shortage:,}", "cases still missing after allocation")
with k4:
    kpi_card("Loads Fully Met", f"{loads_met_count:,}", f"of {total_loads:,} total loads")
with k5:
    kpi_card("Move Next", move_next_value, move_next_sub)

# =====================================================================
# TABS
# =====================================================================
tab_overview, tab_wave, tab_load = st.tabs(["Overview", "Wave Plan", "Load Coverage"])

# ---------------- OVERVIEW ----------------
with tab_overview:
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        top_trailers_chart_df = dock_plan_export.head(15).copy()
        top_trailers_chart_df["Wave"] = top_trailers_chart_df["Wave"].astype(str)
        fig1 = px.bar(
            top_trailers_chart_df, x="Trailer", y="Demand_Served",
            color="Wave", title="Top Trailers by Demand", text="Demand_Served",
            color_discrete_sequence=[STEEL, AMBER]
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
        priority_order = dock_plan_export.sort_values("Wave")
        priority_order_chart_df = priority_order.copy()
        priority_order_chart_df["Wave"] = priority_order_chart_df["Wave"].astype(str)
        fig2 = px.bar(
            priority_order_chart_df, x="Demand_Served", y="Trailer", orientation="h",
            color="Wave", title="Trailer Priority Order — Today",
            category_orders={"Trailer": priority_order_chart_df["Trailer"].tolist()},
            color_discrete_sequence=[STEEL, AMBER]
        )
        st.plotly_chart(style_fig(fig2), use_container_width=True)
        st.caption("Earliest dispatch drives the ranking; cases solved breaks ties.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        short_items = load_export[load_export["Actual_short_cases"] > 0]
        top_skus = short_items.groupby("Item").agg(
            Actual_short_cases=("Actual_short_cases", "sum"),
            Trailer=("Trailer", lambda s: ", ".join(sorted(set(x for x in s if pd.notna(x) and x != ""))) or "No inventory found")
        ).reset_index()
        top_skus = top_skus.sort_values(by="Actual_short_cases", ascending=False).head(10)
        fig4 = px.bar(
            top_skus, x="Item", y="Actual_short_cases", title="Top Shortage Items",
            text="Actual_short_cases", color_discrete_sequence=[DANGER],
            hover_data={"Trailer": True, "Item": False}
        )
        st.plotly_chart(style_fig(fig4), use_container_width=True)
        st.caption("Hover a bar to see the selected trailer, if any, for that item.")
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------- WAVE PLAN ----------------
with tab_wave:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Dock Plan — Trailers in Priority Order")
    st.caption(
        "Only trailers carrying actual shortage inventory appear here. "
        "Priority: earliest dispatch first; ties broken by most short cases carried. "
        "Waves are groups of 4."
    )
    st.dataframe(dock_plan_export, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- LOAD COVERAGE ----------------
with tab_load:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Load Coverage — Demand vs. Available by Trip")
    st.caption("Status: green = Full, yellow = Partial, red = Short.")

    def highlight_status(val):
        color_map = {"Full": "#C7F0D8", "Partial": "#FFE8A3", "Short": "#F5C2C7"}
        return f"background-color: {color_map.get(val, '')}"

    try:
        styled_load = load_export.style.map(highlight_status, subset=["Status"])
    except AttributeError:
        styled_load = load_export.style.applymap(highlight_status, subset=["Status"])
    st.dataframe(styled_load, use_container_width=True, hide_index=True)
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
c1, c2 = st.columns(2)
with c1:
    st.download_button(
        label="Download Full Report (.xlsx)",
        data=build_excel(),
        file_name="Short_Analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
with c2:
    st.download_button(
        label="Download Full Report (.pdf)",
        data=build_full_pdf(
            "Shorts Analysis Results",
            kpis=[
                ("Trailers Involved", str(total_trailers), f"across {total_waves} waves"),
                ("Cases Ordered", f"{total_cases_short:,}", "on short sheet"),
                ("Short Cases", f"{total_shortage:,}", "still missing"),
                ("Loads Fully Met", f"{loads_met_count:,}", f"of {total_loads:,} loads"),
                ("Move Next", move_next_value, move_next_sub),
            ] + ([("Transfer Trailers", str(total_transfer_trailers), "today's transfers")] if total_transfer_trailers is not None else []),
            figs=[
                ("Top Trailers by Demand", fig1),
                ("Load Status Breakdown", fig3),
                ("Trailer Priority Order — Today", fig2),
                ("Top Shortage Items", fig4),
            ],
            wave_df=dock_plan_export,
            load_df=load_export
        ),
        file_name="Shorts_Analysis_Report.pdf",
        mime="application/pdf"
    )
