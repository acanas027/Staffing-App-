import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import Alignment
from io import BytesIO
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image, PageBreak
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

    /* Section card wrapper */
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
    "Full ✅": colors.HexColor("#C7F0D8"),
    "Partial ⚠️": colors.HexColor("#FFE8A3"),
    "Short ❌": colors.HexColor("#F5C2C7"),
}


def build_status_table(df, status_col="Status"):
    """Reportlab Table with the Status column's cells colored green/yellow/red."""
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
    """One combined PDF: Overview (KPIs + charts), Wave Plan, and Load Coverage — each on its own page."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=landscape(letter),
        topMargin=30, bottomMargin=30, leftMargin=30, rightMargin=30
    )
    styles = getSampleStyleSheet()
    story = []

    # ---- Page 1: Overview ----
    story.append(Paragraph(f"{title} — Overview", styles["Title"]))
    story.append(Spacer(1, 14))
    header_row = [k[0] for k in kpis]
    value_row = [k[1] for k in kpis]
    sub_row = [k[2] for k in kpis]
    kpi_table = Table([header_row, value_row, sub_row], colWidths=[150] * len(kpis))
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
    story.append(kpi_table)
    story.append(Spacer(1, 18))

    images = []
    for chart_title, fig in figs:
        png_bytes = fig.to_image(format="png", scale=2, width=560, height=340)
        images.append(Image(BytesIO(png_bytes), width=350, height=213))
    rows = []
    for i in range(0, len(images), 2):
        pair = images[i:i + 2]
        if len(pair) == 1:
            pair.append("")
        rows.append(pair)
    if rows:
        chart_table = Table(rows, colWidths=[360, 360])
        chart_table.setStyle(TableStyle([
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("TOPPADDING", (0, 0), (-1, -1), 10),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(chart_table)

    # ---- Page 2: Wave Plan ----
    story.append(PageBreak())
    story.append(Paragraph(f"{title} — Wave Plan", styles["Title"]))
    story.append(Spacer(1, 10))
    story.append(Paragraph("Dock Plan — Trailers in Priority Order", styles["Heading2"]))
    story.append(Spacer(1, 6))
    story.append(build_status_table(wave_df))

    # ---- Page 3: Load Coverage ----
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
    short_df = pd.read_excel(short_file, header=2)
    short_df = short_df.iloc[:, :11].reset_index(drop=True)

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

    # ---- LOAD COVERAGE + FILL (with real inventory allocation) ----
    # An item's total inventory is shared across every load that needs it — two loads
    # can't both get credit for the same physical cases. Demand is allocated item-by-item
    # in dispatch order (earliest-shipping load first), so once an item's cases run out,
    # later loads correctly show as Partial/Short instead of both showing Full.
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

    # ---- EXCEPTIONS ----
    exceptions_raw = alloc[alloc["Status"] != "Full ✅"].copy()

    # ---- OPTIMIZED TRAILERS ----
    # For each item still short/partial, find the physical trailers that carry it and
    # how many cases each one holds (Fix_Cases — properly summed across every pallet/LPN
    # on that trailer, not just deduplicated rows), and how many distinct loads are
    # waiting on that item (Loads_Impacted). Computed here, before trailer priority, so
    # Loads_Impacted can feed directly into the priority ranking below.
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
    # Blends two things equally: how urgent the trailer's demand is (Earliest_Dispatch —
    # sooner is more urgent) and how many currently-short/partial loads it would help fix
    # (Loads_Impacted — more is better). Both are normalized to a 0-1 scale so neither
    # dominates just because of its raw units, then averaged 50/50 into one score.
    # Priority_Score (demand efficiency) remains as the final tiebreaker.
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

    dispatch_min = trailer_priority["Earliest_Dispatch"].min()
    dispatch_max = trailer_priority["Earliest_Dispatch"].max()
    if dispatch_max > dispatch_min:
        trailer_priority["Urgency_Score"] = 1 - (
            (trailer_priority["Earliest_Dispatch"] - dispatch_min) / (dispatch_max - dispatch_min)
        )
    else:
        trailer_priority["Urgency_Score"] = 1.0

    loads_max = trailer_priority["Loads_Impacted"].max()
    if loads_max > 0:
        trailer_priority["Fix_Score"] = trailer_priority["Loads_Impacted"] / loads_max
    else:
        trailer_priority["Fix_Score"] = 0.0

    trailer_priority["Blended_Score"] = (
        0.5 * trailer_priority["Urgency_Score"] + 0.5 * trailer_priority["Fix_Score"]
    )

    trailer_priority = trailer_priority.sort_values(
        by=["Blended_Score", "Priority_Score"],
        ascending=[False, False]
    ).reset_index(drop=True)

    trailer_priority["Wave"] = (trailer_priority.index // 4) + 1

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

    exceptions = load_trailer[load_trailer["Status"] != "Full ✅"].copy()
    exceptions = exceptions.sort_values(by=["Dispatch", "Status"])

    # ---- FORMAT EXPORTS ----
    dock_plan_export = trailer_priority.drop(columns=["Urgency_Score", "Fix_Score", "Blended_Score"]).copy()
    dock_plan_export["Priority_Score"] = dock_plan_export["Priority_Score"].round(0)
    dock_plan_export.insert(0, "Priority", range(1, len(dock_plan_export) + 1))
    cols = ["Priority", "Wave"] + [c for c in dock_plan_export.columns if c not in ("Priority", "Wave")]
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
total_waves = dock_plan_export["Wave"].nunique()
loads_met_count = int(load_export["Status"].eq("Full ✅").sum())
total_loads = len(load_export)

if not top4_trailers.empty:
    top_trailer = top4_trailers.iloc[0]
    move_next_value = f"Trailer {top_trailer['Trailer']}"
    move_next_sub = f"fixes {int(top_trailer['Fix_Cases']):,} cases across {int(top_trailer['Loads_Impacted'])} load(s)"
else:
    move_next_value = "—"
    move_next_sub = "no shortages to fix"

k1, k2, k3, k4, k5 = st.columns(5)
with k1:
    kpi_card("Trailers Involved", f"{total_trailers}", f"carry items needed today, across {total_waves} waves")
with k2:
    kpi_card("Total Demand", f"{total_demand:,}", "cases ordered")
with k3:
    kpi_card("Shortage Cases", f"{total_shortage:,}", "cases still missing")
with k4:
    kpi_card("Loads Fully Met", f"{loads_met_count:,}", f"of {total_loads:,} total load lines")
with k5:
    kpi_card("Move Next", move_next_value, move_next_sub)

# =====================================================================
# TABS
# =====================================================================
tab_overview, tab_wave, tab_load = st.tabs(
    ["Overview", "Wave Plan", "Load Coverage"]
)

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
        priority_order = dock_plan_export.sort_values("Priority")
        priority_order_chart_df = priority_order.copy()
        priority_order_chart_df["Wave"] = priority_order_chart_df["Wave"].astype(str)
        fig2 = px.bar(
            priority_order_chart_df, x="Demand_Served", y="Trailer", orientation="h",
            color="Wave", text="Priority", title="Trailer Priority Order — Today",
            category_orders={"Trailer": priority_order_chart_df["Trailer"].tolist()},
            color_discrete_sequence=[STEEL, AMBER]
        )
        fig2.update_traces(texttemplate="#%{text}", textposition="outside")
        st.plotly_chart(style_fig(fig2), use_container_width=True)
        st.caption("The order to unload trailers today — #1 first. Bar length shows demand tied to that trailer.")
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
            top_skus, x="Item", y="Shortage_Cases", title="Top Shortage Items",
            text="Shortage_Cases", color_discrete_sequence=[DANGER],
            hover_data={"Trailer": True, "Item": False}
        )
        st.plotly_chart(style_fig(fig4), use_container_width=True)
        st.caption("Hover a bar to see which trailer(s), if any, carry that item.")
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------- WAVE PLAN ----------------
with tab_wave:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Dock Plan — Trailers in Priority Order")
    st.caption("Trailers are listed in the order they should be unloaded today.")
    st.caption(
        "Priority blends two things equally: how soon the trailer's demand ships (urgency) and "
        "how many currently-short/partial loads it would help fix (Loads_Impacted). "
        "Priority_Score (Demand Served ÷ SKU Count — cases served per distinct item) is the tiebreaker."
    )
    st.caption(
        "Loads_Impacted = 0 just means that trailer isn't carrying any item currently causing a "
        "shortage — it can still rank high on urgency alone. It doesn't mean the trailer is unneeded, "
        "only that it isn't one of today's shortage-fixers."
    )
    st.dataframe(
        dock_plan_export.sort_values("Priority"),
        use_container_width=True, hide_index=True
    )
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- LOAD COVERAGE ----------------
with tab_load:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Load Coverage — Demand vs. Available by Trip")
    st.caption("Status: green = Full, yellow = Partial, red = Short.")

    def highlight_status(val):
        color_map = {"Full ✅": "#C7F0D8", "Partial ⚠️": "#FFE8A3", "Short ❌": "#F5C2C7"}
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
            "Dock Optimization Results",
            kpis=[
                ("Trailers Involved", str(total_trailers), f"across {total_waves} waves"),
                ("Total Demand", f"{total_demand:,}", "cases ordered"),
                ("Shortage Cases", f"{total_shortage:,}", "cases still missing"),
                ("Loads Fully Met", f"{loads_met_count:,}", f"of {total_loads:,} total load lines"),
                ("Move Next", move_next_value, move_next_sub),
            ],
            figs=[
                ("Top Trailers by Demand", fig1),
                ("Load Status Breakdown", fig3),
                ("Trailer Priority Order — Today", fig2),
                ("Top Shortage Items", fig4),
            ],
            wave_df=dock_plan_export.sort_values("Priority"),
            load_df=load_export
        ),
        file_name="Dock_Optimization_Report.pdf",
        mime="application/pdf"
    )
