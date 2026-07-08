import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import Alignment, PatternFill
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
LIGHT_PURPLE = "#E6D9F2"
BG = "#F5F6F8"
CARD = "#FFFFFF"
TEXT_MUTED = "#6B7280"

STATUS_COLORS = {
    "Full": SUCCESS,
    "Partial": WARNING,
    "Short": DANGER,
}


def inject_dashboard_style():
    """Professional styling — only applied to the results dashboard,
    never to the upload screen, so that one stays plain/default."""
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


def get_multi_trailer_items(df):
    """Return Item values that appear in Load Coverage from more than one real trailer."""
    if not {"Item", "Trailer"}.issubset(df.columns):
        return set()

    tmp = df[["Item", "Trailer"]].copy()
    tmp["Trailer"] = tmp["Trailer"].astype(str).str.strip()
    tmp = tmp[tmp["Trailer"].ne("No trailer found")]

    multi = tmp.groupby("Item")["Trailer"].nunique()
    return set(multi[multi > 1].index.astype(str))


def build_status_table(df, status_col="Status", repeated_items=None):
    """Reportlab Table with Status cells colored and repeated multi-trailer Item cells light purple."""
    repeated_items = set() if repeated_items is None else set(map(str, repeated_items))

    data = [list(df.columns)] + df.astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    style_commands = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(NAVY)),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]

    # Light purple highlight only on the Item cells where that item is supplied
    # by more than one trailer in the Load Coverage page.
    if repeated_items and "Item" in df.columns:
        item_col_idx = list(df.columns).index("Item")
        for i, item_value in enumerate(df["Item"].astype(str).tolist()):
            if item_value in repeated_items:
                style_commands.append((
                    "BACKGROUND",
                    (item_col_idx, i + 1),
                    (item_col_idx, i + 1),
                    colors.HexColor(LIGHT_PURPLE),
                ))

    if status_col in df.columns:
        col_idx = list(df.columns).index(status_col)
        for i, val in enumerate(df[status_col].astype(str).tolist()):
            color = STATUS_PDF_COLORS.get(val)
            if color:
                style_commands.append(("BACKGROUND", (col_idx, i + 1), (col_idx, i + 1), color))
    table.setStyle(TableStyle(style_commands))
    return table

def build_full_pdf(title, kpis, figs, wave_df, load_df, repeated_items=None):
    """One combined PDF: Overview, Wave Plan, and Load Coverage."""
    buffer = BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=portrait(letter),
        topMargin=30, bottomMargin=30, leftMargin=30, rightMargin=30
    )
    styles = getSampleStyleSheet()
    story = []

    # ---- Page 1: Overview ----
    overview_page = []

    overview_page.append(Paragraph(f"{title} — Overview", styles["Title"]))
    overview_page.append(Spacer(1, 14))

    header_row = [k[0] for k in kpis]
    value_row = [k[1] for k in kpis]
    sub_row = [k[2] for k in kpis]

    kpi_table = Table([header_row, value_row, sub_row], colWidths=[88] * len(kpis))
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
    story.append(build_status_table(load_df, repeated_items=repeated_items))

    doc.build(story)
    buffer.seek(0)
    return buffer


# =====================================================================
# HEADER
# =====================================================================
st.title("Shorts analysis Tool")
st.write("Turn incoming trailer inventory and outbound order shorts into a prioritized unloading plan.")

# =====================================================================
# INPUTS
# =====================================================================
st.markdown("### Data Inputs")
col1, col2 = st.columns(2)
with col1:
    book_file = st.file_uploader("Trailer inventory", type=["xlsx"])
with col2:
    short_file = st.file_uploader("Order shorts (short sheet.xlsx)", type=["xlsx"])

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
    # ---- LOAD TRAILER INVENTORY (BOOK) ----
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

    # Total transfer trailers on the lot today = unique trailers in the inventory
    # (transfers) file, counted from valid inventory rows only.
    total_transfer_trailers = int(clean_df["Trailer"].nunique())

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

    # Forward-fill Trip and Dispatch: the short sheet only writes the Trip number
    # and Dispatch time on the first item row of each trip — continuation rows have
    # blank spaces. Strip spaces, treat blank as NaN, then ffill so every item row
    # inherits its parent trip's values.
    short_clean["Trip"] = short_clean["Trip"].astype(str).str.strip().replace("", None)
    short_clean["Trip"] = short_clean["Trip"].ffill()
    short_clean["Dispatch"] = short_clean["Dispatch"].ffill()

    # =====================================================================
    # CORE MODEL (corrected)
    # Every line on the short sheet is a SHORTAGE — cases an outbound load needs
    # that are not on the pick line yet. The cases that fix them live in the
    # transfer trailers (the inventory file). So:
    #   DEMAND  = short-sheet cases
    #   SUPPLY  = trailer inventory
    # We allocate each item's trailer stock to its short-sheet lines in DISPATCH
    # order (earliest dispatch filled first). A trailer's Fix_Cases is the number
    # of cases it actually contributes toward covering shortages — capped at what
    # is demanded. Every trailer that carries a short item therefore fixes cases.
    # =====================================================================

    # Trailer stock per item (sum across all LPNs of the same SKU on that trailer).
    trailer_stock = clean_df.groupby(["SKU", "Trailer"], as_index=False)["Quantity"].sum()
    trailer_stock = trailer_stock.rename(columns={"SKU": "Item"})

    # Total inventory available per item, across all trailers.
    item_totals = clean_df.groupby("SKU", as_index=False)["Quantity"].sum()
    item_totals = item_totals.rename(columns={"SKU": "Item", "Quantity": "Total_Item_Inventory"})

    fix_rows = []           # trailer-level: how many cases each trailer contributes to each short line
    load_rows = []          # short-line level: demand, allocated, short, status, primary source trailer
    load_detail_rows = []   # load coverage detail: one row per trailer contribution to a short line

    # Process each item independently. Within an item, fill the earliest-dispatch
    # short line first, drawing from trailers (largest stock first as the source order).
    demand_sorted = short_clean.sort_values(["Item", "Dispatch", "Trip"])

    for item, demand_group in demand_sorted.groupby("Item"):
        stock = trailer_stock[trailer_stock["Item"] == item].sort_values(
            "Quantity", ascending=False
        )
        stock_remaining = dict(zip(stock["Trailer"], stock["Quantity"].astype(float)))
        total_item_inv = float(item_totals.loc[item_totals["Item"] == item, "Total_Item_Inventory"].sum())

        for _, line in demand_group.iterrows():
            need = float(line["Cases"])
            allocated_total = 0.0
            sources = []

            for trailer in list(stock_remaining.keys()):
                if need - allocated_total <= 0:
                    break
                avail = stock_remaining[trailer]
                if avail <= 0:
                    continue
                take = min(avail, need - allocated_total)
                if take <= 0:
                    continue
                stock_remaining[trailer] = avail - take
                allocated_total += take

                # Keep the exact trailer contribution so Load Coverage can show
                # one separate row per trailer used on the same load/item.
                sources.append({
                    "Trailer": trailer,
                    "Allocated_Cases": take,
                })

                fix_rows.append({
                    "Trailer": trailer,
                    "Item": item,
                    "Trip": line["Trip"],
                    "Dispatch": line["Dispatch"],
                    "Allocated_Cases": take,
                })

            short_qty = need - allocated_total
            fill_rate = (allocated_total / need) if need > 0 else 0.0
            status = "Full" if fill_rate >= 1 else ("Partial" if fill_rate > 0 else "Short")
            primary_trailer = sources[0]["Trailer"] if sources else np.nan

            # Line-level summary, kept for KPIs, charts, and load status math.
            load_rows.append({
                "Item": item,
                "Trip": line["Trip"],
                "Dispatch": line["Dispatch"],
                "Demand_Cases": need,
                "Allocated_Cases": allocated_total,
                "Total_Item_Inventory": total_item_inv,
                "Fill_Rate": fill_rate,
                "Status": status,
                "Actual_short_cases": short_qty,
                "Trailer": primary_trailer,
            })

            # Detail rows for the Load Coverage page/export.
            # If two trailers cover the same short line, this creates two rows.
            # Example: 3 cases from 198 and 9 cases from 605 = two rows.
            if sources:
                for source in sources:
                    load_detail_rows.append({
                        "Item": item,
                        "Trip": line["Trip"],
                        "Dispatch": line["Dispatch"],
                        "Demand_Cases": need,
                        "Allocated_Cases": source["Allocated_Cases"],
                        "Total_Item_Inventory": total_item_inv,
                        "Fill_Rate": fill_rate,
                        "Status": status,
                        "Actual_short_cases": short_qty,
                        "Trailer": source["Trailer"],
                    })
            else:
                load_detail_rows.append({
                    "Item": item,
                    "Trip": line["Trip"],
                    "Dispatch": line["Dispatch"],
                    "Demand_Cases": need,
                    "Allocated_Cases": 0.0,
                    "Total_Item_Inventory": total_item_inv,
                    "Fill_Rate": fill_rate,
                    "Status": status,
                    "Actual_short_cases": short_qty,
                    "Trailer": np.nan,
                })

    fix_df = pd.DataFrame(fix_rows)
    alloc = pd.DataFrame(load_rows)
    alloc_detail = pd.DataFrame(load_detail_rows)

    _matched = int(alloc["Allocated_Cases"].gt(0).sum())
    st.success(f"{_matched} of {len(alloc)} short lines can be covered (fully or partially) from trailer inventory.")

    # =====================================================================
    # TRAILER-LEVEL SUMMARY (Fix_Cases = cases actually pulled to cover shortages)
    # =====================================================================
    if fix_df.empty:
        optimized_trailers = pd.DataFrame(columns=["Trailer", "Fix_Cases", "Loads_Impacted"])
    else:
        optimized_trailers = fix_df.groupby("Trailer").agg(
            Fix_Cases=("Allocated_Cases", "sum"),
            Loads_Impacted=("Trip", "nunique"),
        ).reset_index().sort_values(by=["Fix_Cases", "Loads_Impacted"], ascending=[False, False])

    top4_trailers = optimized_trailers.head(4).copy()
    if not top4_trailers.empty:
        top4_trailers.insert(0, "Rank", range(1, len(top4_trailers) + 1))

    # =====================================================================
    # TRAILER PRIORITY / WAVE PLAN
    # Earliest dispatch of any short line the trailer helps drives priority.
    # Ties broken by most Fix_Cases, then most loads impacted.
    # =====================================================================
    if fix_df.empty:
        trailer_priority = pd.DataFrame(columns=[
            "Trailer", "Fix_Cases", "Loads_Impacted", "Earliest_Dispatch", "Demand_Served", "SKU_Count"
        ])
    else:
        # Earliest dispatch among the short lines each trailer actually helps.
        earliest_dispatch = fix_df.groupby("Trailer")["Dispatch"].min().rename("Earliest_Dispatch")
        # Total cases the trailer contributes and how many distinct SKUs it helps.
        demand_served = fix_df.groupby("Trailer")["Allocated_Cases"].sum().rename("Demand_Served")
        sku_count = fix_df.groupby("Trailer")["Item"].nunique().rename("SKU_Count")

        trailer_priority = optimized_trailers.merge(
            earliest_dispatch, on="Trailer", how="left"
        ).merge(
            demand_served, on="Trailer", how="left"
        ).merge(
            sku_count, on="Trailer", how="left"
        )

    if not trailer_priority.empty:
        trailer_priority = trailer_priority.sort_values(
            by=["Earliest_Dispatch", "Fix_Cases", "Loads_Impacted"],
            ascending=[True, False, False]
        ).reset_index(drop=True)

        # Waves = groups of 4 in priority order.
        trailer_priority["Trailer_Priority"] = range(1, len(trailer_priority) + 1)
        trailer_priority["Wave"] = ((trailer_priority["Trailer_Priority"] - 1) // 4) + 1
    else:
        trailer_priority["Trailer_Priority"] = []
        trailer_priority["Wave"] = []

    # =====================================================================
    # LOAD COVERAGE TABLE
    # =====================================================================
    load_trailer = alloc_detail.merge(
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
    # Wave Plan: every trailer here fixes shortage cases (Fix_Cases > 0 by construction).
    dock_plan_export = trailer_priority.drop(columns=["Trailer_Priority"]).copy()
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

    # Items that repeat in Load Coverage because they are supplied by different trailers.
    repeated_multi_trailer_items = get_multi_trailer_items(load_export)

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
total_cases_short = int(alloc["Demand_Cases"].sum())
total_shortage = int(alloc["Actual_short_cases"].clip(lower=0).sum())
total_waves = int(dock_plan_export["Wave"].nunique()) if not dock_plan_export.empty else 0

trip_status = alloc.groupby("Trip")["Status"].apply(lambda s: (s == "Full").all())
loads_met_count = int(trip_status.sum())
total_loads = int(trip_status.shape[0])

# Move Next should match the first trailer in the Wave Plan.
if not dock_plan_export.empty:
    next_trailer = dock_plan_export.iloc[0]
    move_next_value = f"Trailer {next_trailer['Trailer']}"
    move_next_sub = (
        f"fixes {int(next_trailer['Fix_Cases']):,} cases "
        f"across {int(next_trailer['Loads_Impacted'])} load(s)"
    )
else:
    move_next_value = "—"
    move_next_sub = "no shortages to fix"

k1, k2, k3, k4, k5, k6 = st.columns(6)
with k1:
    kpi_card("Transfer Trailers", f"{total_transfer_trailers}", "on the lot today")
with k2:
    kpi_card("Trailers Involved", f"{total_trailers}", f"fix shortages, across {total_waves} waves")
with k3:
    kpi_card("Total Cases Short", f"{total_cases_short:,}", "cases on the short sheet")
with k4:
    kpi_card("Still Short", f"{total_shortage:,}", "cases no trailer can cover")
with k5:
    kpi_card("Loads Fully Met", f"{loads_met_count:,}", f"of {total_loads:,} loads coverable in full")
with k6:
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
            top_trailers_chart_df, x="Trailer", y="Fix_Cases",
            color="Wave", title="Top Trailers by Shortage Cases Fixed", text="Fix_Cases",
            color_discrete_sequence=[STEEL, AMBER]
        )
        st.plotly_chart(style_fig(fig1), use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with c2:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        status_summary = alloc.groupby("Status").size().reset_index(name="Count")
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
        priority_order_chart_df = dock_plan_export.sort_values("Wave").copy()
        priority_order_chart_df["Wave"] = priority_order_chart_df["Wave"].astype(str)
        fig2 = px.bar(
            priority_order_chart_df, x="Fix_Cases", y="Trailer", orientation="h",
            color="Wave", title="Trailer Priority Order — Today",
            category_orders={"Trailer": priority_order_chart_df["Trailer"].tolist()},
            color_discrete_sequence=[STEEL, AMBER]
        )
        st.plotly_chart(style_fig(fig2), use_container_width=True)
        st.caption("Earliest dispatch drives the ranking; shortage cases fixed breaks ties.")
        st.markdown('</div>', unsafe_allow_html=True)

    with c4:
        st.markdown('<div class="section-card">', unsafe_allow_html=True)
        short_items = alloc[alloc["Actual_short_cases"] > 0]
        top_skus = short_items.groupby("Item").agg(
            Actual_short_cases=("Actual_short_cases", "sum"),
            Trailer=("Trailer", lambda s: ", ".join(sorted(set(x for x in s if pd.notna(x) and x != ""))) or "No inventory found")
        ).reset_index()
        top_skus = top_skus.sort_values(by="Actual_short_cases", ascending=False).head(10)
        fig4 = px.bar(
            top_skus, x="Item", y="Actual_short_cases", title="Top Still-Short Items",
            text="Actual_short_cases", color_discrete_sequence=[DANGER],
            hover_data={"Trailer": True, "Item": False}
        )
        st.plotly_chart(style_fig(fig4), use_container_width=True)
        st.caption("Items still short after all trailer inventory is allocated.")
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------- WAVE PLAN ----------------
with tab_wave:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Dock Plan — Trailers in Priority Order")
    st.caption(
        "Every trailer here carries cases that fix a shortage. "
        "Priority: earliest dispatch on a short load first; ties broken by most cases fixed. "
        "Waves are groups of 4."
    )
    st.dataframe(dock_plan_export, use_container_width=True, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)

# ---------------- LOAD COVERAGE ----------------
with tab_load:
    st.markdown('<div class="section-card">', unsafe_allow_html=True)
    st.subheader("Load Coverage — Demand vs. Available by Trip")
    st.caption("Each row is one trailer contribution to one load/item. Status: green = Full, yellow = Partial, red = Short.")

    def highlight_load_coverage(row):
        styles = [""] * len(row)
        columns = list(row.index)

        if "Status" in columns:
            status_idx = columns.index("Status")
            color_map = {"Full": "#C7F0D8", "Partial": "#FFE8A3", "Short": "#F5C2C7"}
            status_color = color_map.get(row["Status"], "")
            if status_color:
                styles[status_idx] = f"background-color: {status_color}"

        if "Item" in columns and str(row["Item"]) in repeated_multi_trailer_items:
            item_idx = columns.index("Item")
            styles[item_idx] = f"background-color: {LIGHT_PURPLE}; font-weight: 700"

        return styles

    styled_load = load_export.style.apply(highlight_load_coverage, axis=1)
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

        # Light purple highlight on repeated Item cells in the Load Coverage sheet.
        ws = writer.sheets.get("Load Coverage")
        if ws is not None:
            headers = [cell.value for cell in ws[1]]
            if "Item" in headers:
                item_col = headers.index("Item") + 1
                purple_fill = PatternFill(start_color="E6D9F2", end_color="E6D9F2", fill_type="solid")
                for row_num in range(2, ws.max_row + 1):
                    item_value = str(ws.cell(row=row_num, column=item_col).value)
                    if item_value in repeated_multi_trailer_items:
                        ws.cell(row=row_num, column=item_col).fill = purple_fill
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
                ("Transfer Trailers", str(total_transfer_trailers), "on the lot today"),
                ("Trailers Involved", str(total_trailers), f"across {total_waves} waves"),
                ("Cases Short", f"{total_cases_short:,}", "on short sheet"),
                ("Still Short", f"{total_shortage:,}", "no trailer covers"),
                ("Loads Fully Met", f"{loads_met_count:,}", f"of {total_loads:,} loads"),
                ("Move Next", move_next_value, move_next_sub),
            ],
            figs=[
                ("Top Trailers by Fix Cases", fig1),
                ("Load Status Breakdown", fig3),
                ("Trailer Priority Order — Today", fig2),
                ("Top Still-Short Items", fig4),
            ],
            wave_df=dock_plan_export,
            load_df=load_export,
            repeated_items=repeated_multi_trailer_items
        ),
        file_name="Shorts_Analysis_Report.pdf",
        mime="application/pdf"
    )
