"""
4_Performance_Scorecard.py
==========================
Daily, weekly, or monthly view of how well we met our goals, from the same
persistent log the shift closeout writes to. Pick Day / Week / Month and a date,
see the same KPIs the closeout uses (OC Service Target, OC Shorts Target,
CPU Service Target, Daily Goal Met), the top reasons loads missed requirements,
the per-shift breakdown, and every miss. A one-page PDF of the selected period
is available to hand to the DC manager.
"""

import datetime
import io
from collections import Counter

import pandas as pd
import streamlit as st

import shift_log

try:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER
    from reportlab.lib.pagesizes import letter
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.platypus import (
        SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    )
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False


# ============================================================
#  HELPERS
# ============================================================

def _metric(column, label, block, met_key, total_key):
    rate = block.get("rate")
    met = block.get(met_key, 0)
    total = block.get(total_key, 0)
    if rate is None:
        column.metric(label, "—")
        column.caption("No data this period.")
    else:
        column.metric(label, f"{rate}%")
        column.caption(f"{met} of {total}")


def _oc_shorts_block(score):
    """
    OC Shorts Target, computed the same way the closeout's rolling scorecard does it:
    scorable OC loads (the OC Service Target denominator) minus OC loads flagged short,
    over scorable OC loads. Returns a {rate, met, total} block.
    """
    oc_service = score.get("oc_signoff") or {}
    total = oc_service.get("required", oc_service.get("total", 0)) or 0
    try:
        total = int(total)
    except Exception:
        total = 0
    if total <= 0:
        return {"rate": None, "met": 0, "total": 0}

    shorts = 0
    for row in score.get("misses", []):
        if str(row.get("type", "")).strip().upper() != "OC":
            continue
        if str(row.get("short", "")).strip().upper() in ("Y", "YES"):
            shorts += 1

    met = max(total - shorts, 0)
    rate = round(100 * met / total)
    return {"rate": rate, "met": met, "total": total}


def _recent_month_options(count=12):
    today = datetime.date.today()
    out = []
    y, m = today.year, today.month
    for _ in range(count):
        out.append((y, m, datetime.date(y, m, 1).strftime("%B %Y")))
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    return out


def rank_top_miss_reasons(misses, top_n=3):
    """Count standardized miss reasons and return [(reason, count), ...] + total."""
    counter = Counter()
    for m in misses or []:
        reason = str(m.get("miss_reason", "") or "").strip() or "Unspecified"
        counter[reason] += 1
    return counter.most_common(top_n), sum(counter.values())


def build_scorecard_pdf(score, period_label, top_reasons, total_miss_events):
    """One-page scorecard PDF for any period (day/week/month). Returns bytes, or None."""
    if not REPORTLAB_AVAILABLE:
        return None

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=letter,
        rightMargin=0.6 * inch, leftMargin=0.6 * inch,
        topMargin=0.6 * inch, bottomMargin=0.6 * inch,
    )
    base = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "T", parent=base["Title"], fontSize=18, alignment=TA_CENTER,
        textColor=colors.HexColor("#0F5B78"), spaceAfter=4,
    )
    sub_style = ParagraphStyle(
        "S", parent=base["Normal"], fontSize=10, alignment=TA_CENTER,
        textColor=colors.HexColor("#555555"), spaceAfter=14,
    )
    h_style = ParagraphStyle(
        "H", parent=base["Heading2"], fontSize=12,
        textColor=colors.HexColor("#0F5B78"), spaceBefore=10, spaceAfter=6,
    )
    body = ParagraphStyle("B", parent=base["Normal"], fontSize=9, leading=12)

    story = [
        Paragraph("Performance Scorecard", title_style),
        Paragraph(f"{period_label} &nbsp;|&nbsp; {score['shifts_logged']} closeout(s) recorded", sub_style),
    ]

    def _cell(block, met_key, total_key):
        rate = block.get("rate")
        if rate is None:
            return "No data"
        return f"{rate}%  ({block.get(met_key, 0)} of {block.get(total_key, 0)})"

    oc_shorts = _oc_shorts_block(score)

    goal_data = [
        ["Goal Area", "This Period"],
        ["OC Service Target (<= 120 min)", _cell(score["oc_signoff"], "met", "required")],
        ["OC Shorts Target (0 short)", _cell(oc_shorts, "met", "total")],
        ["CPU Service Target (<= 120 min)", _cell(score["cpu_on_time"], "met", "total")],
        ["Daily Goal Met", _cell(score["shift_goal"], "met", "total")],
        ["Loads Completed (total)", str(score["loads_completed_total"])],
        ["Shorts (total)", str(score["shorts_total"])],
    ]
    goal_table = Table(goal_data, colWidths=[3.2 * inch, 3.9 * inch])
    goal_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F5B78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7F9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(goal_table)

    # Top reasons loads missed requirements
    story.append(Paragraph("Top reasons loads missed requirements", h_style))
    if top_reasons:
        rdata = [["#", "Reason", "Count", "% of misses"]]
        for i, (reason, count) in enumerate(top_reasons, 1):
            share = f"{round(100 * count / total_miss_events)}%" if total_miss_events else "—"
            rdata.append([str(i), Paragraph(str(reason), body), str(count), share])
        rt = Table(rdata, colWidths=[0.4 * inch, 4.5 * inch, 0.9 * inch, 1.3 * inch], repeatRows=1)
        rt.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#C00000")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 9),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(rt)
    else:
        story.append(Paragraph("No misses recorded this period.", body))

    # Full miss list
    story.append(Paragraph("Misses this period", h_style))
    if score["misses"]:
        miss_data = [["Date", "Shift", "Type", "Load", "Customer", "Reason"]]
        for m in score["misses"]:
            miss_data.append([
                Paragraph(str(m.get("date", "")), body),
                Paragraph(str(m.get("shift", "")), body),
                Paragraph(str(m.get("type", "")), body),
                Paragraph(str(m.get("load", "")), body),
                Paragraph(str(m.get("customer", "")), body),
                Paragraph(str(m.get("miss_reason", "") or "—"), body),
            ])
        miss_table = Table(
            miss_data,
            colWidths=[0.8 * inch, 0.55 * inch, 0.6 * inch, 0.8 * inch, 1.9 * inch, 2.45 * inch],
            repeatRows=1,
        )
        miss_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#C00000")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(miss_table)
    else:
        story.append(Paragraph("No misses recorded this period.", body))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
#  PAGE
# ============================================================

st.set_page_config(page_title="Performance Scorecard", layout="wide")
st.title("Performance Scorecard")
st.write(
    "How much of our goals we've met. Pick a day, a week, or a month. This updates "
    "automatically every time a shift is closed out, and uses the same KPIs as the "
    "daily closeout: OC Service Target, OC Shorts Target, CPU Service Target, and "
    "Daily Goal Met."
)

if not shift_log.is_configured():
    st.error(
        "The shift log isn't connected yet, so there's no data to show. "
        f"Reason: {shift_log.setup_hint()}"
    )
    st.info(
        "See the setup steps at the top of shift_log.py: create a Google Sheet, "
        "add a service account, and put the credentials + sheet ID in Streamlit secrets."
    )
    st.stop()

period = st.radio("Period", ["Day", "Week", "Month"], horizontal=True)

if period == "Day":
    day = st.date_input("Operating date", value=datetime.date.today())
    try:
        score = shift_log.get_daily_scorecard(day)
    except Exception as e:
        st.error(f"Could not load the daily scorecard: {e}")
        st.stop()
    period_label = score["date"]
    file_tag = f"Daily_{score['date'].replace('/', '-')}"

elif period == "Week":
    end = st.date_input("Week ending", value=datetime.date.today())
    try:
        score = shift_log.get_weekly_scorecard(end)
    except Exception as e:
        st.error(f"Could not load the weekly scorecard: {e}")
        st.stop()
    period_label = f"{score['start']} – {score['end']}"
    file_tag = f"Weekly_{score['end'].replace('/', '-')}"

else:
    options = _recent_month_options(12)
    labels = [o[2] for o in options]
    picked_label = st.selectbox("Month", labels, index=0)
    year, month, _ = next(o for o in options if o[2] == picked_label)
    try:
        score = shift_log.get_monthly_scorecard(year, month)
    except Exception as e:
        st.error(f"Could not load the monthly scorecard: {e}")
        st.stop()
    period_label = picked_label
    file_tag = f"Monthly_{year}-{month:02d}"

if score["shifts_logged"] == 0:
    st.warning(f"No closeouts have been recorded for {period_label} yet.")
    st.stop()

st.caption(f"Based on {score['shifts_logged']} closeout(s) — {period_label}.")

# Headline KPIs — matched to the daily closeout's Rolling Scorecard.
oc_shorts_block = _oc_shorts_block(score)
m1, m2, m3, m4 = st.columns(4)
_metric(m1, "OC Service Target", score["oc_signoff"], "met", "required")
_metric(m2, "OC Shorts Target", oc_shorts_block, "met", "total")
_metric(m3, "CPU Service Target", score["cpu_on_time"], "met", "total")
_metric(m4, "Daily Goal Met", score["shift_goal"], "met", "total")

# Period totals
t1, t2, t3 = st.columns(3)
t1.metric("Closeouts recorded", score["shifts_logged"])
t2.metric("Loads completed (total)", score["loads_completed_total"])
t3.metric("Shorts (total)", score["shorts_total"])

# Top 3 reasons loads missed requirements
top_reasons, total_miss_events = rank_top_miss_reasons(score["misses"], top_n=3)
st.markdown("---")
st.subheader("Top 3 reasons loads missed requirements")
if top_reasons:
    st.table(pd.DataFrame([
        {"Reason": r, "Count": c,
         "% of misses": f"{round(100 * c / total_miss_events)}%" if total_miss_events else "—"}
        for r, c in top_reasons
    ]))
else:
    st.success("No misses recorded this period.")

# Per-shift breakdown
st.markdown("---")
st.subheader("Closeout-by-closeout")
if score["per_shift"]:
    st.dataframe(
        pd.DataFrame([
            {
                "Date": r.get("date"),
                "Shift": r.get("shift"),
                "Goal met": r.get("goal_met"),
                "Loads done": r.get("loads_completed"),
                "Shorts": r.get("total_shorts"),
                "OC service met": r.get("oc_signoff_met"),
                "OC loads": r.get("oc_total"),
                "CPU service": f"{r.get('cpu_on_time')}/{r.get('cpu_total')}",
                "Notes": r.get("notes"),
            }
            for r in score["per_shift"]
        ]),
        use_container_width=True,
        height=320,
    )
else:
    st.caption("No per-closeout summary rows for this period.")

# Full miss list
st.markdown("---")
st.subheader("Misses")
if score["misses"]:
    st.dataframe(
        pd.DataFrame([
            {
                "Date": m.get("date"), "Shift": m.get("shift"),
                "Type": m.get("type"), "Load": m.get("load"),
                "Customer": m.get("customer"), "Appt": m.get("appt_time"),
                "On time": m.get("on_time"), "OC service": m.get("signoff_done"),
                "Short": m.get("short"),
                "Reason": m.get("miss_reason"),
            }
            for m in score["misses"]
        ]),
        use_container_width=True,
        height=300,
    )
else:
    st.success("No misses recorded this period.")

# Download the scorecard for the selected period
st.markdown("---")
pdf_bytes = build_scorecard_pdf(score, period_label, top_reasons, total_miss_events)
if pdf_bytes:
    st.download_button(
        f"Download {period} Scorecard (PDF)",
        data=pdf_bytes,
        file_name=f"Performance_Scorecard_{file_tag}.pdf",
        mime="application/pdf",
    )
elif not REPORTLAB_AVAILABLE:
    st.caption("PDF download unavailable — add reportlab to requirements.txt to enable it.")
