"""
4_Monthly_Scorecard.py
=====================
Cumulative monthly view of how well we met our goals. It reads the same persistent
log the shift closeout writes to, so it updates automatically every time a shift is
closed out — nothing extra to save. Pick a month, see the headline goal percentages,
the per-shift breakdown, and every miss. A one-page PDF is available to hand off.
"""

import datetime
import io

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
    """Render a goal metric, handling the no-data case."""
    rate = block.get("rate")
    met = block.get(met_key, 0)
    total = block.get(total_key, 0)
    if rate is None:
        column.metric(label, "—")
        column.caption("No data this month.")
    else:
        column.metric(label, f"{rate}%")
        column.caption(f"{met} of {total}")


def _recent_month_options(count=12):
    """Return [(year, month, label)] for the last `count` months, newest first."""
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


def build_monthly_pdf(score, month_label):
    """One-page monthly goal scorecard PDF. Returns bytes, or None."""
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
        Paragraph("Monthly Goal Scorecard", title_style),
        Paragraph(
            f"{month_label} &nbsp;|&nbsp; {score['shifts_logged']} shift(s) closed out",
            sub_style,
        ),
    ]

    def _cell(block, met_key, total_key):
        rate = block.get("rate")
        if rate is None:
            return "No data"
        return f"{rate}%  ({block.get(met_key, 0)} of {block.get(total_key, 0)})"

    goal_data = [
        ["Goal Area", "This Month"],
        ["OC Sign-Off Compliance", _cell(score["oc_signoff"], "met", "required")],
        ["OC Photo Compliance", _cell(score["oc_photos"], "met", "required")],
        ["CPU On-Time", _cell(score["cpu_on_time"], "met", "total")],
        ["Shift Goal Met", _cell(score["shift_goal"], "met", "total")],
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

    story.append(Paragraph("Misses this month", h_style))
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
        story.append(Paragraph("No misses recorded this month.", body))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


# ============================================================
#  PAGE
# ============================================================

st.set_page_config(page_title="Monthly Scorecard", layout="wide")
st.title("Monthly Goal Scorecard")
st.write(
    "How much of our goals we've met this month. This updates automatically every "
    "time a shift is closed out."
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

options = _recent_month_options(12)
labels = [o[2] for o in options]
picked_label = st.selectbox("Month", labels, index=0)
year, month, _ = next(o for o in options if o[2] == picked_label)

try:
    score = shift_log.get_monthly_scorecard(year, month)
except Exception as e:
    st.error(f"Could not load the monthly scorecard: {e}")
    st.stop()

if score["shifts_logged"] == 0:
    st.warning(f"No shifts have been closed out in {picked_label} yet.")
    st.stop()

st.caption(f"Based on {score['shifts_logged']} shift(s) closed out in {picked_label}.")

# Headline goal metrics
m1, m2, m3, m4 = st.columns(4)
_metric(m1, "OC Sign-Off", score["oc_signoff"], "met", "required")
_metric(m2, "OC Photos", score["oc_photos"], "met", "required")
_metric(m3, "CPU On-Time", score["cpu_on_time"], "met", "total")
_metric(m4, "Shift Goal Met", score["shift_goal"], "met", "total")

# Month totals
t1, t2, t3 = st.columns(3)
t1.metric("Shifts closed out", score["shifts_logged"])
t2.metric("Loads completed (total)", score["loads_completed_total"])
t3.metric("Shorts (total)", score["shorts_total"])

# Per-shift breakdown
st.markdown("---")
st.subheader("Shift-by-shift this month")
if score["per_shift"]:
    st.dataframe(
        pd.DataFrame([
            {
                "Date": r.get("date"),
                "Shift": r.get("shift"),
                "Goal met": r.get("goal_met"),
                "Loads done": r.get("loads_completed"),
                "Shorts": r.get("total_shorts"),
                "OC sign-off met": r.get("oc_signoff_met"),
                "OC photos met": r.get("oc_photos_met"),
                "CPU on time": f"{r.get('cpu_on_time')}/{r.get('cpu_total')}",
                "Notes": r.get("notes"),
            }
            for r in score["per_shift"]
        ]),
        use_container_width=True,
        height=320,
    )
else:
    st.caption("No per-shift summary rows for this month.")

# Misses
st.markdown("---")
st.subheader("Misses this month")
if score["misses"]:
    st.dataframe(
        pd.DataFrame([
            {
                "Date": m.get("date"), "Shift": m.get("shift"),
                "Type": m.get("type"), "Load": m.get("load"),
                "Customer": m.get("customer"), "Appt": m.get("appt_time"),
                "On time": m.get("on_time"), "Sign-off": m.get("signoff_done"),
                "Photos": m.get("photos_done"), "Short": m.get("short"),
                "Reason": m.get("miss_reason"),
            }
            for m in score["misses"]
        ]),
        use_container_width=True,
        height=300,
    )
else:
    st.success("No misses recorded this month.")

# Download
st.markdown("---")
pdf_bytes = build_monthly_pdf(score, picked_label)
if pdf_bytes:
    st.download_button(
        "Download Monthly Scorecard (PDF)",
        data=pdf_bytes,
        file_name=f"Monthly_Scorecard_{year}-{month:02d}.pdf",
        mime="application/pdf",
    )
elif not REPORTLAB_AVAILABLE:
    st.caption("PDF download unavailable — add reportlab to requirements.txt to enable it.")
