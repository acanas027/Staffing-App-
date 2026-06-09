"""
3_Shift_Closeout.py
===================
End-of-shift screen. The supervisor opens this when the shift is wrapping up.

It loads the commitments + shift goal snapshotted during the morning run, asks the
supervisor to confirm what actually happened, writes the result to the persistent
log, and produces a one-page End-of-Shift report comparing expectations vs actual.

The supervisor uploads nothing. The only inputs are confirmations (mostly taps) of
the soft facts a spreadsheet can't see, plus loads completed, shorts, and notes.
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


YES_NO = ["Yes", "No"]

# Standardized miss/late reasons. "Other" reveals a free-text box so nothing is lost,
# but every common cause is now countable across shifts.
MISS_REASONS = [
    "(none)",
    "Late inbound / product not received",
    "Short product / inventory shortage",
    "No driver / carrier no-show",
    "Labor gap / short-staffed",
    "Equipment / dock issue",
    "Dock congestion / staging",
    "Other (explain)",
]


# ============================================================
#  HELPERS
# ============================================================

def _norm_na(value):
    """Map a Yes/No/NA selection to Y / N / NA."""
    v = str(value).strip().upper()
    if v in ("YES", "Y"):
        return "Y"
    if v in ("NO", "N"):
        return "N"
    return "NA"

def _resolve_miss_reason(choice, other_text):
    """Turn the dropdown selection into the stored reason string."""
    if choice in ("(none)", "", None):
        return ""
    if choice == "Other (explain)":
        return (other_text or "").strip() or "Other (unspecified)"
    return choice

def _yn_or_na(shipped, on_time):
    """On-time only means something if the load shipped; otherwise NA."""
    if str(shipped).strip().upper() not in ("YES", "Y"):
        return "NA"
    return _norm_na(on_time)

def _appt_minutes(appt_time):
    """Parse an appt like '18:00' or '6:00' into minutes since midnight. None if unparseable."""
    import re
    text = str(appt_time or "").strip()
    m = re.match(r"^(\d{1,2}):(\d{2})", text)
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def _in_shift_window(appt_time, shift):
    """
    Keep only loads whose appointment falls in this shift's window.
    1st shift: 06:00–16:00 (inclusive of 06:00, up to 16:00).
    2nd shift: 17:00–05:00 next day (wraps midnight).
    Loads with no parseable appt time are kept (can't confidently exclude them).
    """
    mins = _appt_minutes(appt_time)
    if mins is None:
        return True  # don't drop a commitment just because the time is blank/odd

    if str(shift).strip() == "1st":
        return 6 * 60 <= mins <= 16 * 60          # 360..960
    # 2nd shift wraps midnight: 17:00..23:59 OR 00:00..05:00
    return mins >= 17 * 60 or mins <= 5 * 60      # >=1020 or <=300


def _build_summary(outcome_rows, loads_completed, total_shorts, goal_met, shift_goal, notes):
    """Roll per-commitment outcomes into the one-row shift summary."""
    oc = [o for o in outcome_rows if o.get("type") == "OC"]
    cpu = [o for o in outcome_rows if o.get("type") == "CPU"]

    return {
        "loads_completed": loads_completed,
        "total_shorts": total_shorts,
        "goal_met": _norm_na(goal_met),
        "shift_goal": shift_goal,
        "oc_total": len(oc),
        "oc_signoff_met": sum(1 for o in oc if o.get("signoff_done") == "Y"),
        "oc_photos_met": sum(1 for o in oc if o.get("photos_done") == "Y"),
        "cpu_total": len(cpu),
        "cpu_on_time": sum(1 for o in cpu if o.get("on_time") == "Y"),
        "notes": notes,
    }


def _metric(column, label, block, met_key, total_key):
    """Render a compliance metric, handling the no-data case."""
    rate = block.get("rate")
    met = block.get(met_key, 0)
    total = block.get(total_key, 0)
    if rate is None:
        column.metric(label, "—")
        column.caption("No data yet.")
    else:
        column.metric(label, f"{rate}%")
        column.caption(f"{met} of {total}")


def _status(ok, required=True):
    """Return a status word for a comparison row."""
    if not required:
        return "—"
    return "On target" if ok else "Missed"


def build_report_rows(outcome_rows, loads_completed, total_shorts, goal_met, shift_goal):
    """
    Build the expectations-vs-actual comparison rows.
    Each row: area, expected, actual, status.
    """
    oc = [o for o in outcome_rows if o.get("type") == "OC"]
    cpu = [o for o in outcome_rows if o.get("type") == "CPU"]

    oc_signoff_req = sum(1 for o in oc if o.get("signoff_done") in ("Y", "N"))
    oc_signoff_met = sum(1 for o in oc if o.get("signoff_done") == "Y")
    oc_photos_req = sum(1 for o in oc if o.get("photos_done") in ("Y", "N"))
    oc_photos_met = sum(1 for o in oc if o.get("photos_done") == "Y")
    oc_total = len(oc)
    oc_on_time = sum(1 for o in oc if o.get("on_time") == "Y")
    cpu_total = len(cpu)
    cpu_on_time = sum(1 for o in cpu if o.get("on_time") == "Y")

    goal_norm = _norm_na(goal_met)
    goal_actual = {"Y": "Met", "N": "Not met"}.get(goal_norm, "Not recorded")
    goal_status = "On target" if goal_norm == "Y" else ("Missed" if goal_norm == "N" else "—")

    rows = [
        {
            "area": "Shift Goal",
            "expected": shift_goal or "Not recorded",
            "actual": goal_actual,
            "status": goal_status,
        },
        {
            "area": "OC Sign-Off",
            "expected": f"{oc_signoff_req} required" if oc_signoff_req else "None required",
            "actual": f"{oc_signoff_met} collected",
            "status": _status(oc_signoff_met >= oc_signoff_req, required=oc_signoff_req > 0),
        },
        {
            "area": "OC Photos",
            "expected": f"{oc_photos_req} required" if oc_photos_req else "None required",
            "actual": f"{oc_photos_met} taken",
            "status": _status(oc_photos_met >= oc_photos_req, required=oc_photos_req > 0),
        },
        {
            "area": "OC On-Time",
            "expected": f"{oc_total} load(s)" if oc_total else "No OC loads",
            "actual": f"{oc_on_time} on time",
            "status": _status(oc_on_time >= oc_total, required=oc_total > 0),
        },
        {
            "area": "CPU On-Time",
            "expected": f"{cpu_total} appt(s)" if cpu_total else "No CPUs",
            "actual": f"{cpu_on_time} on time",
            "status": _status(cpu_on_time >= cpu_total, required=cpu_total > 0),
        },
        {
            "area": "Shorts",
            "expected": "Target 0",
            "actual": f"{int(total_shorts)}",
            "status": _status(int(total_shorts) == 0),
        },
        {
            "area": "Loads Completed",
            "expected": "—",
            "actual": f"{int(loads_completed)}",
            "status": "—",
        },
    ]

    misses = []
    for o in outcome_rows:
        if (
            str(o.get("shipped")).upper() == "N"
            or str(o.get("on_time")).upper() == "N"
            or str(o.get("signoff_done")).upper() == "N"
            or str(o.get("photos_done")).upper() == "N"
            or str(o.get("short")).upper() in ("Y", "YES")
        ):
            misses.append(o)

    return rows, misses

def build_report_pdf(operating_date, shift, report_rows, misses, notes):
    """Build the one-page End-of-Shift report PDF. Returns bytes, or None."""
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
        Paragraph("End-of-Shift Report", title_style),
        Paragraph(f"{operating_date} &nbsp;|&nbsp; {shift} shift &nbsp;|&nbsp; Expectations vs Actual", sub_style),
    ]

    # Comparison table
    data = [["Area", "Expected", "Actual", "Result"]]
    for r in report_rows:
        data.append([
            Paragraph(str(r["area"]), body),
            Paragraph(str(r["expected"]), body),
            Paragraph(str(r["actual"]), body),
            Paragraph(str(r["status"]), body),
        ])

    table = Table(data, colWidths=[1.4 * inch, 3.0 * inch, 1.7 * inch, 1.0 * inch], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F5B78")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F4F7F9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    # Color the Result cell per row.
    for i, r in enumerate(report_rows, start=1):
        s = r["status"]
        if s in ("On target", "Met"):
            style.append(("BACKGROUND", (3, i), (3, i), colors.HexColor("#C6EFCE")))
        elif s in ("Missed", "Not met"):
            style.append(("BACKGROUND", (3, i), (3, i), colors.HexColor("#FFC7CE")))
        else:
            style.append(("BACKGROUND", (3, i), (3, i), colors.HexColor("#ECECEC")))
    table.setStyle(TableStyle(style))
    story.append(table)

    # Misses
    story.append(Paragraph("Misses this shift", h_style))
    if misses:
        miss_data = [["Type", "Load", "Customer", "Appt", "Reason"]]
        for m in misses:
            miss_data.append([
                Paragraph(str(m.get("type", "")), body),
                Paragraph(str(m.get("load", "")), body),
                Paragraph(str(m.get("customer", "")), body),
                Paragraph(str(m.get("appt_time", "")), body),
                Paragraph(str(m.get("miss_reason", "") or "—"), body),
            ])
        miss_table = Table(miss_data, colWidths=[0.7 * inch, 0.9 * inch, 2.2 * inch, 0.8 * inch, 2.5 * inch], repeatRows=1)
        miss_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#C00000")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8.5),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(miss_table)
    else:
        story.append(Paragraph("No misses recorded this shift.", body))

    # Notes
    story.append(Paragraph("Operational notes", h_style))
    story.append(Paragraph(str(notes).strip() or "—", body))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def render_report_table(report_rows):
    """On-screen version of the comparison table."""
    df = pd.DataFrame([
        {"Area": r["area"], "Expected": r["expected"], "Actual": r["actual"], "Result": r["status"]}
        for r in report_rows
    ])
    st.table(df)


# ============================================================
#  PAGE
# ============================================================

st.set_page_config(page_title="Shift Closeout", layout="wide")
st.title("Shift Closeout")
st.write(
    "Confirm how today's commitments closed out. This records the proof of how we "
    "did against the shift goal, OC, and CPU expectations."
)

if not shift_log.is_configured():
    st.error(
        "The shift log isn't connected yet, so closeouts can't be saved. "
        f"Reason: {shift_log.setup_hint()}"
    )
    st.info(
        "See the setup steps at the top of shift_log.py: create a Google Sheet, "
        "add a service account, and put the credentials + sheet ID in Streamlit secrets."
    )
    st.stop()

col_a, col_b = st.columns(2)
operating_date = col_a.date_input("Operating date", value=datetime.date.today())
shift = col_b.selectbox("Shift", ["1st", "2nd"])
operating_date_str = operating_date.strftime("%m/%d/%Y")

try:
    commitments = shift_log.load_commitments(operating_date_str, shift)
except Exception as e:
    st.error(f"Could not load commitments: {e}")
    st.stop()

if not commitments:
    st.warning(
        f"No commitments were snapshotted for {operating_date_str} {shift} shift. "
        "Run the morning staffing report for this day first — it captures the shift "
        "goal and the OC/CPU commitments that this screen closes out."
    )
    st.stop()

oc_commitments = [
    c for c in commitments
    if str(c.get("type")) == "OC" and _in_shift_window(c.get("appt_time"), shift)
]
cpu_commitments = [
    c for c in commitments
    if str(c.get("type")) == "CPU" and _in_shift_window(c.get("appt_time"), shift)
]
shift_goal = shift_log.get_shift_goal(commitments)

already_closed = shift_log.outcomes_exist(operating_date_str, shift)
if already_closed:
    st.warning(
        "This shift has already been closed out. Submitting again will overwrite "
        "the earlier record."
    )

st.caption(
    f"Loaded {len(oc_commitments)} OC commitment(s) and {len(cpu_commitments)} "
    f"CPU commitment(s) from the morning run."
)

if shift_goal:
    st.info(f"**Shift goal (from this morning):** {shift_goal}")
else:
    st.caption("No shift goal was recorded for this day.")


# ── The closeout form ───────────────────────────────────────────────────────
with st.form("closeout_form"):
    outcome_rows = []

    if oc_commitments:
        st.subheader("Opportunity Customer loads")
        for c in oc_commitments:
            load = str(c.get("load", ""))
            cust = str(c.get("customer", ""))
            appt = str(c.get("appt_time", ""))
            signoff_required = str(c.get("signoff_required")).upper() == "Y"
            photos_required = str(c.get("photos_required")).upper() == "Y"

            with st.expander(f"OC  •  Load {load}  •  {cust}  •  appt {appt}", expanded=True):
                if c.get("requirement"):
                    st.caption(f"Requirement: {c.get('requirement')}")

                row1 = st.columns(2)
                shipped = row1[0].selectbox("Shipped?", YES_NO, key=f"oc_ship_{load}")
                on_time = row1[1].selectbox(
                    "On time vs appointment?", YES_NO, key=f"oc_ot_{load}",
                    help="Did it leave by its appointment time?",
                )

                row2 = st.columns(2)
                if signoff_required:
                    signoff_done = row2[0].selectbox(
                        "Supervisor sign-off collected?", YES_NO, key=f"oc_sign_{load}"
                    )
                else:
                    signoff_done = "NA"
                    row2[0].caption("Sign-off not required.")

                if photos_required:
                    photos_done = row2[1].selectbox(
                        "Photos taken (3 dock + 3 loading)?", YES_NO, key=f"oc_pho_{load}"
                    )
                else:
                    photos_done = "NA"
                    row2[1].caption("Photos not required.")

                row3 = st.columns(2)
                short = row3[0].selectbox("Loaded short?", YES_NO, index=1, key=f"oc_short_{load}")
                miss_reason_choice = row3[1].selectbox(
                    "Miss reason (if something went wrong)",
                    MISS_REASONS, index=0, key=f"oc_miss_{load}",
                )
                miss_reason_other = ""
                if miss_reason_choice == "Other (explain)":
                    miss_reason_other = row3[1].text_input("Describe", key=f"oc_miss_other_{load}")
                miss_reason = _resolve_miss_reason(miss_reason_choice, miss_reason_other)

                outcome_rows.append({
                    "type": "OC", "load": load, "customer": cust, "appt_time": appt,
                    "shipped": shipped,
                    "on_time": _yn_or_na(shipped, on_time),
                    "signoff_done": _norm_na(signoff_done),
                    "photos_done": _norm_na(photos_done),
                    "short": short, "miss_reason": miss_reason,
                })

    if cpu_commitments:
        st.subheader("CPU appointments")
        for c in cpu_commitments:
            load = str(c.get("load", ""))
            cust = str(c.get("customer", ""))
            appt = str(c.get("appt_time", ""))

            with st.expander(f"CPU  •  Load {load}  •  {cust}  •  appt {appt}", expanded=True):
                row1 = st.columns(2)
                shipped = row1[0].selectbox("Shipped?", YES_NO, key=f"cpu_ship_{load}")
                on_time = row1[1].selectbox("Customer left on time?", YES_NO, key=f"cpu_ot_{load}")
                row2 = st.columns(2)
                short = row2[0].selectbox("Loaded short?", YES_NO, index=1, key=f"cpu_short_{load}")
                miss_reason_choice = row2[1].selectbox(
                    "Miss reason", MISS_REASONS, index=0, key=f"cpu_miss_{load}",
                )
                miss_reason_other = ""
                if miss_reason_choice == "Other (explain)":
                    miss_reason_other = row2[1].text_input("Describe", key=f"cpu_miss_other_{load}")
                miss_reason = _resolve_miss_reason(miss_reason_choice, miss_reason_other)

                outcome_rows.append({
                    "type": "CPU", "load": load, "customer": cust, "appt_time": appt,
                    "shipped": shipped,
                    "on_time": _yn_or_na(shipped, on_time),
                    "signoff_done": "NA", "photos_done": "NA",
                    "short": short, "miss_reason": miss_reason,
                })

    # ----- Shift goal result -----
    st.subheader("Shift goal")
    if shift_goal:
        st.caption(shift_goal)
        goal_met = st.selectbox("Did we meet the shift goal?", YES_NO, key="goal_met")
    else:
        goal_met = "NA"
        st.caption("No shift goal was recorded, so there's nothing to mark here.")

    # ----- Shift totals (OT removed) -----
    st.subheader("Shift totals")
    s1, s2 = st.columns(2)
    loads_completed = s1.number_input("Loads completed this shift", min_value=0, step=1, value=0)
    total_shorts = s2.number_input(
        "Loads shipped short this shift",
        min_value=0, step=1, value=0,
        help="Total number of loads that shipped short across the whole shift, "
             "including any OC/CPU loads you already marked short above. Count loads, not cases.",
    )
    notes = st.text_area("Operational notes")

    submitted = st.form_submit_button("Save Closeout & Build Report")


# ── Handle submission ───────────────────────────────────────────────────────
if submitted:
    shorts_marked_above = sum(
        1 for o in outcome_rows if str(o.get("short")).strip().upper() in ("Y", "YES")
    )
    if int(total_shorts) < shorts_marked_above:
        st.error(
            f"You marked {shorts_marked_above} load(s) short above, but entered "
            f"{int(total_shorts)} for total shorts. The shift total can't be less than "
            f"the loads you already marked short. Fix the total (or the per-load answers) "
            f"and submit again."
        )
        st.stop()
    summary = _build_summary(
        outcome_rows, loads_completed, total_shorts, goal_met, shift_goal, notes
    )
    report_rows, misses = build_report_rows(
        outcome_rows, loads_completed, total_shorts, goal_met, shift_goal
    )
    try:
        result = shift_log.save_outcomes(operating_date_str, shift, outcome_rows, summary)
        pdf_bytes = build_report_pdf(operating_date_str, shift, report_rows, misses, notes)
        st.session_state["closeout_report"] = {
            "date": operating_date_str,
            "shift": shift,
            "rows": report_rows,
            "misses": misses,
            "notes": notes,
            "pdf": pdf_bytes,
        }
        st.success(
            f"Closeout saved — {result['outcomes_written']} commitment outcome(s) "
            f"recorded for {operating_date_str} {shift} shift."
        )
    except Exception as e:
        st.error(f"Could not save closeout: {e}")


# ── End-of-Shift report (persists across reruns via session_state) ──────────
report = st.session_state.get("closeout_report")
if report and report["date"] == operating_date_str and report["shift"] == shift:
    st.markdown("---")
    st.subheader("End-of-Shift Report — Expectations vs Actual")
    render_report_table(report["rows"])

    if report["misses"]:
        st.markdown("**Misses this shift**")
        st.dataframe(
            pd.DataFrame([
                {
                    "Type": m.get("type"), "Load": m.get("load"),
                    "Customer": m.get("customer"), "Appt": m.get("appt_time"),
                    "Reason": m.get("miss_reason") or "—",
                }
                for m in report["misses"]
            ]),
            use_container_width=True,
        )

    if report.get("pdf"):
        st.download_button(
            "Download End-of-Shift Report (PDF)",
            data=report["pdf"],
            file_name=f"End_of_Shift_{operating_date_str.replace('/', '-')}_{shift}.pdf",
            mime="application/pdf",
        )
    elif not REPORTLAB_AVAILABLE:
        st.caption("PDF download unavailable — add reportlab to requirements.txt to enable it.")


# ── Rolling scorecard ───────────────────────────────────────────────────────
st.markdown("---")
st.subheader("Rolling 30-Day Scorecard")

try:
    score = shift_log.get_recent_scorecard(days=30)
except Exception as e:
    st.error(f"Could not load scorecard: {e}")
    score = None

if score:
    st.caption(f"Based on {score['shifts_logged']} shift(s) closed out in the last 30 days.")
    m1, m2, m3, m4 = st.columns(4)
    _metric(m1, "OC Sign-Off", score["oc_signoff"], "met", "required")
    _metric(m2, "OC Photos", score["oc_photos"], "met", "required")
    _metric(m3, "CPU On-Time", score["cpu_on_time"], "met", "total")
    _metric(m4, "Shift Goal Met", score["shift_goal"], "met", "total")

    if score["misses"]:
        st.markdown("**Itemized misses (last 30 days)**")
        st.dataframe(
            pd.DataFrame([
                {
                    "Date": r.get("date"), "Shift": r.get("shift"),
                    "Type": r.get("type"), "Load": r.get("load"),
                    "Customer": r.get("customer"), "Appt": r.get("appt_time"),
                    "On time": r.get("on_time"), "Sign-off": r.get("signoff_done"),
                    "Photos": r.get("photos_done"), "Short": r.get("short"),
                    "Reason": r.get("miss_reason"),
                }
                for r in score["misses"]
            ]),
            use_container_width=True, height=300,
        )
    else:
        st.success("No misses recorded in the last 30 days.")
