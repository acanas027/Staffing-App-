"""
3_Shift_Closeout.py
===================
End-of-shift screen. The supervisor opens this when the shift is wrapping up.

It loads the commitments + shift goal snapshotted during the morning run, asks the
supervisor to confirm what actually happened, writes the result to the persistent
log, and produces a one-page End-of-Shift report comparing expectations vs actual.

The supervisor uploads the OpenDock export. The app auto-scores service time for
all outbound loads in the uploaded file, then asks only for the human confirmations
still needed: OC shorts/reasons, shift goal, totals, and notes.
"""

import datetime
import io
import re

import pandas as pd
import streamlit as st

import shift_log
import dc_config

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
SERVICE_TARGET_MINUTES = 120

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
    Window boundaries come from dc_config (1st 06:00-16:30, 2nd 17:00-05:00),
    so they stay in sync with the staffing report. Loads with no parseable appt
    time are kept (can't confidently exclude them).
    """
    return dc_config.in_shift_window(_appt_minutes(appt_time), str(shift).strip())


def _fmt_minutes(mins):
    """Minutes since midnight -> 'HH:MM'."""
    mins = int(mins) % (24 * 60)
    return f"{mins // 60:02d}:{mins % 60:02d}"


# ============================================================
#  OPENDOCK HELPERS
# ============================================================

def _clean_text(value):
    """Return a safe display string for uploaded spreadsheet values."""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    return str(value).strip()


def _norm_header(value):
    """Normalize spreadsheet headers so minor spacing/punctuation changes do not break the upload."""
    return re.sub(r"[^a-z0-9]+", "", str(value or "").strip().lower())


def _find_col(df, candidates):
    """Find the first matching column from a list of possible OpenDock column names."""
    lookup = {_norm_header(c): c for c in df.columns}
    for candidate in candidates:
        key = _norm_header(candidate)
        if key in lookup:
            return lookup[key]
    return None


def _norm_load_id(value):
    """
    Normalize load references for matching.
    Examples: 173760 -> 173760, LD174921 -> 174921.
    """
    text = _clean_text(value).upper()
    if not text:
        return ""
    digits = re.findall(r"\d+", text)
    return digits[-1] if digits else text


def _fmt_opendock_time(value):
    """Normalize OpenDock appointment times to HH:MM when possible."""
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    if isinstance(value, datetime.datetime):
        return value.strftime("%H:%M")
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M")
    if isinstance(value, (int, float)) and 0 <= float(value) < 1:
        mins = int(round(float(value) * 24 * 60))
        return _fmt_minutes(mins)
    text = str(value).strip()
    m = re.match(r"^(\d{1,2}):(\d{2})", text)
    if m:
        return f"{int(m.group(1)):02d}:{m.group(2)}"
    return text


def _opendock_blank_datetime(value):
    """
    Detect blank date/time cells from OpenDock.
    Some exports/templates show placeholder 45 in blank date/time fields, so date/time
    checks treat 45 as blank. Service Time does NOT use this helper because 45 minutes
    is a valid service time.
    """
    text = _clean_text(value)
    return text == "" or text.lower() in ("nan", "nat", "none", "null") or text == "45"


def _service_minutes(value):
    """Convert OpenDock Service Time to an integer number of minutes."""
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    num = pd.to_numeric(value, errors="coerce")
    if pd.isna(num):
        return None
    return int(round(float(num)))


def _score_opendock_service(status, arrival_date, arrival_time, departure_date, departure_time, service_time):
    """Apply the end-of-shift service-time rules to one OpenDock row."""
    status_text = _clean_text(status)
    status_key = status_text.lower().replace(" ", "")
    arrived = not (_opendock_blank_datetime(arrival_date) and _opendock_blank_datetime(arrival_time))
    departed = not (_opendock_blank_datetime(departure_date) and _opendock_blank_datetime(departure_time))
    service_min = _service_minutes(service_time)

    if status_key in ("scheduled", "noshow", "no-show"):
        return {
            "result_type": "no_show",
            "service_minutes": service_min,
            "delay_minutes": None,
            "service_result": "Customer no-show",
        }

    if status_key == "cancelled":
        return {
            "result_type": "cancelled",
            "service_minutes": service_min,
            "delay_minutes": None,
            "service_result": "Cancelled appointment",
        }

    if arrived and not departed:
        return {
            "result_type": "no_departure",
            "service_minutes": service_min,
            "delay_minutes": None,
            "service_result": "No departure recorded",
        }

    if service_min is None:
        return {
            "result_type": "missing_service_time",
            "service_minutes": None,
            "delay_minutes": None,
            "service_result": "Service time not recorded",
        }

    if service_min <= SERVICE_TARGET_MINUTES:
        return {
            "result_type": "target_met",
            "service_minutes": service_min,
            "delay_minutes": 0,
            "service_result": "Service time target met",
        }

    delay = service_min - SERVICE_TARGET_MINUTES
    return {
        "result_type": "delayed",
        "service_minutes": service_min,
        "delay_minutes": delay,
        "service_result": f"Delayed service (over {SERVICE_TARGET_MINUTES})",
    }


def build_opendock_service_audit(uploaded_file, operating_date, shift):
    """
    Read the uploaded OpenDock Excel export and build service-time audit rows for all
    outbound loads on the selected operating date and shift.
    Returns: (service_rows, service_by_load)
    """
    if uploaded_file is None:
        return [], {}

    uploaded_file.seek(0)
    df = pd.read_excel(uploaded_file)
    df.columns = [str(c).strip() for c in df.columns]

    cols = {
        "appt_date": _find_col(df, ["Appt Date", "Appointment Date"]),
        "appt_time": _find_col(df, ["Appt Time", "Appointment Time"]),
        "status": _find_col(df, ["Status"]),
        "arrival_date": _find_col(df, ["Arrival Date"]),
        "arrival_time": _find_col(df, ["Arrival Time"]),
        "departure_date": _find_col(df, ["Departure Date"]),
        "departure_time": _find_col(df, ["Departure Time"]),
        "service_time": _find_col(df, ["Service Time", "Service Time (mins)", "Service Time mins"]),
        "load_reference": _find_col(df, ["Load Reference", "Load Ref", "Load"]),
        "carrier": _find_col(df, ["Carrier Company", "Carrier", "Customer"]),
        "load_type": _find_col(df, ["Load Type"]),
        "direction": _find_col(df, ["Direction"]),
        "dock": _find_col(df, ["Dock"]),
    }

    required = ["appt_date", "appt_time", "status", "arrival_date", "arrival_time",
                "departure_date", "departure_time", "service_time", "load_reference"]
    missing = [name for name in required if not cols.get(name)]
    if missing:
        raise ValueError(
            "OpenDock upload is missing required column(s): " + ", ".join(missing)
        )

    work = df.copy()
    work["_appt_date_parsed"] = pd.to_datetime(work[cols["appt_date"]], errors="coerce").dt.date
    work["_appt_time_text"] = work[cols["appt_time"]].apply(_fmt_opendock_time)
    work["_load_norm"] = work[cols["load_reference"]].apply(_norm_load_id)

    # End-of-shift report is outbound-focused. Remove this filter if you later want inbound audited too.
    if cols.get("direction"):
        work = work[work[cols["direction"]].astype(str).str.strip().str.upper().eq("OUTBOUND")]

    work = work[work["_appt_date_parsed"].eq(operating_date)]
    work = work[work["_appt_time_text"].apply(lambda x: _in_shift_window(x, shift))]

    service_rows = []
    for _, row in work.iterrows():
        score = _score_opendock_service(
            row.get(cols["status"]),
            row.get(cols["arrival_date"]), row.get(cols["arrival_time"]),
            row.get(cols["departure_date"]), row.get(cols["departure_time"]),
            row.get(cols["service_time"]),
        )
        load = _clean_text(row.get(cols["load_reference"]))
        appt = _fmt_opendock_time(row.get(cols["appt_time"]))
        service_rows.append({
            "load": load,
            "load_norm": _norm_load_id(load),
            "customer": _clean_text(row.get(cols["carrier"])) if cols.get("carrier") else "",
            "appt_date": operating_date.strftime("%m/%d/%Y"),
            "appt_time": appt,
            "status": _clean_text(row.get(cols["status"])),
            "service_minutes": score["service_minutes"],
            "delay_minutes": score["delay_minutes"],
            "service_result": score["service_result"],
            "result_type": score["result_type"],
            "load_type": _clean_text(row.get(cols["load_type"])) if cols.get("load_type") else "",
            "dock": _clean_text(row.get(cols["dock"])) if cols.get("dock") else "",
        })

    service_by_load = {}
    for r in service_rows:
        exact_key = _clean_text(r.get("load", "")).upper()
        norm_key = r.get("load_norm")
        if exact_key:
            service_by_load[exact_key] = r
        if norm_key and norm_key not in service_by_load:
            # Only use the digit-only fallback when there is not an exact text match.
            # This prevents LD173153 from overwriting a different 173153 appointment.
            service_by_load[norm_key] = r

    return service_rows, service_by_load


def _opendock_counts(service_rows):
    """Summarize OpenDock service audit rows for report scoring."""
    counts = {
        "total": len(service_rows),
        "target_met": 0,
        "delayed": 0,
        "no_show": 0,
        "no_departure": 0,
        "missing_service_time": 0,
        "cancelled": 0,
    }
    for r in service_rows:
        rt = r.get("result_type")
        if rt in counts:
            counts[rt] += 1
    counts["issues"] = (
        counts["delayed"] + counts["no_show"] +
        counts["no_departure"] + counts["missing_service_time"]
    )
    return counts


def _commitment_auto_fields(commitment, service_by_load):
    """
    Convert OpenDock service result into the existing outcome fields used by shift_log.
    This lets the old scorecard keep working while the supervisor no longer manually
    answers shipped/on-time for every load.
    """
    raw_key = _clean_text(commitment.get("load", "")).upper()
    norm_key = _norm_load_id(commitment.get("load", ""))
    service = service_by_load.get(raw_key) or service_by_load.get(norm_key)
    if not service:
        return {
            "shipped": "No",
            "on_time": "NA",
            "auto_reason": "Not found in OpenDock upload",
        }

    rt = service.get("result_type")
    if rt == "target_met":
        return {"shipped": "Yes", "on_time": "Y", "auto_reason": ""}
    if rt == "delayed":
        return {"shipped": "Yes", "on_time": "N", "auto_reason": service.get("service_result", "")}
    if rt in ("no_show", "no_departure", "missing_service_time", "cancelled"):
        return {"shipped": "No", "on_time": "NA", "auto_reason": service.get("service_result", "")}
    return {"shipped": "No", "on_time": "NA", "auto_reason": service.get("service_result", "Review OpenDock status")}


def _combine_reasons(*parts):
    """Join non-empty reason strings without losing either manual or OpenDock context."""
    clean = [str(p).strip() for p in parts if str(p or "").strip()]
    return " | ".join(dict.fromkeys(clean))


def _goal_predicted_cutoff(shift_goal):
    """
    Pull the predicted appointment cutoff out of the morning goal text, which reads
    '...through appointment 14:00 (...) by shift end 16:30.' We want the appointment
    cutoff, never the shift-end time, so we only match the time right after the word
    'appointment'. Returns minutes since midnight, or None if there's no real cutoff
    (e.g. the goal said 'none - before first appt').
    """
    import re
    m = re.search(r"appointment\s+(\d{1,2}):(\d{2})", str(shift_goal or ""), flags=re.IGNORECASE)
    if not m:
        return None
    return int(m.group(1)) * 60 + int(m.group(2))


def _cutoff_variance(shift_goal, actual_cutoff):
    """
    Compare predicted cutoff (from the goal) to the actual cutoff controlled to.
    Later actual = controlled further into the day = AHEAD of plan; earlier = BEHIND.
    Returns a dict (predicted/actual/delta/direction/message) or None if either time
    is missing or unparseable.
    """
    predicted = _goal_predicted_cutoff(shift_goal)
    actual = _appt_minutes(actual_cutoff)
    if predicted is None or actual is None:
        return None
    delta = actual - predicted  # + later = ahead, - earlier = behind
    if delta > 0:
        direction = "AHEAD"
        message = (f"Controlled to {actual_cutoff} vs predicted {_fmt_minutes(predicted)} - "
                   f"{delta} min ahead of plan.")
    elif delta < 0:
        direction = "BEHIND"
        message = (f"Controlled to {actual_cutoff} vs predicted {_fmt_minutes(predicted)} - "
                   f"{abs(delta)} min behind plan.")
    else:
        direction = "ON TARGET"
        message = f"Controlled to {actual_cutoff}, exactly the predicted cutoff."
    return {"predicted_min": predicted, "actual_min": actual,
            "delta_min": delta, "direction": direction, "message": message}


def _build_summary(outcome_rows, loads_completed, total_shorts, goal_met, shift_goal, notes,
                   actual_cutoff=""):
    """Roll per-commitment outcomes into the one-row shift summary."""
    oc = [o for o in outcome_rows if o.get("type") == "OC"]
    cpu = [o for o in outcome_rows if o.get("type") == "CPU"]

    oc_service_met = sum(1 for o in oc if o.get("on_time") == "Y")

    return {
        "loads_completed": loads_completed,
        "total_shorts": total_shorts,
        "goal_met": _norm_na(goal_met),
        "shift_goal": shift_goal,
        "actual_cutoff": "",
        "oc_total": len(oc),
        "oc_service_target_met": oc_service_met,
        "oc_service_target_total": len(oc),
        # Keep the old keys so existing shift_log / Google Sheet setups do not break.
        # They are no longer used as supervisor KPIs on this page.
        "oc_signoff_met": 0,
        "oc_photos_met": 0,
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


def build_report_rows(outcome_rows, loads_completed, total_shorts, goal_met, shift_goal, service_rows=None):
    """
    Build the expectations-vs-actual comparison rows.
    Each row: area, expected, actual, status.
    """
    service_rows = service_rows or []
    oc = [o for o in outcome_rows if o.get("type") == "OC"]
    cpu = [o for o in outcome_rows if o.get("type") == "CPU"]

    oc_total = len(oc)
    oc_on_time = sum(1 for o in oc if o.get("on_time") == "Y")
    oc_shorts = sum(1 for o in oc if str(o.get("short")).strip().upper() in ("Y", "YES"))
    cpu_total = len(cpu)
    cpu_on_time = sum(1 for o in cpu if o.get("on_time") == "Y")

    service_counts = _opendock_counts(service_rows)

    goal_norm = _norm_na(goal_met)
    goal_actual = {"Y": "Met", "N": "Not met"}.get(goal_norm, "Not recorded")
    goal_status = "On target" if goal_norm == "Y" else ("Missed" if goal_norm == "N" else "—")

    if service_counts["total"]:
        service_actual = (
            f"{service_counts['target_met']} met / {service_counts['total']} load(s); "
            f"{service_counts['delayed']} delayed; "
            f"{service_counts['no_show']} no-show; "
            f"{service_counts['no_departure']} no departure"
        )
        service_status = "On target" if service_counts["issues"] == 0 else "Missed"
    else:
        service_actual = "No OpenDock rows found for this date/shift"
        service_status = "—"

    rows = [
        {
            "area": "Shift Goal",
            "expected": shift_goal or "Not recorded",
            "actual": goal_actual,
            "status": goal_status,
        },
        {
            "area": "OpenDock Service Time",
            "expected": f"Target <= {SERVICE_TARGET_MINUTES} min for all outbound loads",
            "actual": service_actual,
            "status": service_status,
        },
        {
            "area": "OC Service Target",
            "expected": f"{oc_total} OC load(s) at <= {SERVICE_TARGET_MINUTES} min" if oc_total else "No OC loads",
            "actual": f"{oc_on_time} met service target",
            "status": _status(oc_on_time >= oc_total, required=oc_total > 0),
        },
        {
            "area": "OC Shorts",
            "expected": f"0 short across {oc_total} OC load(s)" if oc_total else "No OC loads",
            "actual": f"{oc_shorts} short",
            "status": _status(oc_shorts == 0, required=oc_total > 0),
        },
        {
            "area": "CPU Service Target",
            "expected": f"{cpu_total} CPU appointment(s) auto-scored from OpenDock" if cpu_total else "No CPUs",
            "actual": f"{cpu_on_time} met service target",
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


def build_report_pdf(operating_date, shift, report_rows, misses, notes, service_rows=None):
    """Build the End-of-Shift report PDF. Returns bytes, or None."""
    service_rows = service_rows or []
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

    # OpenDock service audit
    story.append(Paragraph("OpenDock service-time audit", h_style))
    if service_rows:
        service_data = [["Load", "Customer/Carrier", "Appt", "Status", "Svc Min", "Result"]]
        for r in service_rows:
            service_min = "—" if r.get("service_minutes") is None else str(r.get("service_minutes"))
            service_data.append([
                Paragraph(str(r.get("load", "")), body),
                Paragraph(str(r.get("customer", "")), body),
                Paragraph(str(r.get("appt_time", "")), body),
                Paragraph(str(r.get("status", "")), body),
                Paragraph(service_min, body),
                Paragraph(str(r.get("service_result", "")), body),
            ])
        service_table = Table(
            service_data,
            colWidths=[0.8 * inch, 1.7 * inch, 0.65 * inch, 0.8 * inch, 0.6 * inch, 2.55 * inch],
            repeatRows=1,
        )
        service_style = [
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#0F5B78")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 8),
            ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#B7B7B7")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ]
        for i, r in enumerate(service_rows, start=1):
            if r.get("result_type") == "target_met":
                service_style.append(("BACKGROUND", (5, i), (5, i), colors.HexColor("#C6EFCE")))
            elif r.get("result_type") in ("delayed", "no_show", "no_departure", "missing_service_time"):
                service_style.append(("BACKGROUND", (5, i), (5, i), colors.HexColor("#FFC7CE")))
            else:
                service_style.append(("BACKGROUND", (5, i), (5, i), colors.HexColor("#ECECEC")))
        service_table.setStyle(TableStyle(service_style))
        story.append(service_table)
    else:
        story.append(Paragraph("No OpenDock service rows were included.", body))

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
    "Confirm how today's commitments closed out. OpenDock auto-scores service time "
    "for all loads, while the supervisor confirms OC shorts and the shift goal."
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


# ── OpenDock upload + automatic service audit ───────────────────────────────
st.subheader("OpenDock service audit")
opendock_file = st.file_uploader(
    "Upload today's OpenDock appointment export",
    type=["xlsx", "xls"],
    key="opendock_file",
    help="The report will automatically score every outbound load in this file against the 120-minute service target.",
)

opendock_service_rows = []
opendock_by_load = {}
opendock_upload_error = ""
if opendock_file is not None:
    try:
        opendock_service_rows, opendock_by_load = build_opendock_service_audit(
            opendock_file, operating_date, shift
        )
        st.caption(
            f"OpenDock loaded: {len(opendock_service_rows)} outbound load(s) found for "
            f"{operating_date_str} {shift} shift. Details will appear in the PDF report."
        )
    except Exception as e:
        opendock_upload_error = str(e)
        st.error(f"Could not read OpenDock upload: {e}")
else:
    st.info("Upload OpenDock before saving closeout so service time is audited automatically.")


# ── The closeout form ───────────────────────────────────────────────────────
with st.form("closeout_form"):
    outcome_rows = []

    if oc_commitments:
        st.subheader("Opportunity Customer loads — short confirmation")
        st.caption(
            "OpenDock now auto-scores service time in the background. For OC loads, "
            "the supervisor only confirms whether anything shipped short and why."
        )
        for c in oc_commitments:
            load = str(c.get("load", ""))
            cust = str(c.get("customer", ""))
            appt = str(c.get("appt_time", ""))
            auto = _commitment_auto_fields(c, opendock_by_load)

            with st.expander(f"OC  •  Load {load}  •  {cust}  •  appt {appt}", expanded=True):
                if c.get("requirement"):
                    st.caption(f"Requirement: {c.get('requirement')}")

                row = st.columns(2)
                short = row[0].selectbox("Loaded short?", YES_NO, index=1, key=f"oc_short_{load}")
                miss_reason_choice = row[1].selectbox(
                    "Short reason / miss reason",
                    MISS_REASONS, index=0, key=f"oc_miss_{load}",
                )
                miss_reason_other = ""
                if miss_reason_choice == "Other (explain)":
                    miss_reason_other = row[1].text_input("Describe", key=f"oc_miss_other_{load}")
                manual_reason = _resolve_miss_reason(miss_reason_choice, miss_reason_other)
                miss_reason = _combine_reasons(manual_reason, auto.get("auto_reason"))

                outcome_rows.append({
                    "type": "OC", "load": load, "customer": cust, "appt_time": appt,
                    "shipped": auto.get("shipped", "No"),
                    "on_time": auto.get("on_time", "NA"),
                    "signoff_done": "NA",
                    "photos_done": "NA",
                    "short": short,
                    "miss_reason": miss_reason,
                })

    # CPU outcomes are auto-scored from OpenDock and saved in the background.
    # Nothing is shown to the supervisor because there is no manual CPU input needed.
    if cpu_commitments:
        for c in cpu_commitments:
            load = str(c.get("load", ""))
            cust = str(c.get("customer", ""))
            appt = str(c.get("appt_time", ""))
            auto = _commitment_auto_fields(c, opendock_by_load)
            outcome_rows.append({
                "type": "CPU", "load": load, "customer": cust, "appt_time": appt,
                "shipped": auto.get("shipped", "No"),
                "on_time": auto.get("on_time", "NA"),
                "signoff_done": "NA", "photos_done": "NA",
                "short": "No",
                "miss_reason": auto.get("auto_reason", ""),
            })

    # ----- Shift goal result -----
    st.subheader("Shift goal")
    if shift_goal:
        st.caption(shift_goal)
        goal_met = st.selectbox("Did we meet the shift goal?", YES_NO, key="goal_met")
    else:
        goal_met = "NA"
        st.caption("No shift goal was recorded, so there's nothing to mark here.")

    # ----- Shift totals -----
    st.subheader("Shift totals")
    s1, s2 = st.columns(2)
    loads_completed = s1.number_input("Loads completed this shift", min_value=0, step=1, value=0)
    total_shorts = s2.number_input(
        "Loads shipped short this shift",
        min_value=0, step=1, value=0,
        help="Total number of loads that shipped short across the whole shift, "
             "including any OC loads you already marked short above. Count loads, not cases.",
    )
    notes = st.text_area("Operational notes")

    submitted = st.form_submit_button("Save Closeout & Build Report")


# ── Handle submission ───────────────────────────────────────────────────────
if submitted:
    if opendock_file is None:
        st.error("Upload the OpenDock appointment export before saving closeout.")
        st.stop()
    if opendock_upload_error:
        st.error(f"Fix the OpenDock upload before saving closeout: {opendock_upload_error}")
        st.stop()
    if not opendock_service_rows:
        st.error(
            "OpenDock loaded, but no outbound loads were found for this operating date and shift. "
            "Check the Operating date, Shift, and the Appt Date/Appt Time columns in the upload."
        )
        st.stop()

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
        outcome_rows, loads_completed, total_shorts, goal_met, shift_goal, notes,
    )
    report_rows, misses = build_report_rows(
        outcome_rows, loads_completed, total_shorts, goal_met, shift_goal,
        service_rows=opendock_service_rows,
    )
    try:
        result = shift_log.save_outcomes(operating_date_str, shift, outcome_rows, summary)
        pdf_bytes = build_report_pdf(
            operating_date_str, shift, report_rows, misses, notes,
            service_rows=opendock_service_rows,
        )
        st.session_state["closeout_report"] = {
            "date": operating_date_str,
            "shift": shift,
            "rows": report_rows,
            "misses": misses,
            "notes": notes,
            "service_rows": opendock_service_rows,
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

    if report.get("service_rows"):
        with st.expander("OpenDock service-time audit included in report", expanded=False):
            st.dataframe(
                pd.DataFrame([
                    {
                        "Load": r.get("load"),
                        "Customer/Carrier": r.get("customer"),
                        "Appt": r.get("appt_time"),
                        "Status": r.get("status"),
                        "Service Min": r.get("service_minutes"),
                        "Result": r.get("service_result"),
                    }
                    for r in report.get("service_rows", [])
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
    m1, m2, m3 = st.columns(3)
    oc_service_block = score.get("oc_service_target") or score.get("oc_on_time")
    if oc_service_block:
        _metric(m1, "OC Service Target", oc_service_block, "met", "total")
    else:
        m1.metric("OC Service Target", "—")
        m1.caption("Saved from this version forward.")
    _metric(m2, "CPU Service Target", score["cpu_on_time"], "met", "total")
    _metric(m3, "Shift Goal Met", score["shift_goal"], "met", "total")

    if score["misses"]:
        st.markdown("**Itemized misses (last 30 days)**")
        st.dataframe(
            pd.DataFrame([
                {
                    "Date": r.get("date"), "Shift": r.get("shift"),
                    "Type": r.get("type"), "Load": r.get("load"),
                    "Customer": r.get("customer"), "Appt": r.get("appt_time"),
                    "Service target": r.get("on_time"),
                    "Short": r.get("short"),
                    "Reason": r.get("miss_reason"),
                }
                for r in score["misses"]
            ]),
            use_container_width=True, height=300,
        )
    else:
        st.success("No misses recorded in the last 30 days.")
