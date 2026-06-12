"""
shift_log.py
============
Persistent commitment + outcome log for the staffing tool.

WHY THIS EXISTS
---------------
The morning run KNOWS the day's commitments: every Opportunity Customer (OC) load
with its sign-off/photo requirements, every CPU load with its appointment time, and
the AI-generated shift goal. This module persists all of that, then lets the shift
closeout screen record what actually happened. The gap between the two is the proof
your boss wants: OC sign-off %, OC photo %, CPU on-time %, shift-goal-met %, with
every miss itemized.

The tabs (commitments / outcomes / shift_summary) are created automatically.

READ CACHING
------------
All tab reads go through _read_tab_records(), which is cached for 120 seconds. This
keeps repeated Streamlit reruns (every widget click) from hitting the Google Sheets
per-minute read quota. The cache is cleared right after every write so closeouts and
snapshots show fresh data immediately.
"""

import datetime

import streamlit as st

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except Exception:
    GSPREAD_AVAILABLE = False


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

COMMITMENTS_TAB = "commitments"
OUTCOMES_TAB = "outcomes"
SUMMARY_TAB = "shift_summary"

COMMITMENTS_HEADER = [
    "snapshot_id", "date", "shift", "type", "load", "customer", "appt_time",
    "priority", "requirement", "signoff_required", "photos_required",
    "morning_status", "created_at",
]

OUTCOMES_HEADER = [
    "snapshot_id", "date", "shift", "type", "load", "customer", "appt_time",
    "shipped", "on_time", "signoff_done", "photos_done", "short",
    "miss_reason", "closed_at",
]

# OT hours removed. Shift goal + goal_met + actual_cutoff added.
SUMMARY_HEADER = [
    "snapshot_id", "date", "shift", "loads_completed", "total_shorts",
    "goal_met", "shift_goal", "actual_cutoff",
    "oc_total", "oc_signoff_met", "oc_photos_met",
    "cpu_total", "cpu_on_time", "notes", "closed_at",
]


# ============================================================
#  CONFIG / CONNECTION
# ============================================================

def is_configured():
    """True only when both the library and the required secrets are present."""
    if not GSPREAD_AVAILABLE:
        return False
    try:
        return (
            "gcp_service_account" in st.secrets
            and "shift_log_sheet_id" in st.secrets
        )
    except Exception:
        return False


def setup_hint():
    """A short, human-readable reason the log isn't ready, for the UI."""
    if not GSPREAD_AVAILABLE:
        return "The gspread / google-auth packages aren't installed. Add them to requirements.txt."
    try:
        if "gcp_service_account" not in st.secrets:
            return "Missing [gcp_service_account] block in Streamlit secrets."
        if "shift_log_sheet_id" not in st.secrets:
            return "Missing shift_log_sheet_id in Streamlit secrets."
    except Exception:
        return "Streamlit secrets are not available."
    return "Unknown configuration issue."


@st.cache_resource(show_spinner=False)
def _get_spreadsheet():
    """Authorize once per session and return the open spreadsheet handle."""
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]), scopes=SCOPES
    )
    client = gspread.authorize(creds)
    return client.open_by_key(st.secrets["shift_log_sheet_id"])


def _get_tab(name, header):
    """Return the worksheet, creating it with a header row if missing."""
    sh = _get_spreadsheet()
    try:
        ws = sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=1000, cols=max(20, len(header)))
        ws.append_row(header, value_input_option="USER_ENTERED")
        return ws

    first_row = ws.row_values(1)
    if not first_row:
        ws.append_row(header, value_input_option="USER_ENTERED")
    return ws


@st.cache_data(ttl=120, show_spinner=False)
def _read_tab_records(name, header_tuple):
    """
    Cached read of an entire tab's records. TTL 120s means repeated reruns within
    two minutes serve from memory instead of hitting Google — this is what keeps
    the app under the Sheets per-minute read quota. The cache is cleared on write
    (see save_outcomes / snapshot_commitments) so data is never stale after a save.

    header_tuple is a tuple (not a list) because cache-key arguments must be hashable.
    """
    ws = _get_tab(name, list(header_tuple))
    return ws.get_all_records(expected_headers=list(header_tuple))


def _yn(value):
    """Normalize a boolean-ish value to 'Y' / 'N' for the sheet."""
    return "Y" if str(value).strip().upper() in ("Y", "YES", "TRUE", "1") else "N"


def make_snapshot_id(operating_date, shift):
    """Stable key for a single operating date + shift."""
    return f"{operating_date}_{shift}".replace("/", "-").replace(" ", "")


# ============================================================
#  WRITE: MORNING COMMITMENT SNAPSHOT
# ============================================================

def snapshot_commitments(operating_date, shift, oc_load_matches, cpu_commitments, shift_goal="",  total_loads_for_day=""):
    """
    Persist today's commitments and the shift goal. Idempotent per (date, shift):
    re-running the morning report replaces the prior snapshot.

    oc_load_matches : list of dicts as produced by find_oc_load_matches(), with
        keys load, customer_on_board, oc_name, time, priority, requirements,
        sign_off, pictures, status.
    cpu_commitments : list of dicts with keys load, customer, appt_time, morning_status.
    shift_goal      : the AI-generated shift goal string (stored as a GOAL row).
    """
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    snapshot_id = make_snapshot_id(operating_date, shift)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = []

    # Shift goal stored as a single GOAL row (goal text lives in the requirement column).
    rows.append([
        snapshot_id, operating_date, shift, "GOAL",
        "", "", "", "", str(shift_goal or ""), "", "", str(total_loads_for_day or ""), now,
    ])

    for m in oc_load_matches or []:
        rows.append([
            snapshot_id, operating_date, shift, "OC",
            str(m.get("load", "")),
            str(m.get("customer_on_board") or m.get("oc_name", "")),
            str(m.get("time", "")),
            str(m.get("priority", "")),
            str(m.get("requirements", "")),
            _yn(m.get("sign_off")),
            _yn(m.get("pictures")),
            str(m.get("status", "")),
            now,
        ])
    for c in cpu_commitments or []:
        rows.append([
            snapshot_id, operating_date, shift, "CPU",
            str(c.get("load", "")),
            str(c.get("customer", "")),
            str(c.get("appt_time", "")),
            "", "", "N", "N",
            str(c.get("morning_status", "")),
            now,
        ])

    ws = _get_tab(COMMITMENTS_TAB, COMMITMENTS_HEADER)
    _replace_rows_for_snapshot(ws, COMMITMENTS_HEADER, snapshot_id, rows)

    _read_tab_records.clear()  # invalidate cached reads so the new snapshot shows immediately
    return {
        "snapshot_id": snapshot_id,
        "oc_count": len(oc_load_matches or []),
        "cpu_count": len(cpu_commitments or []),
        "shift_goal": shift_goal or "",
        "total": len(rows),
    }


# ============================================================
#  READ: COMMITMENTS FOR A GIVEN DAY (for the closeout screen)
# ============================================================

def load_commitments(operating_date, shift):
    """Return the list of commitment dicts snapshotted for this date+shift (incl. GOAL row)."""
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    snapshot_id = make_snapshot_id(operating_date, shift)
    records = _read_tab_records(COMMITMENTS_TAB, tuple(COMMITMENTS_HEADER))
    return [r for r in records if str(r.get("snapshot_id")) == snapshot_id]


def get_shift_goal(commitments):
    """Pull the goal text out of a loaded commitment list."""
    for c in commitments or []:
        if str(c.get("type")) == "GOAL":
            return str(c.get("requirement", ""))
    return ""

def get_total_loads_for_day(commitments):
    """Pull the day's total outbound-loads input out of a loaded commitment list.
    Stored on the GOAL row's morning_status column by snapshot_commitments()."""
    for c in commitments or []:
        if str(c.get("type")) == "GOAL":
            value = c.get("morning_status")
            try:
                return int(float(str(value).strip()))
            except (TypeError, ValueError):
                return None
    return None

def outcomes_exist(operating_date, shift):
    """True if this shift has already been closed out (used to warn on re-submit)."""
    if not is_configured():
        return False
    snapshot_id = make_snapshot_id(operating_date, shift)
    records = _read_tab_records(OUTCOMES_TAB, tuple(OUTCOMES_HEADER))
    return any(str(r.get("snapshot_id")) == snapshot_id for r in records)


# ============================================================
#  WRITE: CLOSEOUT OUTCOMES + SHIFT SUMMARY
# ============================================================

def save_outcomes(operating_date, shift, outcome_rows, summary):
    """
    Persist per-commitment outcomes plus a one-row shift summary. Idempotent per
    (date, shift) so a corrected re-submit overwrites rather than duplicates.

    summary : dict with keys loads_completed, total_shorts, goal_met, shift_goal,
        actual_cutoff, oc_total, oc_signoff_met, oc_photos_met, cpu_total,
        cpu_on_time, notes.
    """
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    snapshot_id = make_snapshot_id(operating_date, shift)
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    rows = []
    for o in outcome_rows:
        rows.append([
            snapshot_id, operating_date, shift,
            str(o.get("type", "")),
            str(o.get("load", "")),
            str(o.get("customer", "")),
            str(o.get("appt_time", "")),
            _yn(o.get("shipped")),
            str(o.get("on_time", "")),       # Y / N / NA
            str(o.get("signoff_done", "")),  # Y / N / NA
            str(o.get("photos_done", "")),   # Y / N / NA
            _yn(o.get("short")),
            str(o.get("miss_reason", "")),
            now,
        ])

    ws_out = _get_tab(OUTCOMES_TAB, OUTCOMES_HEADER)
    _replace_rows_for_snapshot(ws_out, OUTCOMES_HEADER, snapshot_id, rows)

    summary_row = [
        snapshot_id, operating_date, shift,
        summary.get("loads_completed", 0),
        summary.get("total_shorts", 0),
        str(summary.get("goal_met", "")),     # Y / N / NA
        str(summary.get("shift_goal", "")),
        str(summary.get("actual_cutoff", "")),
        summary.get("oc_total", 0),
        summary.get("oc_signoff_met", 0),
        summary.get("oc_photos_met", 0),
        summary.get("cpu_total", 0),
        summary.get("cpu_on_time", 0),
        str(summary.get("notes", "")),
        now,
    ]
    ws_sum = _get_tab(SUMMARY_TAB, SUMMARY_HEADER)
    _replace_rows_for_snapshot(ws_sum, SUMMARY_HEADER, snapshot_id, [summary_row])

    _read_tab_records.clear()  # invalidate cached reads so the new data shows immediately
    return {"snapshot_id": snapshot_id, "outcomes_written": len(rows)}


# ============================================================
#  READ: ROLLING SCORECARD (the "walk into the manager's office" view)
# ============================================================

def get_recent_scorecard(days=30):
    """
    Compute compliance rates over the last N days. OC/CPU rates come from the
    outcomes log; shift-goal-met comes from the shift_summary log.
    """
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    records = _read_tab_records(OUTCOMES_TAB, tuple(OUTCOMES_HEADER))

    cutoff = datetime.date.today() - datetime.timedelta(days=days)
    recent = []
    for r in records:
        d = _parse_date(r.get("date"))
        if d is not None and d >= cutoff:
            recent.append(r)

    oc = [r for r in recent if str(r.get("type")) == "OC"]
    cpu = [r for r in recent if str(r.get("type")) == "CPU"]

    def _rate(items, key):
        relevant = [r for r in items if str(r.get(key)).upper() in ("Y", "N")]
        if not relevant:
            return None, 0, 0
        met = sum(1 for r in relevant if str(r.get(key)).upper() == "Y")
        return round(100 * met / len(relevant)), met, len(relevant)

    signoff_rate, signoff_met, signoff_req = _rate(oc, "signoff_done")
    photos_rate, photos_met, photos_req = _rate(oc, "photos_done")
    cpu_rate, cpu_met, cpu_total = _rate(cpu, "on_time")

    # Shift-goal-met from the summary tab.
    goal_rate, goal_met, goal_total = None, 0, 0
    try:
        sum_records = _read_tab_records(SUMMARY_TAB, tuple(SUMMARY_HEADER))
        recent_sum = [
            r for r in sum_records
            if (_parse_date(r.get("date")) is not None and _parse_date(r.get("date")) >= cutoff)
        ]
        for r in recent_sum:
            counts, was_met = _goal_met_pair(r.get("goal_met"))
            if counts:
                goal_total += 1
                if was_met:
                    goal_met += 1
        if goal_total:
            goal_rate = round(100 * goal_met / goal_total)
    except Exception:
        pass

    misses = []
    for r in recent:
        if (
            str(r.get("on_time")).upper() == "N"
            or str(r.get("signoff_done")).upper() == "N"
            or str(r.get("photos_done")).upper() == "N"
            or str(r.get("short")).upper() == "Y"
        ):
            misses.append(r)

    return {
        "days": days,
        "shifts_logged": len({r.get("snapshot_id") for r in recent}),
        "oc_signoff": {"rate": signoff_rate, "met": signoff_met, "required": signoff_req},
        "oc_photos": {"rate": photos_rate, "met": photos_met, "required": photos_req},
        "cpu_on_time": {"rate": cpu_rate, "met": cpu_met, "total": cpu_total},
        "shift_goal": {"rate": goal_rate, "met": goal_met, "total": goal_total},
        "misses": misses,
    }


def get_monthly_scorecard(year, month):
    """
    Cumulative goal performance for one calendar month. Reads the same persistent
    log the closeout writes, so it is always current. Returns headline rates, the
    month totals, a per-shift breakdown, and every miss.
    """
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    def _in_month(record):
        d = _parse_date(record.get("date"))
        return d is not None and d.year == int(year) and d.month == int(month)

    records = _read_tab_records(OUTCOMES_TAB, tuple(OUTCOMES_HEADER))
    month_rows = [r for r in records if _in_month(r)]

    oc = [r for r in month_rows if str(r.get("type")) == "OC"]
    cpu = [r for r in month_rows if str(r.get("type")) == "CPU"]

    def _rate(items, key):
        relevant = [r for r in items if str(r.get(key)).upper() in ("Y", "N")]
        if not relevant:
            return None, 0, 0
        met = sum(1 for r in relevant if str(r.get(key)).upper() == "Y")
        return round(100 * met / len(relevant)), met, len(relevant)

    signoff_rate, signoff_met, signoff_req = _rate(oc, "signoff_done")
    photos_rate, photos_met, photos_req = _rate(oc, "photos_done")
    cpu_rate, cpu_met, cpu_total = _rate(cpu, "on_time")

    # Goal-met + month totals + per-shift breakdown from the summary tab.
    goal_rate, goal_met, goal_total = None, 0, 0
    loads_completed_total = 0
    shorts_total = 0
    per_shift = []
    try:
        sum_records = _read_tab_records(SUMMARY_TAB, tuple(SUMMARY_HEADER))
        month_sum = [r for r in sum_records if _in_month(r)]
        for r in month_sum:
            counts, was_met = _goal_met_pair(r.get("goal_met"))
            if counts:
                goal_total += 1
                if was_met:
                    goal_met += 1
            loads_completed_total += _to_int(r.get("loads_completed"))
            shorts_total += _to_int(r.get("total_shorts"))
            per_shift.append({
                "date": r.get("date"),
                "shift": r.get("shift"),
                "goal_met": r.get("goal_met"),
                "shift_goal": r.get("shift_goal"),
                "loads_completed": _to_int(r.get("loads_completed")),
                "total_shorts": _to_int(r.get("total_shorts")),
                "oc_total": _to_int(r.get("oc_total")),
                "oc_signoff_met": _to_int(r.get("oc_signoff_met")),
                "oc_photos_met": _to_int(r.get("oc_photos_met")),
                "cpu_total": _to_int(r.get("cpu_total")),
                "cpu_on_time": _to_int(r.get("cpu_on_time")),
                "notes": r.get("notes", ""),
            })
        if goal_total:
            goal_rate = round(100 * goal_met / goal_total)
    except Exception:
        pass

    per_shift.sort(key=lambda r: (str(r.get("date")), str(r.get("shift"))))

    misses = []
    for r in month_rows:
        if (
            str(r.get("on_time")).upper() == "N"
            or str(r.get("signoff_done")).upper() == "N"
            or str(r.get("photos_done")).upper() == "N"
            or str(r.get("short")).upper() == "Y"
        ):
            misses.append(r)

    shifts_logged = len(per_shift) or len({r.get("snapshot_id") for r in month_rows})

    return {
        "year": int(year),
        "month": int(month),
        "shifts_logged": shifts_logged,
        "loads_completed_total": loads_completed_total,
        "shorts_total": shorts_total,
        "oc_signoff": {"rate": signoff_rate, "met": signoff_met, "required": signoff_req},
        "oc_photos": {"rate": photos_rate, "met": photos_met, "required": photos_req},
        "cpu_on_time": {"rate": cpu_rate, "met": cpu_met, "total": cpu_total},
        "shift_goal": {"rate": goal_rate, "met": goal_met, "total": goal_total},
        "per_shift": per_shift,
        "misses": misses,
    }
def get_daily_scorecard(operating_date):
    """
    Goal performance for a single operating date. Same shape as the weekly/monthly
    scorecards so one page renders all three. operating_date may be a datetime.date
    or an mm/dd/YYYY string.
    """
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    if isinstance(operating_date, str):
        day = _parse_date(operating_date) or datetime.date.today()
    else:
        day = operating_date

    def _on_day(record):
        d = _parse_date(record.get("date"))
        return d is not None and d == day

    records = _read_tab_records(OUTCOMES_TAB, tuple(OUTCOMES_HEADER))
    day_rows = [r for r in records if _on_day(r)]

    oc = [r for r in day_rows if str(r.get("type")) == "OC"]
    cpu = [r for r in day_rows if str(r.get("type")) == "CPU"]

    def _rate(items, key):
        relevant = [r for r in items if str(r.get(key)).upper() in ("Y", "N")]
        if not relevant:
            return None, 0, 0
        met = sum(1 for r in relevant if str(r.get(key)).upper() == "Y")
        return round(100 * met / len(relevant)), met, len(relevant)

    signoff_rate, signoff_met, signoff_req = _rate(oc, "signoff_done")
    photos_rate, photos_met, photos_req = _rate(oc, "photos_done")
    cpu_rate, cpu_met, cpu_total = _rate(cpu, "on_time")

    goal_rate, goal_met, goal_total = None, 0, 0
    loads_completed_total = 0
    shorts_total = 0
    per_shift = []
    try:
        sum_records = _read_tab_records(SUMMARY_TAB, tuple(SUMMARY_HEADER))
        day_sum = [r for r in sum_records if _on_day(r)]
        for r in day_sum:
            counts, was_met = _goal_met_pair(r.get("goal_met"))
            if counts:
                goal_total += 1
                if was_met:
                    goal_met += 1
            loads_completed_total += _to_int(r.get("loads_completed"))
            shorts_total += _to_int(r.get("total_shorts"))
            per_shift.append({
                "date": r.get("date"),
                "shift": r.get("shift"),
                "goal_met": r.get("goal_met"),
                "shift_goal": r.get("shift_goal"),
                "loads_completed": _to_int(r.get("loads_completed")),
                "total_shorts": _to_int(r.get("total_shorts")),
                "oc_total": _to_int(r.get("oc_total")),
                "oc_signoff_met": _to_int(r.get("oc_signoff_met")),
                "oc_photos_met": _to_int(r.get("oc_photos_met")),
                "cpu_total": _to_int(r.get("cpu_total")),
                "cpu_on_time": _to_int(r.get("cpu_on_time")),
                "notes": r.get("notes", ""),
            })
        if goal_total:
            goal_rate = round(100 * goal_met / goal_total)
    except Exception:
        pass

    per_shift.sort(key=lambda r: (str(r.get("date")), str(r.get("shift"))))

    misses = []
    for r in day_rows:
        if (
            str(r.get("on_time")).upper() == "N"
            or str(r.get("signoff_done")).upper() == "N"
            or str(r.get("photos_done")).upper() == "N"
            or str(r.get("short")).upper() == "Y"
        ):
            misses.append(r)

    shifts_logged = len(per_shift) or len({r.get("snapshot_id") for r in day_rows})

    return {
        "date": day.strftime("%m/%d/%Y"),
        "shifts_logged": shifts_logged,
        "loads_completed_total": loads_completed_total,
        "shorts_total": shorts_total,
        "oc_signoff": {"rate": signoff_rate, "met": signoff_met, "required": signoff_req},
        "oc_photos": {"rate": photos_rate, "met": photos_met, "required": photos_req},
        "cpu_on_time": {"rate": cpu_rate, "met": cpu_met, "total": cpu_total},
        "shift_goal": {"rate": goal_rate, "met": goal_met, "total": goal_total},
        "per_shift": per_shift,
        "misses": misses,
    }

def get_weekly_scorecard(end_date):
    """
    Goal performance for the 7-day window ending on end_date (inclusive).
    Same shape as get_monthly_scorecard so one page renders both.
    end_date may be a datetime.date or an mm/dd/YYYY string.
    """
    if not is_configured():
        raise RuntimeError(f"Shift log not configured: {setup_hint()}")

    if isinstance(end_date, str):
        end = _parse_date(end_date) or datetime.date.today()
    else:
        end = end_date
    start = end - datetime.timedelta(days=6)

    def _in_week(record):
        d = _parse_date(record.get("date"))
        return d is not None and start <= d <= end

    records = _read_tab_records(OUTCOMES_TAB, tuple(OUTCOMES_HEADER))
    week_rows = [r for r in records if _in_week(r)]

    oc = [r for r in week_rows if str(r.get("type")) == "OC"]
    cpu = [r for r in week_rows if str(r.get("type")) == "CPU"]

    def _rate(items, key):
        relevant = [r for r in items if str(r.get(key)).upper() in ("Y", "N")]
        if not relevant:
            return None, 0, 0
        met = sum(1 for r in relevant if str(r.get(key)).upper() == "Y")
        return round(100 * met / len(relevant)), met, len(relevant)

    signoff_rate, signoff_met, signoff_req = _rate(oc, "signoff_done")
    photos_rate, photos_met, photos_req = _rate(oc, "photos_done")
    cpu_rate, cpu_met, cpu_total = _rate(cpu, "on_time")

    goal_rate, goal_met, goal_total = None, 0, 0
    loads_completed_total = 0
    shorts_total = 0
    per_shift = []
    try:
        sum_records = _read_tab_records(SUMMARY_TAB, tuple(SUMMARY_HEADER))
        week_sum = [r for r in sum_records if _in_week(r)]
        for r in week_sum:
            counts, was_met = _goal_met_pair(r.get("goal_met"))
            if counts:
                goal_total += 1
                if was_met:
                    goal_met += 1
            loads_completed_total += _to_int(r.get("loads_completed"))
            shorts_total += _to_int(r.get("total_shorts"))
            per_shift.append({
                "date": r.get("date"),
                "shift": r.get("shift"),
                "goal_met": r.get("goal_met"),
                "shift_goal": r.get("shift_goal"),
                "loads_completed": _to_int(r.get("loads_completed")),
                "total_shorts": _to_int(r.get("total_shorts")),
                "oc_total": _to_int(r.get("oc_total")),
                "oc_signoff_met": _to_int(r.get("oc_signoff_met")),
                "oc_photos_met": _to_int(r.get("oc_photos_met")),
                "cpu_total": _to_int(r.get("cpu_total")),
                "cpu_on_time": _to_int(r.get("cpu_on_time")),
                "notes": r.get("notes", ""),
            })
        if goal_total:
            goal_rate = round(100 * goal_met / goal_total)
    except Exception:
        pass

    per_shift.sort(key=lambda r: (str(r.get("date")), str(r.get("shift"))))

    misses = []
    for r in week_rows:
        if (
            str(r.get("on_time")).upper() == "N"
            or str(r.get("signoff_done")).upper() == "N"
            or str(r.get("photos_done")).upper() == "N"
            or str(r.get("short")).upper() == "Y"
        ):
            misses.append(r)

    shifts_logged = len(per_shift) or len({r.get("snapshot_id") for r in week_rows})

    return {
        "start": start.strftime("%m/%d/%Y"),
        "end": end.strftime("%m/%d/%Y"),
        "shifts_logged": shifts_logged,
        "loads_completed_total": loads_completed_total,
        "shorts_total": shorts_total,
        "oc_signoff": {"rate": signoff_rate, "met": signoff_met, "required": signoff_req},
        "oc_photos": {"rate": photos_rate, "met": photos_met, "required": photos_req},
        "cpu_on_time": {"rate": cpu_rate, "met": cpu_met, "total": cpu_total},
        "shift_goal": {"rate": goal_rate, "met": goal_met, "total": goal_total},
        "per_shift": per_shift,
        "misses": misses,
    }


# ============================================================
#  PRIVATE HELPERS
# ============================================================

def _to_int(value):
    try:
        return int(float(str(value).strip()))
    except Exception:
        return 0


def _goal_met_pair(value):
    """
    Interpret a stored goal_met value as (counts_toward_total, was_met).

    The closeout stores Daily Goal as a percent string ("87", "100"). This also
    still understands the legacy Y/N convention. A day "meets" the goal when it
    reaches 100% of the achievable goal. NA / blank / unparseable values do not
    count toward the denominator at all.
    """
    text = str(value).strip().upper()
    if text in ("Y", "N"):
        return True, text == "Y"
    if text in ("", "NA", "N/A", "NONE"):
        return False, False
    try:
        pct = float(text.rstrip("%"))
    except ValueError:
        return False, False
    return True, pct >= 100


def _replace_rows_for_snapshot(ws, header, snapshot_id, new_rows):
    """
    Remove existing rows for this snapshot_id, then append new_rows. Read-filter-
    rewrite is fine at this volume and keeps writes idempotent so re-runs never
    duplicate a day.
    """
    all_values = ws.get_all_values()
    body = all_values[1:] if all_values else []

    id_col = header.index("snapshot_id")
    kept = [row for row in body if (len(row) > id_col and row[id_col] != snapshot_id)]

    ws.clear()
    ws.append_row(header, value_input_option="USER_ENTERED")
    final = kept + new_rows
    if final:
        ws.append_rows(final, value_input_option="USER_ENTERED")


def _parse_date(text):
    for fmt in ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"):
        try:
            return datetime.datetime.strptime(str(text).strip(), fmt).date()
        except Exception:
            continue
    return None
