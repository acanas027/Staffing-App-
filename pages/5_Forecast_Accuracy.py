"""
5_Forecast_Accuracy.py
======================
How well the MORNING PLAN predicted the day — separate from the Performance
Scorecard on purpose. The Performance Scorecard grades the shift (did the
warehouse hit its targets). This page grades the TOOL: did the morning staffing
report's predicted appointment cutoff match what actually happened, and how many
loads fell behind the cutoff goal.

Reads the forecast_accuracy tab the daily closeout writes (1st shift for now).
This is NOT a supervisor metric — a "BEHIND" day can mean the model was too
optimistic, not that the shift underperformed. Trends only become meaningful
once several days of closeouts accumulate.
"""

import datetime

import pandas as pd
import streamlit as st

import shift_log


# ============================================================
#  PAGE
# ============================================================

st.set_page_config(page_title="Forecast Accuracy", layout="wide")
st.title("Forecast Accuracy")
st.write(
    "How well the morning plan predicted the day. This measures the planning tool, "
    "not the supervisor: it compares the morning report's predicted appointment "
    "cutoff against what actually happened, and counts loads that fell behind the "
    "cutoff goal. A 'behind' day can mean the model was too optimistic, not that "
    "the shift underperformed. Currently 1st shift only."
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

window = st.radio("Window", ["Last 7 days", "Last 30 days", "Last 90 days"], horizontal=True, index=1)
days = {"Last 7 days": 7, "Last 30 days": 30, "Last 90 days": 90}[window]

try:
    fa = shift_log.get_forecast_accuracy(days=days)
except Exception as e:
    st.error(f"Could not load forecast accuracy: {e}")
    st.stop()

if fa["count"] == 0:
    st.warning(
        f"No forecast-accuracy rows recorded in the {window.lower()} yet. "
        "Each daily closeout writes one row (1st shift). Once a few days are in, "
        "this page shows how the morning predictions are trending."
    )
    st.stop()

st.caption(f"Based on {fa['count']} closed day(s) in the {window.lower()}.")

# Headline roll-ups
m1, m2, m3 = st.columns(3)

m1.metric("Avg loads short of cutoff", fa["avg_loads_short"] if fa["avg_loads_short"] is not None else "—")
m1.caption("Avg controllable loads behind the cutoff goal, per day.")

m2.metric("Days behind plan", f"{fa['behind_days']} of {fa['count']}")
m2.caption("Days the actual cutoff came in earlier than predicted.")

m3.metric("No-departure loads (total)", fa["total_no_dep"])
m3.caption("Loads with no departure logged — verify with the OpenDock clerk.")

# Per-day table
st.markdown("---")
st.subheader("Day by day")
st.dataframe(
    pd.DataFrame([
        {
            "Date": r.get("date"),
            "Shift": r.get("shift"),
            "Predicted cutoff": r.get("predicted_cutoff") or "—",
            "Actual cutoff": r.get("actual_cutoff") or "—",
            "Direction": r.get("direction") or "—",
            "Loads short": r.get("loads_short"),
            "No departure logged": r.get("no_departure_count"),
        }
        for r in fa["rows"]
    ]),
    use_container_width=True,
    height=360,
)

st.caption(
    "Direction compares the actual appointment cutoff to the morning prediction. "
    "'Loads short' counts controllable loads due by the goal cutoff that did not "
    "depart on time (drops excluded — carrier timing). 'No departure logged' is a "
    "data-quality check, not a miss: it flags loads the OpenDock clerk may not have "
    "scanned out."
)
