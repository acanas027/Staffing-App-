"""
dc_config.py
============
Single source of truth for the distribution center's operational constants.

WHY THIS FILE EXISTS
--------------------
These numbers were previously scattered across the staffing report and the shift
closeout. When the same concept lived in two files, the two could drift apart (e.g.
the shift-1 end time was 16:30 in one place and 16:00 in another). Holding them here
means each value is defined once; every page imports it, so they can never disagree.

WHAT BELONGS HERE
-----------------
Only constants — values that are true about the DC and change rarely (a few times a
year at most): work rates, staffing floors, shift times, plant assumptions. Anything
that changes per shift (who is present, today's loads, the selected day) is a runtime
input and does NOT belong here.

HOW TO USE
----------
    import dc_config
    ...
    capacity = pickers * dc_config.PICK_RATE * hours

Change a value here once and every page that imports it picks up the change.
"""

# ============================================================
#  WORK RATES  (per person, per hour)
# ============================================================
PICK_RATE = 185.0    # cases picked per hour per person
PULL_RATE = 25.0     # full pallets pulled per hour per person
LOAD_RATE = 1.0      # trailers loaded per hour per person
UNLOAD_RATE = 44.0   # inbound pallets unloaded per hour per person


# ============================================================
#  STAFFING RULES
# ============================================================
# Always-on Tasking floor (replenishment + putaway) before any full-pallet
# pull taskers are added on top.
TASK_FLOOR = 4

# Reserved inbound crew taken off the top before the Picking/Loading/pull-tasker
# split. These protect inbound flow no matter how heavy outbound is.
MIN_UNLOADERS = 2
MIN_RECEIVERS = 2

# 1st-shift loading target = this share of the selected-day outbound loads,
# on top of whatever is already completed/loaded.
LOAD_TARGET_SHARE = 0.52


# ============================================================
#  SHIFT TIMES
# ============================================================
# 1st shift runs 06:00 to 16:30. 16:30 is the single end-of-shift boundary used
# everywhere — the planning target AND the closeout's "which loads are mine" window.
# 2nd shift runs 17:00 to 05:00 the next day (wraps past midnight).
#
# Two forms of each value are provided because different code wants different forms:
#   *_LABEL  -> human-readable "HH:MM" string (for goals, report text)
#   *_MIN    -> minutes since midnight as an int (for time-window math/comparisons)

SHIFT_1_START_LABEL = "06:00"
SHIFT_1_END_LABEL = "16:30"
SHIFT_2_START_LABEL = "17:00"
SHIFT_2_END_LABEL = "05:00"

SHIFT_1_START_MIN = 6 * 60          # 360
SHIFT_1_END_MIN = 16 * 60 + 30      # 990  (16:30)
SHIFT_2_START_MIN = 17 * 60         # 1020
SHIFT_2_END_MIN = 5 * 60            # 300  (05:00 next day; window wraps midnight)


def shift_end_label(shift):
    """Return the end-of-shift label for '1st' or '2nd'."""
    return SHIFT_1_END_LABEL if "1" in str(shift) else SHIFT_2_END_LABEL


def in_shift_window(appt_minutes, shift):
    """
    True if an appointment (given in minutes since midnight) falls in this shift's
    window. 1st shift is a simple range; 2nd shift wraps past midnight.
    Pass None through as True so a blank/unparseable time is never silently dropped.
    """
    if appt_minutes is None:
        return True
    if "1" in str(shift):
        return SHIFT_1_START_MIN <= appt_minutes <= SHIFT_1_END_MIN
    # 2nd shift: 17:00..23:59 OR 00:00..05:00
    return appt_minutes >= SHIFT_2_START_MIN or appt_minutes <= SHIFT_2_END_MIN


# ============================================================
#  PLANT INBOUND ASSUMPTIONS  (pallets per open plant)
# ============================================================
# Used to estimate inbound workload when a plant is open. These are planning
# assumptions, not live counts.
PLANT_INBOUND_PALLETS = {
    "Crossroads": 700,
    "Deer Creek": 500,
    "MSB": 640,
}
