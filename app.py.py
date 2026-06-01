import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from io import BytesIO
import os
import shutil
import re
import json
import datetime
from openai import OpenAI


st.set_page_config(page_title="Staffing Report Generator", layout="wide")

st.title("Staffing Report Generator")
st.write("Enter daily inputs, select who is present, and generate the staffing report.")

TEMPLATE_FILE = "staffing_template.xlsx"


if not os.path.exists(TEMPLATE_FILE):
    st.error("Template file not found. Put staffing_template.xlsx in the same folder as report.py.")
    st.stop()


#  OPPORTUNITY CUSTOMER LIST (loaded from Excel) 
# File must be in the same folder as report.py.
# Sheet: "OC Customer List"
# Row 6  = headers (skipped by name check)
# Row 7  = example row — skipped (customer name contains "market x" / "example")
# Rows 8+ = real data
# Columns:
#   A: Resers DC   B: Customer #   C: Customer Name   D: Address
#   E: Profile/Why OC   F: DC Requirements   G: Sign Off (Y/N)
#   H: Pictures (Y/N)   I: Other (Y/N)

OC_FILE = "Resers DCs Opportunity Cusotmer List.xlsx"
OC_SHEET = "OC Customer List"
OC_HEADER_ROW = 6   # 1-based row number of the header row
OC_DATA_START = 8   # first real data row (row 7 is the example, skip it)


@st.cache_data
def load_oc_customer_list():
    """
    Read the OC Excel file and return a list of customer dicts identical in
    shape to the old embedded OC_CUSTOMER_LIST.  Cached so it only reads once
    per app session.
    """
    if not os.path.exists(OC_FILE):
        st.error(
            f"OC customer list file not found: '{OC_FILE}'. "
            "Make sure it is in the same folder as report.py."
        )
        return []

    try:
        wb = load_workbook(OC_FILE, data_only=True)
        if OC_SHEET not in wb.sheetnames:
            st.error(f"Sheet '{OC_SHEET}' not found in {OC_FILE}.")
            return []

        ws = wb[OC_SHEET]
        customers = []

        for row_idx in range(OC_DATA_START, ws.max_row + 1):
            raw_name = ws.cell(row_idx, 3).value  # col C
            if not raw_name:
                continue

            name_clean = str(raw_name).strip().strip('"').lower()

            # Skip the example row
            if "market x" in name_clean or "example" in name_clean:
                continue

            raw_cust_num = ws.cell(row_idx, 2).value  # col B
            raw_issue    = ws.cell(row_idx, 5).value  # col E
            raw_reqs     = ws.cell(row_idx, 6).value  # col F
            raw_signoff  = ws.cell(row_idx, 7).value  # col G
            raw_pictures = ws.cell(row_idx, 8).value  # col H

            issue = str(raw_issue).strip() if raw_issue else ""
            reqs  = str(raw_reqs).strip()  if raw_reqs  else ""

            sign_off = str(raw_signoff).strip().upper() == "Y" if raw_signoff else False
            pictures = str(raw_pictures).strip().upper() == "Y" if raw_pictures else False

            # Priority: HIGH if sign-off or pictures required, else MEDIUM
            priority = "HIGH" if (sign_off or pictures) else "MEDIUM"

            # Build search aliases from the customer name
            # e.g. "Sobey's - All Loads" → also match "sobey", "sobeys", "sobey's"
            base = name_clean.rstrip(" -").split(" - ")[0].strip()
            aliases = []
            # Strip common suffixes to get short-match terms
            for suffix in [" - all loads", " all loads", " fresh dc", " (olathe)"]:
                if base.endswith(suffix):
                    aliases.append(base.replace(suffix, "").strip())
            # Add apostrophe variants
            if "'" in base:
                aliases.append(base.replace("'", ""))
                aliases.append(base.replace("'s", ""))
            # Common known aliases
            known_aliases = {
                "target rialto":        ["target"],
                "sobey's - all loads":  ["sobeys", "sobey", "sobey's"],
                "sysco kc (olathe)":    ["sysco kc", "sysco kansas city", "sysco olathe", "sysco kc olathe"],
                "pfs virgina":          ["pfs virginia", "pfs va"],
                "metro toronto fresh dc": ["metro toronto", "metro fresh"],
                "jewel's":              ["jewels", "jewel"],
                "awg":                  ["associated wholesale grocers"],
                "whataburguer":         ["whataburger"],
            }
            if name_clean in known_aliases:
                aliases += known_aliases[name_clean]

            # Deduplicate aliases, remove if same as name
            aliases = list(dict.fromkeys(
                a for a in aliases if a and a != name_clean
            ))

            customers.append({
                "name": name_clean,
                "aliases": aliases,
                "customer_number": str(raw_cust_num).strip() if raw_cust_num else None,
                "issue": issue,
                "requirements": reqs,
                "sign_off": sign_off,
                "pictures": pictures,
                "priority": priority,
            })

        return customers

    except Exception as e:
        st.error(f"Error loading OC customer list: {e}")
        return []


def find_oc_customers_in_board(board_text):
    oc_list = load_oc_customer_list()
    board_lower = board_text.lower()
    matches = []
    for customer in oc_list:
        search_terms = [customer["name"]] + customer.get("aliases", [])
        found_terms = [term for term in search_terms if term.lower() in board_lower]
        if found_terms:
            matches.append({"customer": customer, "matched_on": found_terms})
    return matches


def build_oc_alert_text(oc_matches):
    if not oc_matches:
        return None
    lines = [
        "=== OPPORTUNITY CUSTOMER (OC) ALERT ===",
        "The following loads belong to customers on the Opportunity Customer List.",
        "These customers have a documented history of complaints and require special handling.",
        "Flag these loads explicitly in your analysis and include action items for each.",
        "",
    ]
    for match in oc_matches:
        c = match["customer"]
        lines.append(f"CUSTOMER: {c['name'].upper()}")
        lines.append(f"Matched on: {', '.join(match['matched_on'])}")
        lines.append(f"Priority: {c['priority']}")
        lines.append(f"Issue History: {c['issue']}")
        lines.append(f"DC Requirements: {c['requirements']}")
        if c["sign_off"]:
            lines.append("DC Supervisor Sign-Off REQUIRED before this load ships.")
        if c["pictures"]:
            lines.append("Photos REQUIRED: 3 on dock + 3 during loading (6 total). Email to manager.")
        lines.append("")
    lines += [
        "IMPORTANT: For every OC load identified above:",
        "1. Flag it clearly in your Board Summary section.",
        "2. Add a dedicated OC Action section with specific steps before this load ships.",
        "3. Recommend who should own the sign-off and photo process.",
        "4. Include this as one of the Top 3 Action Items if the load is active or upcoming.",
        "=== END OC ALERT ===",
    ]
    return "\n".join(lines)


def get_groq_client():
    if "GROQ_API_KEY" not in st.secrets:
        return None
    return OpenAI(
        api_key=st.secrets["GROQ_API_KEY"],
        base_url="https://api.groq.com/openai/v1",
    )


@st.cache_data
def load_names():
    wb = load_workbook(TEMPLATE_FILE, data_only=False)
    ws = wb["Inputs"]
    names = []
    # Dynamic scan — no hardcoded row limit. Stops after 10 consecutive empty rows.
    consecutive_empty = 0
    for row in range(3, ws.max_row + 1):
        name = ws[f"E{row}"].value
        if name and str(name).strip():
            names.append(str(name).strip())
            consecutive_empty = 0
        else:
            consecutive_empty += 1
            if consecutive_empty >= 10:
                break
    return names


names = load_names()


def whole_workers(value):
    return int(float(value or 0) + 0.7)


def is_present(row):
    return str(row["Present"]).strip().lower() == "x"


def has_skill(row, code):
    return code in str(row["Skills"])


def best_fit(row, text):
    return text.lower() in str(row["Best Fit"]).lower()


def name_contains(row, text):
    return text.lower() in str(row["Name"]).lower()


def calculate_input_values(day, shift, total_cases):
    first_shift_pick = {
        "Sunday": 0.20, "Monday": 0.18, "Tuesday": 0.18, "Wednesday": 0.19,
        "Thursday": 0.19, "Friday": 0.18, "Saturday": 0.21,
    }
    second_shift_pick = {
        "Sunday": 0.19, "Monday": 0.15, "Tuesday": 0.15, "Wednesday": 0.17,
        "Thursday": 0.17, "Friday": 0.17, "Saturday": 0.19,
    }
    first_shift_fp = {
        "Sunday": 0.28, "Monday": 0.32, "Tuesday": 0.40, "Wednesday": 0.35,
        "Thursday": 0.35, "Friday": 0.36, "Saturday": 0.31,
    }
    second_shift_fp = {
        "Sunday": 0.32, "Monday": 0.33, "Tuesday": 0.27, "Wednesday": 0.29,
        "Thursday": 0.28, "Friday": 0.30, "Saturday": 0.30,
    }
    if shift == "1st":
        cases_to_pick = total_cases * first_shift_pick.get(day, 0)
        full_pallets = (total_cases * first_shift_fp.get(day, 0)) / 70
    else:
        cases_to_pick = total_cases * second_shift_pick.get(day, 0)
        full_pallets = (total_cases * second_shift_fp.get(day, 0)) / 70
    return cases_to_pick, full_pallets


def calculate_needed(
    day, shift, total_cases, hours_remaining, total_outbound_loads_actual,
    crossroads_open, deer_creek_open, msb_open,
):
    if hours_remaining <= 0:
        hours_remaining = 1
    cases_to_pick, full_pallets = calculate_input_values(day, shift, total_cases)
    inbound_pallets = 0
    if crossroads_open == "YES":
        inbound_pallets += 700
    if deer_creek_open == "YES":
        inbound_pallets += 500
    if msb_open == "YES":
        inbound_pallets += 640
    raw_needed = {
        "Unloading": (inbound_pallets / 4) / (44 * hours_remaining),
        "Receiving": (inbound_pallets / 4) / (44 * hours_remaining),
        "Putaway": (inbound_pallets / 2) / (25 * hours_remaining),
        "Picking": cases_to_pick / (185 * hours_remaining),
        "Replenishment": (cases_to_pick / 70) / (25 * 8.5),
        "Full Pallets": full_pallets / (25 * hours_remaining),
        "Loading": total_outbound_loads_actual / hours_remaining,
    }
    needed = {
        "Unloading": max(2, whole_workers(raw_needed["Unloading"])),
        "Receiving": max(2, whole_workers(raw_needed["Receiving"])),
        "Picking": whole_workers(raw_needed["Picking"]),
        "Tasking": whole_workers(
            raw_needed["Putaway"] + raw_needed["Replenishment"] + raw_needed["Full Pallets"]
        ),
        "Loading": whole_workers(raw_needed["Loading"]),
    }
    return needed, raw_needed, cases_to_pick, full_pallets, inbound_pallets


def generate_recommendations(staff, needed):
    assigned = {task: 0 for task in needed}
    staff["Recommended Task"] = ""
    present_indexes = staff[staff.apply(is_present, axis=1)].index.tolist()

    def assign_if_needed(task, idx):
        if assigned[task] < needed[task]:
            staff.at[idx, "Recommended Task"] = task
            assigned[task] += 1
            return True
        return False

    for idx in present_indexes:
        row = staff.loc[idx]
        if name_contains(row, "Dale"):
            staff.at[idx, "Recommended Task"] = "Receiving"
            assigned["Receiving"] += 1
        elif name_contains(row, "Alex"):
            staff.at[idx, "Recommended Task"] = "Unloading"
            assigned["Unloading"] += 1

    for idx in present_indexes:
        if staff.at[idx, "Recommended Task"] != "":
            continue
        row = staff.loc[idx]
        if str(row["Skills"]).strip() == "P":
            assign_if_needed("Picking", idx)

    best_fit_steps = [
        ("Unloading", "Unload", "U"),
        ("Loading", "Load", "L"),
        ("Receiving", "Receiv", "R"),
        ("Picking", "Pick", "P"),
        ("Tasking", "Task", "T"),
    ]
    for task, fit_text, skill in best_fit_steps:
        for idx in present_indexes:
            if staff.at[idx, "Recommended Task"] != "":
                continue
            row = staff.loc[idx]
            if best_fit(row, fit_text) and has_skill(row, skill):
                assign_if_needed(task, idx)

    skill_map = {"Unloading": "U", "Receiving": "R", "Loading": "L", "Picking": "P", "Tasking": "T"}
    for task, skill in skill_map.items():
        for idx in present_indexes:
            if assigned[task] >= needed[task]:
                break
            if staff.at[idx, "Recommended Task"] != "":
                continue
            row = staff.loc[idx]
            if has_skill(row, skill):
                assign_if_needed(task, idx)

    backup_tasks = ["Unloading", "Receiving", "Loading", "Picking", "Tasking"]
    for task in backup_tasks:
        while assigned[task] < needed[task]:
            found_worker = False
            for idx in present_indexes:
                if staff.at[idx, "Recommended Task"] != "":
                    continue
                row = staff.loc[idx]
                if best_fit(row, "Task") and (has_skill(row, "T") or has_skill(row, "L") or has_skill(row, "P")):
                    assign_if_needed(task, idx)
                    found_worker = True
                    break
            if found_worker:
                continue
            for idx in present_indexes:
                if staff.at[idx, "Recommended Task"] != "":
                    continue
                row = staff.loc[idx]
                if has_skill(row, "T") or has_skill(row, "L") or has_skill(row, "P"):
                    assign_if_needed(task, idx)
                    found_worker = True
                    break
            if not found_worker:
                break

    for idx in present_indexes:
        if staff.at[idx, "Recommended Task"] == "":
            if assigned["Tasking"] < needed["Tasking"]:
                staff.at[idx, "Recommended Task"] = "Tasking"
                assigned["Tasking"] += 1
            else:
                staff.at[idx, "Recommended Task"] = "Lead/Extra"

    preferred_extra_names = ["will", "antonio"]
    preferred_idxs = [
        idx for idx in present_indexes
        if any(name in str(staff.at[idx, "Name"]).lower() for name in preferred_extra_names)
    ]
    current_extra_idxs = [
        idx for idx in present_indexes
        if staff.at[idx, "Recommended Task"] == "Lead/Extra"
    ]
    for preferred_idx in preferred_idxs:
        if not current_extra_idxs:
            break
        if staff.at[preferred_idx, "Recommended Task"] == "Lead/Extra":
            continue
        swap_idx = None
        for extra_idx in current_extra_idxs:
            if not any(name in str(staff.at[extra_idx, "Name"]).lower() for name in preferred_extra_names):
                swap_idx = extra_idx
                break
        if swap_idx is None:
            break
        old_task = staff.at[preferred_idx, "Recommended Task"]
        staff.at[preferred_idx, "Recommended Task"] = "Lead/Extra"
        staff.at[swap_idx, "Recommended Task"] = old_task
        current_extra_idxs.remove(swap_idx)

    return staff


def build_summary(staff, needed):
    present_recommendations = staff[
        staff["Present"].astype(str).str.strip().str.lower().eq("x")
        & staff["Recommended Task"].astype(str).str.strip().ne("")
    ].copy()
    needed_list = pd.Series(needed, name="Needed")
    assigned_list = present_recommendations["Recommended Task"].value_counts().rename("Assigned")
    summary_table = pd.concat([needed_list, assigned_list], axis=1).fillna(0)
    summary_table["Needed"] = summary_table["Needed"].astype(int)
    summary_table["Assigned"] = summary_table["Assigned"].astype(int)
    summary_table["Difference"] = summary_table["Assigned"] - summary_table["Needed"]
    summary_table["Status"] = summary_table["Difference"].apply(
        lambda x: "Good" if x == 0 else ("Overstaffed" if x > 0 else "Understaffed")
    )
    return present_recommendations, summary_table


def build_recommendations(summary_table, present_recommendations, raw_needed, hours_remaining, notes):
    recommendations = []
    total_labor_gap = int(summary_table["Difference"].sum())
    labor_hours_gap = total_labor_gap * hours_remaining
    recommendations.append(
        f"Current labor balance estimate: {labor_hours_gap:+.1f} labor-hours. "
        f"Positive means extra capacity; negative means short capacity."
    )
    for task, row in summary_table.iterrows():
        diff = int(row["Difference"])
        if diff < 0:
            recommendations.append(
                f"{task}: approximately {abs(diff * hours_remaining):.1f} labor-hours behind based on current staffing vs need."
            )
        elif diff > 0:
            recommendations.append(f"{task}: approximately {diff * hours_remaining:.1f} labor-hours ahead / available capacity.")
        else:
            recommendations.append(f"{task}: Staffing is balanced.")

    picking_gap = int(summary_table.loc["Picking", "Difference"]) if "Picking" in summary_table.index else 0
    tasking_gap = int(summary_table.loc["Tasking", "Difference"]) if "Tasking" in summary_table.index else 0
    receiving_gap = int(summary_table.loc["Receiving", "Difference"]) if "Receiving" in summary_table.index else 0
    unloading_gap = int(summary_table.loc["Unloading", "Difference"]) if "Unloading" in summary_table.index else 0
    loading_gap = int(summary_table.loc["Loading", "Difference"]) if "Loading" in summary_table.index else 0
    lead_gap = int(summary_table.loc["Lead/Extra", "Difference"]) if "Lead/Extra" in summary_table.index else 0

    if picking_gap < 0:
        recommendations.append("High picking short risk detected. Consider moving tasking labor into replenishment to protect pickers.")
        recommendations.append("Avoid pulling pickers into unloading or loading unless outbound service is critical.")
        if tasking_gap > 0:
            recommendations.append(f"Tasking currently has {tasking_gap} extra worker(s). Consider temporarily assigning them to replenishment.")
        if lead_gap > 0:
            recommendations.append("Lead/Extra capacity available. Consider flexing extra labor into replenishment or picking support.")
    if unloading_gap < 0 or receiving_gap < 0:
        recommendations.append("Inbound flow risk detected. Falling behind may create dock congestion and delayed putaway.")
        recommendations.append("Consider moving flexible tasking labor into unloading or receiving temporarily.")
        if tasking_gap > 1:
            recommendations.append("Tasking has available labor that can support inbound operations.")
    if loading_gap < 0:
        recommendations.append("Outbound loading risk detected. Late departures and service failures may increase.")
        recommendations.append("Protect loading labor before reallocating to non-critical work.")
        if lead_gap > 0:
            recommendations.append("Use Lead/Extra labor to support outbound staging or trailer cleanup.")
    if total_labor_gap > 1:
        recommendations.append("Operation currently has excess labor capacity.")
        recommendations.append("Consider deep cleaning, trailer audits, replenishment cleanup, or cross-training.")
        recommendations.append("Extra labor could be used proactively to prevent later picking shortages.")

    inbound_pressure = raw_needed["Unloading"] + raw_needed["Receiving"] + raw_needed["Putaway"]
    outbound_pressure = raw_needed["Picking"] + raw_needed["Loading"]
    if inbound_pressure > outbound_pressure * 1.3:
        recommendations.append("Inbound workload is significantly heavier than outbound.")
        recommendations.append("Focus on unloading, receiving, and putaway to avoid congestion.")
    elif outbound_pressure > inbound_pressure * 1.3:
        recommendations.append("Outbound workload is significantly heavier than inbound.")
        recommendations.append("Prioritize replenishment and picking continuity to avoid shorts.")

    if hours_remaining <= 4:
        recommendations.append("Shift is entering final hours. Prioritize completion work and outbound execution.")
    elif hours_remaining >= 8:
        recommendations.append("Enough shift time remains to strategically rebalance labor before bottlenecks form.")

    lower_notes = notes.lower()
    if "late" in lower_notes:
        recommendations.append("Manager notes mention late loads. Prioritize outbound execution and trailer readiness.")
    if "short" in lower_notes:
        recommendations.append("Manager notes indicate short risk. Protect replenishment and picking flow.")
    if "live" in lower_notes:
        recommendations.append("Live loads detected in notes. Prioritize those doors before drop trailers.")
    if "cpu" in lower_notes:
        recommendations.append("CPU loads referenced. Ensure loading labor is protected.")

    return recommendations


#  BOARD EXCEL READING 
# Python reads ALL cell values and color flags directly from the xlsx.
# Board layout (columns A–M):
#   A: Load#  B: Destination  C: Carrier  D: Time  E: Door  F: Trailer
#   G: Status  H: TT4  I: Loader  J: Comments  K: Pulls  L: Picks  M: Priority
#
# Day header rows have a day name in col A and a date in col B.
# All other meaningful rows have a 5-9 digit load number in col A.

BOARD_DAY_NAMES = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]


def normalize_board_text(value):
    """Convert any Excel cell value to a clean string."""
    if value is None:
        return ""
    # Handle datetime.time objects (Excel stores times this way)
    if isinstance(value, datetime.time):
        return value.strftime("%H:%M")
    # Handle datetime.datetime objects
    if isinstance(value, datetime.datetime):
        return value.strftime("%m/%d/%Y")
    # Handle pandas NaT / NaN
    try:
        import math
        if isinstance(value, float) and math.isnan(value):
            return ""
    except Exception:
        pass
    text = str(value).replace("\n", " ").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def normalize_board_date(value):
    text = normalize_board_text(value)
    if not text:
        return ""
    for fmt in ["%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d"]:
        try:
            return pd.to_datetime(text, format=fmt).strftime("%m/%d/%Y")
        except Exception:
            pass
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%m/%d/%Y")
    except Exception:
        pass
    return text


def normalize_board_time(value):
    text = normalize_board_text(value)
    if not text:
        return ""
    if re.fullmatch(r"\d{1,2}:\d{2}", text):
        h, m = text.split(":")
        return f"{int(h):02d}:{m}"
    try:
        parsed = pd.to_datetime(text, errors="coerce")
        if pd.notna(parsed):
            return parsed.strftime("%H:%M")
    except Exception:
        pass
    return text


def looks_like_board_load(value):
    """Return the digit string if value looks like a load number, else ''."""
    text = normalize_board_text(value)
    if not text:
        return ""
    if text in BOARD_DAY_NAMES:
        return ""
    if re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", text):
        return ""
    if re.fullmatch(r"\d{1,2}:\d{2}", text):
        return ""
    digits = re.sub(r"[^0-9]", "", text)
    if 5 <= len(digits) <= 9:
        return digits
    return ""


def detect_board_status(value):
    text = normalize_board_text(value).upper()
    if not text:
        return ""
    if "LOADED SHORT" in text:
        return "Loaded Short"
    if "PICKING/SHORT" in text or "PICKING SHORT" in text:
        return "Picking/Short"
    if "READY/SHORT" in text or re.search(r"\bR/S\b", text):
        return "R/S"
    if re.search(r"\bRTL\b", text) or "READY TO LOAD" in text:
        return "RTL"
    if "NO DRIVER" in text:
        return "No Driver"
    if "PICKING" in text:
        return "Picking"
    if "COMPLETED" in text or re.search(r"\bCOMPLETE\b", text):
        return "Completed"
    if re.search(r"\bLATE\b", text):
        return "Late"
    if re.search(r"\bLOADED\b", text):
        return "Loaded"
    return ""


def detect_trailer_field_late(trailer_value):
    """
    The board sometimes writes LATE or ETA <time> in the trailer/door column (col 6)
    instead of the status column. Detect this specifically for col 6 only —
    NOT for the time column (col 4) where ETA just means the driver hasn't arrived yet.
    """
    text = normalize_board_text(trailer_value).upper()
    if not text:
        return False
    if re.search(r"\bLATE\b", text):
        return True
    # ETA in the trailer field = driver running late, load overdue
    if re.match(r"^ETA\b", text):
        return True
    return False


def board_cell_flags(cell):
    """Read fill color and font color to detect special flags."""
    flags = []
    fill_color = ""
    font_color = ""
    try:
        fill = cell.fill
        if fill and fill.fgColor:
            if fill.fgColor.type == "rgb":
                fill_color = str(fill.fgColor.rgb).upper()
            elif fill.fgColor.type == "indexed":
                fill_color = str(fill.fgColor.indexed).upper()
    except Exception:
        pass
    try:
        font = cell.font
        if font and font.color and font.color.type == "rgb":
            font_color = str(font.color.rgb).upper()
    except Exception:
        pass
    # Yellow fill → load check
    if fill_color in ("FFFFFF00", "00FFFF00", "FFFF00", "0000000D"):
        flags.append("LOAD-CHECK")
    # Light blue fill → TT4 needed
    if fill_color in ("FFADD8E6", "FF87CEEB", "FFADD8FF", "FFB0E0E6", "FF00BFFF"):
        flags.append("TT4-NEEDED")
    # Red font → Canadian load
    if font_color in ("FFFF0000", "00FF0000"):
        flags.append("CANADIAN")
    return flags


def parse_number(value):
    text = normalize_board_text(value)
    if not text or text.strip() in ("", " "):
        return 0
    digits = re.sub(r"[^0-9]", "", text)
    return int(digits) if digits else 0


def board_records_from_excel(board_file):
    """
    Read every sheet of the uploaded .xlsx board file.
    Returns a list of structured load-row dicts.
    Handles:
      - datetime.time values in the Time column
      - Pure numeric load numbers (Excel stores as int/float)
      - Day header rows that set the current day/date context
      - Color flags: LOAD-CHECK (yellow), TT4-NEEDED (blue), CANADIAN (red font)
    """
    board_file.seek(0)
    file_name = board_file.name.lower()
    all_rows = []

    if file_name.endswith(".xls"):
        # Legacy .xls — openpyxl can't read it, use pandas (no color support)
        sheets = pd.read_excel(board_file, sheet_name=None, header=None, engine="xlrd")
        for sheet_name, df in sheets.items():
            df = df.fillna("")
            current_day = ""
            current_date = ""
            for idx, row in df.iterrows():
                values = [normalize_board_text(v) for v in row.tolist()]
                while len(values) < 13:
                    values.append("")
                first_cell = values[0]
                if first_cell in BOARD_DAY_NAMES:
                    current_day = first_cell
                    current_date = normalize_board_date(values[1])
                    continue
                load_number = looks_like_board_load(values[0])
                if not load_number:
                    continue
                if detect_trailer_field_late(values[5]):
                    status = "Late"
                else:
                    status = detect_board_status(values[6])
                if not status:
                    status = detect_board_status(" ".join(values))
                trailer_text = values[5].upper()
                type_value = "Live" if "LIVE" in trailer_text else ("CPU - Live" if "CPU" in trailer_text else ("Drop" if "DROP" in trailer_text else ""))
                all_rows.append({
                    "source": sheet_name,
                    "day": current_day,
                    "date": current_date,
                    "load_number": load_number,
                    "customer": values[1],
                    "carrier": values[2],
                    "appt_time": normalize_board_time(values[3]),
                    "door": values[4],
                    "trailer": values[5],
                    "status": status,
                    "type": type_value,
                    "tt4": values[7],
                    "loader": values[8],
                    "comments": values[9],
                    "pulls": parse_number(values[10]),
                    "picks": parse_number(values[11]),
                    "flags": [],
                    "raw_row": " | ".join(v for v in values if v),
                })
        return all_rows

    # .xlsx — full read with color flag detection
    # Only read the Outbound sheet. Other sheets (Inbound, formats, etc.)
    # contain non-outbound data that would create phantom "Unknown Day" loads.
    wb = load_workbook(board_file, data_only=True)
    outbound_sheet = None
    for candidate in ["Outbound", "outbound", "OUTBOUND"]:
        if candidate in wb.sheetnames:
            outbound_sheet = candidate
            break
    sheets_to_read = [outbound_sheet] if outbound_sheet else wb.sheetnames
    for sheet_name in sheets_to_read:
        ws = wb[sheet_name]
        current_day = ""
        current_date = ""
        consecutive_empty = 0  # stop early when Excel phantom rows start
        for row_idx in range(1, ws.max_row + 1):
            values = []
            flags = []
            has_content = False
            for col_idx in range(1, 14):
                cell = ws.cell(row_idx, col_idx)
                if cell.value is not None:
                    has_content = True
                values.append(normalize_board_text(cell.value))
                for flag in board_cell_flags(cell):
                    flags.append(flag)

            if not has_content:
                consecutive_empty += 1
                if consecutive_empty >= 15:
                    break  # Excel reports ws.max_row incorrectly due to phantom formatting
                continue
            consecutive_empty = 0

            first_cell = values[0]
            if first_cell in BOARD_DAY_NAMES:
                current_day = first_cell
                current_date = normalize_board_date(values[1])
                continue

            load_number = looks_like_board_load(values[0])
            if not load_number:
                continue

            # Check col 6 (trailer/door field) for LATE or ETA FIRST —
            # the board writes these there for overdue loads from prior days.
            # Late takes priority over RTL/R/S since those still apply but
            # the load is also overdue.
            if detect_trailer_field_late(values[5]):
                status = "Late"
            else:
                status = detect_board_status(values[6])
            if not status:
                status = detect_board_status(" ".join(values))

            trailer_text = values[5].upper()
            type_value = ""
            if "LIVE" in trailer_text:
                type_value = "Live"
            elif "CPU" in trailer_text:
                type_value = "CPU - Live"
            elif "DROP" in trailer_text:
                type_value = "Drop"

            all_rows.append({
                "source": sheet_name,
                "row_number": row_idx,
                "day": current_day,
                "date": current_date,
                "load_number": load_number,
                "customer": values[1],
                "carrier": values[2],
                "appt_time": normalize_board_time(values[3]),
                "door": values[4],
                "trailer": values[5],
                "status": status,
                "type": type_value,
                "tt4": values[7],
                "loader": values[8],
                "comments": values[9],
                "pulls": parse_number(values[10]),
                "picks": parse_number(values[11]),
                "flags": sorted(set(flags)),
                "raw_row": " | ".join(v for v in values if v),
            })

    return all_rows


def board_records_from_csv(board_file):
    board_file.seek(0)
    df = pd.read_csv(board_file, header=None).fillna("")
    current_day = ""
    current_date = ""
    all_rows = []
    for idx, row in df.iterrows():
        values = [normalize_board_text(v) for v in row.tolist()]
        while len(values) < 13:
            values.append("")
        first_cell = values[0]
        if first_cell in BOARD_DAY_NAMES:
            current_day = first_cell
            current_date = normalize_board_date(values[1])
            continue
        load_number = looks_like_board_load(values[0])
        if not load_number:
            continue
        if detect_trailer_field_late(values[5]):
            status = "Late"
        else:
            status = detect_board_status(values[6])
        if not status:
            status = detect_board_status(" ".join(values))
        trailer_text = values[5].upper()
        type_value = "Live" if "LIVE" in trailer_text else ("CPU - Live" if "CPU" in trailer_text else ("Drop" if "DROP" in trailer_text else ""))
        all_rows.append({
            "source": "CSV Board",
            "day": current_day,
            "date": current_date,
            "load_number": load_number,
            "customer": values[1],
            "carrier": values[2],
            "appt_time": normalize_board_time(values[3]),
            "door": values[4],
            "trailer": values[5],
            "status": status,
            "type": type_value,
            "tt4": values[7],
            "loader": values[8],
            "comments": values[9],
            "pulls": parse_number(values[10]),
            "picks": parse_number(values[11]),
            "flags": [],
            "raw_row": " | ".join(v for v in values if v),
        })
    return all_rows


def board_records_from_inbound_sheet(board_file):
    """
    Read the Inbound sheet from the board Excel file.
    Layout (1-indexed columns):
      1: Load Number   2: Carrier   3: Dispatch Time   4: Type (Live/Drop)
      5: Trailer #     6: Status    7: Receiver         8: From (origin plant)
      9: OR#           10: Notes    11: Start
    Day header rows: col1 = day name, col2 = date, col3 = expected count
    """
    board_file.seek(0)
    try:
        wb = load_workbook(board_file, data_only=True)
    except Exception:
        return []

    inbound_sheet = None
    for candidate in ["Inbound", "inbound", "INBOUND"]:
        if candidate in wb.sheetnames:
            inbound_sheet = candidate
            break
    if not inbound_sheet:
        return []

    ws = wb[inbound_sheet]
    all_rows = []
    current_day = ""
    current_date = ""

    for row_idx in range(1, ws.max_row + 1):
        has_content = any(ws.cell(row_idx, c).value is not None for c in range(1, 12))
        if not has_content:
            continue

        col1 = normalize_board_text(ws.cell(row_idx, 1).value)

        # Day header row — check case-insensitively (board may use SUNDAY vs Sunday)
        col1_title = col1.strip().title()
        if col1_title in BOARD_DAY_NAMES:
            current_day = col1_title
            current_date = normalize_board_date(ws.cell(row_idx, 2).value)
            continue

        # Skip column header row
        if col1.lower() in ("load number", "load #", "load"):
            continue

        # Must be a load number
        load_number = looks_like_board_load(col1)
        if not load_number:
            continue

        all_rows.append({
            "source": inbound_sheet,
            "day": current_day,
            "date": current_date,
            "load_number": load_number,
            "carrier": normalize_board_text(ws.cell(row_idx, 2).value),
            "appt_time": normalize_board_time(ws.cell(row_idx, 3).value),
            "type": normalize_board_text(ws.cell(row_idx, 4).value),
            "trailer": normalize_board_text(ws.cell(row_idx, 5).value),
            "status": normalize_board_text(ws.cell(row_idx, 6).value),
            "receiver": normalize_board_text(ws.cell(row_idx, 7).value),
            "origin": normalize_board_text(ws.cell(row_idx, 8).value),
            "or_number": normalize_board_text(ws.cell(row_idx, 9).value),
            "notes": normalize_board_text(ws.cell(row_idx, 10).value),
        })

    return all_rows


def build_python_inbound_summary(inbound_rows):
    summary = {
        "loads_read_from_inbound": len(inbound_rows),
        "loads_by_day": {},
        "live_loads": 0,
        "drop_loads": 0,
        "on_lot": 0,
        "at_door": 0,
        "loads_with_receiver": 0,
        "loads_missing_receiver": 0,
    }
    for row in inbound_rows:
        day_key = row.get("day") or "Unknown Day"
        summary["loads_by_day"][day_key] = summary["loads_by_day"].get(day_key, 0) + 1

        type_upper = row.get("type", "").upper()
        status_upper = row.get("status", "").upper()

        if "LIVE" in type_upper:
            summary["live_loads"] += 1
        if "DROP" in type_upper:
            summary["drop_loads"] += 1
        if "ON LOT" in status_upper:
            summary["on_lot"] += 1
        if "DOOR" in status_upper:
            summary["at_door"] += 1
        if row.get("receiver"):
            summary["loads_with_receiver"] += 1
        else:
            summary["loads_missing_receiver"] += 1

    return summary


def build_python_board_summary(board_rows):
    summary = {
        "loads_read_from_board": len(board_rows),
        "loads_by_day": {},
        "loads_by_date": {},
        "status_counts": {},
        "late_loads": 0,
        "rtl_loads": 0,
        "rs_loads": 0,
        "picking_loads": 0,
        "picking_short_loads": 0,
        "loaded_short_loads": 0,
        "completed_loads": 0,
        "blank_or_not_started_loads": 0,
        "live_loads": 0,
        "drop_loads": 0,
        "cpu_loads": 0,
        "tt4_needed_loads": 0,
        "load_check_loads": 0,
        "canadian_loads": 0,
        "loads_with_loader_assigned": 0,
        "loads_missing_loader": 0,
        "late_load_details": [],
        "rs_load_details": [],
        "picking_short_details": [],
        "loaded_short_details": [],
        "rtl_details": [],
        "blank_or_not_started_details": [],
        "priority_load_details": [],
    }

    for row in board_rows:
        day_key = row.get("day") or "Unknown Day"
        date_key = row.get("date") or "Unknown Date"
        status = row.get("status") or "Blank/Not Started"
        status_upper = status.upper()
        raw_upper = row.get("raw_row", "").upper()
        flags = row.get("flags", [])

        summary["loads_by_day"][day_key] = summary["loads_by_day"].get(day_key, 0) + 1
        summary["loads_by_date"][date_key] = summary["loads_by_date"].get(date_key, 0) + 1
        summary["status_counts"][status] = summary["status_counts"].get(status, 0) + 1

        if "LATE" in status_upper or "LATE " in f" {raw_upper} ":
            summary["late_loads"] += 1
            summary["late_load_details"].append(row)
        if status_upper == "RTL" or "READY TO LOAD" in status_upper:
            summary["rtl_loads"] += 1
            summary["rtl_details"].append(row)
        if status_upper in ["R/S", "READY/SHORT"] or "R/S" in raw_upper:
            summary["rs_loads"] += 1
            summary["rs_load_details"].append(row)
        if status_upper == "PICKING":
            summary["picking_loads"] += 1
        if "PICKING/SHORT" in status_upper or "PICKING SHORT" in status_upper:
            summary["picking_short_loads"] += 1
            summary["picking_short_details"].append(row)
        if "LOADED SHORT" in status_upper:
            summary["loaded_short_loads"] += 1
            summary["loaded_short_details"].append(row)
        if "COMPLETED" in status_upper or status_upper == "COMPLETE":
            summary["completed_loads"] += 1
        if not row.get("status"):
            summary["blank_or_not_started_loads"] += 1
            summary["blank_or_not_started_details"].append(row)
        if "LIVE" in raw_upper:
            summary["live_loads"] += 1
            summary["priority_load_details"].append(row)
        if "DROP" in raw_upper:
            summary["drop_loads"] += 1
        if "CPU" in raw_upper:
            summary["cpu_loads"] += 1
            summary["priority_load_details"].append(row)
        if "TT4-NEEDED" in flags:
            summary["tt4_needed_loads"] += 1
            summary["priority_load_details"].append(row)
        if "LOAD-CHECK" in flags:
            summary["load_check_loads"] += 1
            summary["priority_load_details"].append(row)
        if "CANADIAN" in flags:
            summary["canadian_loads"] += 1
            summary["priority_load_details"].append(row)
        if row.get("loader"):
            summary["loads_with_loader_assigned"] += 1
        else:
            summary["loads_missing_loader"] += 1

    seen = set()
    unique_priority = []
    for item in summary["priority_load_details"]:
        key = (item.get("load_number"), item.get("row_number"), item.get("source"))
        if key not in seen:
            seen.add(key)
            unique_priority.append(item)
    summary["priority_load_details"] = unique_priority

    return summary


def compact_board_rows_for_ai(board_rows):
    """Strip raw_row and row_number — send only what the AI needs."""
    compact_rows = []
    for row in board_rows:
        compact_rows.append({
            "day": row.get("day", ""),
            "date": row.get("date", ""),
            "load": row.get("load_number", ""),
            "customer": row.get("customer", ""),
            "carrier": row.get("carrier", ""),
            "time": row.get("appt_time", ""),
            "door": row.get("door", ""),
            "trailer": row.get("trailer", ""),
            "status": row.get("status", ""),
            "type": row.get("type", ""),
            "tt4": row.get("tt4", ""),
            "loader": row.get("loader", ""),
            "picks": row.get("picks", 0),
            "pulls": row.get("pulls", 0),
            "flags": row.get("flags", []),
            "comments": row.get("comments", ""),
        })
    return compact_rows


def read_board_file_to_text(board_file):
    """
    Main entry point: reads outbound and inbound sheets, builds Python-verified
    summaries, and returns a JSON string for the AI prompt.
    """
    board_file.seek(0)
    file_name = board_file.name.lower()

    try:
        if file_name.endswith(".csv"):
            board_rows = board_records_from_csv(board_file)
            inbound_rows = []
        else:
            board_rows = board_records_from_excel(board_file)
            board_file.seek(0)
            inbound_rows = board_records_from_inbound_sheet(board_file)

        board_summary = build_python_board_summary(board_rows)
        inbound_summary = build_python_inbound_summary(inbound_rows)
        compact_rows = compact_board_rows_for_ai(board_rows)

        payload = {
            "python_verified_outbound_summary": board_summary,
            "python_verified_inbound_summary": inbound_summary,
            "structured_outbound_rows": compact_rows,
            "structured_inbound_rows": inbound_rows,
            "instructions_for_ai": [
                "Use python_verified_outbound_summary for outbound counts.",
                "Use python_verified_inbound_summary for inbound counts.",
                "Outbound and inbound are separate — never mix their counts.",
                "All times use 24-hour clock.",
                "Blank status means load not yet started.",
                "Flags: LOAD-CHECK=yellow fill, TT4-NEEDED=blue fill, CANADIAN=red font.",
            ],
        }

        return json.dumps(payload, indent=2, ensure_ascii=False)

    except Exception as e:
        error_message = str(e)
        st.error(f"BOARD PARSER ERROR: {error_message}")
        st.exception(e)
        return json.dumps({
            "error": f"Could not read board file: {error_message}",
            "python_verified_outbound_summary": {},
            "python_verified_inbound_summary": {},
            "structured_outbound_rows": [],
            "structured_inbound_rows": [],
        }, indent=2, ensure_ascii=False)


def analyze_board_with_groq(
    board_text, day, shift, total_cases, hours_remaining, total_outbound_loads,
    crossroads_open, deer_creek_open, msb_open, needed, summary_table,
    cases_to_pick, inbound_pallets, notes, oc_alert_text=None,
):
    client = get_groq_client()
    if client is None:
        return (
            "Board analysis could not be completed because GROQ_API_KEY is missing. "
            "Add GROQ_API_KEY in Streamlit Cloud Secrets."
        )

    staffing_lines = []
    for task, row in summary_table.iterrows():
        staffing_lines.append(
            f" {task}: Need {int(row['Needed'])}, Have {int(row['Assigned'])}, "
            f"Gap {int(row['Difference'])} ({row['Status']})"
        )
    staffing_summary = "\n".join(staffing_lines)

    plants_open = [
        p for p, status in [("Crossroads", crossroads_open), ("Deer Creek", deer_creek_open), ("MSB", msb_open)]
        if status == "YES"
    ]

    oc_section = f"\n\n{oc_alert_text}\n" if oc_alert_text else ""

    base_context = f"""
You are an experienced warehouse operations shift manager analyzing an outbound load board that was read directly from an Excel file (cell values, not a screenshot or image). All data is clean and structured — treat every field as accurate cell content.
Use short bullet points. don't over explain.

When reading: separate loads and their data by day, focus on today but still mention when there are still loads on the board from days before, from what day and what is happening with them.

Additional warehouse operation context:
This is a high-volume outbound grocery distribution center operation. This is the first shift and it starts from 6 am to 4:30 pm with 9.5 workable hours. Setting up the second shift for success can vary, but if my morning shift has all loads RTL and the appointments are until 4pm that is still success, not behind. 
The outbound board represents live warehouse execution, not future planning. The board uses 24 hour clock instead of 12.

The manager using this system is focused on:
- Preventing shorts
- Keeping pickers productive
- Avoiding late departures
- Protecting dock flow
- Prioritizing live loads correctly
- Reducing congestion
- Getting ahead instead of reacting late

Operational priorities from highest to lowest:
1. Prevent shorts on customer orders
2. Protect outbound departures
3. Maintain picking flow
4. Prevent inbound congestion
5. Use extra labor proactively

Operational definitions:
- RTL = Ready To Load: Product is staged and ready. Loader can execute.
- R/S = Ready/Short: Load is mostly ready but missing full pallets or replenishment inventory. This is a major operational risk and can quickly become late.
- Picking = Order currently being picked.
- Picking/Short = Picking in progress but inventory shortages are occurring. This usually means replenishment or manufacturing support is needed.
- Loaded Short = Trailer loaded but missing product. This is a severe service risk.
- Live = Trailer physically waiting at the dock. Live loads always have higher priority than drop trailers.
- Drop = Trailer can wait longer and has lower urgency.
- Late = Appointment time already missed or at risk.

Important labor behavior rules:
- Pickers should stay picking whenever possible.
- Tasking/replenishment exists mainly to protect pickers from running out of product.
- If replenishment falls behind, pickers stop producing.
- Loading labor should only be pulled if outbound risk is low.
- Receiving and unloading can temporarily absorb delays better than picking.
- Lead/Extra labor should be used proactively before the operation falls behind.

Operational productivity assumptions:
- 1 picker averages 185 cases/hour
- 1 loader averages 1 trailer/hour
- 1 unloader averages 44 pallets/hour
- 1 replenishment/tasking worker averages 25 pallet moves/hour

Risk interpretation rules:
- Multiple Picking/Short loads means replenishment is failing.
- Multiple R/S loads means outbound may miss appointments.
- Late live loads are highest priority.
- Loads with no door, no trailer, or no loader are operational risks.
- If many loads are blank/not started, the operation is behind schedule.
- If outbound workload is heavier than staffing, recommend labor moves immediately.

Management philosophy:
The goal is not only to survive the shift. The goal is to get ahead early enough that later appointments are protected. We only send people to manufacturing if it's going to benefit us.

The manager prefers:
- proactive recommendations
- actionable labor moves
- operational risk analysis
- realistic achievable goals
- time-based recommendations
- practical warehouse language
- direct communication without corporate fluff

When making recommendations:
- Specify EXACTLY where labor should move from and to
- Explain WHY
- Explain operational consequences if no action is taken
- Give achievable operational goals for the next 30 minutes and next 2 hours
- Prioritize live loads, shorts, and dock flow
- Think like an experienced outbound operations manager

Today's operational context:
- Day: {day}, Shift: {shift}
- Total cases forecast for today: {total_cases:,}
- Cases to pick this shift: {cases_to_pick:,.0f}
- Hours remaining in shift: {hours_remaining}
- Total outbound loads scheduled today: {total_outbound_loads}
- Inbound pallets expected: {inbound_pallets:,} (Plants open: {", ".join(plants_open) if plants_open else "None"})
- Manager notes: {notes if notes.strip() else "None"}

Current staffing vs. what we need:
{staffing_summary}
{oc_section}
Board data rules and operation rules:
- All data below was extracted directly from Excel cells — treat it as accurate.
- Cells annotated with [LOAD-CHECK] had a yellow fill in Excel, meaning that load needs a load check.
- Cells annotated with [CANADIAN] had red font in Excel, meaning it is a Canadian load.
- If a color annotation is absent, the cell had no special flag — do not guess.
- Blank status on the board means the load is not currently being worked.
- R/S means Ready to load but still short on full pallets.
- Our average productivity:
  - Picking: 185 cases per hour per worker allocated
  - Loading: 1 trailer per hour per worker allocated
  - Unloading: 44 pallets per hour per worker allocated
  - Full pallets / replenishment movement: 25 full pallets per hour per worker allocated
- Picking is measured in tickets on the board, but analyze everything in cases. Our average is 60 cases per picking ticket.
- If a column or value is unclear or missing, say "unclear" — do not invent information.

Here is the outbound board data extracted directly from the Excel file:
{board_text}
"""

    output_structure = """
Read the board carefully row by row.

Give me a clear, practical warehouse manager analysis in plain English covering:

1. Board Summary:
- Break loads down by status and day: RTL, R/S, Late, Picking, Picking/Short, Loaded Short, Completed, blank/not started, etc.
- Specify how many loads are completed today out of the total for the day.
- Specify any late loads, from when, if they are occupying a door, and which door.

2.  Opportunity Customer (OC) Alerts:
- List every load on the board that belongs to a customer on the Opportunity Customer List.
- For each OC load: state the load number, customer name, current status, appointment time, and EXACTLY what special actions are required before this load ships.
- If pictures are required, state when they should be taken and who should own it.
- If supervisor sign-off is required, state who should sign off and when.
- If no OC customers are on the board today, state clearly: "No Opportunity Customers detected on today's board."

3. Picking & Short Risk:
- How many loads have not been started?
- Given cases-to-pick and current staffing, are we at risk of falling further behind? In easy words, yes or no and why.
- How big is the risk? Explain what are the risk factors.
- Can we get ahead? Yes or no and why
- Given all this information, how far ahead can we finish this shift?
- Give me the load appointment times we should be picking by the end of this shift.
- Specify people from what areas we can move from and to where. Should we consider sending people to manufacturing to reduce short risks? Specify people from what areas we can move staff from and to where.

4. Prioritization:
- Are there any loads we should prioritize? Be specific, add load numbers.
- How and why should we prioritize them?

5. Cross-Analysis with Staffing:
- Given staffing gaps or surpluses, which problems can we actually fix right now?
- Where should labor move first?
- Based on staffing and demand, what should be an achievable goal for this shift?
- How ahead or behind should we finish this shift?

6. Top 3 Action Items:
- What are the 3 most important things the manager should do in the next 30 minutes?
- What are the 3 most important things the manager should do in the next 2 hours to achieve today's goal?

Make sure every recommendation and suggested action is achievable and following the same direction.
Have somewhere where you clearly set the expectations for the shift and explain why. I want this easy to identify.
Keep the tone like a smart, experienced ops manager talking to another manager.
No corporate fluff.
Be clear, practical, and actionable.
Add times and case/pallet numbers to every goal so progress is measurable.
Include what-if scenarios: if X happens, here is what to do.
Only use data, do not guess.
Talk about how you are heading the second shift for success.
When suggesting to think about moving staff specify from where to where.
Remember even though we have loads for the day it is separated in 2 shifts. We load approximately 52% of loads in the first shift. Take that into consideration, we still can have the loads ready to load for second shift. Read the board and check the times.
When making suggestions that we should be ready to load up to a specific hour do not use a range, be specific.
OC loads (Opportunity Customers) must ALWAYS be called out explicitly and early in the analysis — never buried at the bottom.
"""

    try:
        initial_response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": base_context + output_structure}],
            temperature=0.2,
            max_completion_tokens=2500,
        )
        initial_analysis = initial_response.choices[0].message.content

        validation_response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are a meticulous warehouse operations auditor. "
                        "Your only job is to validate the operational analysis below against the raw board data and operational context provided. "
                        "Check for: incorrect load counts by status, wrong load numbers referenced, "
                        "contradictory statements (e.g. saying a load is RTL and also Picking), "
                        "math errors in labor or case projections, recommendations that conflict with stated priorities, "
                        "and any invented data not present in the board. "
                        "ALSO verify: if an Opportunity Customer (OC) was flagged in the context, confirm it was addressed "
                        "explicitly in the analysis with correct requirements. If it was missed, flag it. "
                        "Be specific about each issue found. If something is correct, confirm it. "
                        "Do not rewrite the full analysis — only list what needs to be corrected and what is confirmed accurate. "
                        "Keep it concise and factual."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"=== ORIGINAL BOARD DATA AND CONTEXT ===\n{base_context}\n\n"
                        f"=== INITIAL ANALYSIS TO VALIDATE ===\n{initial_analysis}"
                    ),
                },
            ],
            temperature=0.1,
            max_completion_tokens=1200,
        )
        validation_notes = validation_response.choices[0].message.content

        final_response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "system",
                    "content": (
                        "You are an experienced warehouse operations shift manager. "
                        "You have an initial operational analysis and a validation audit that flags any errors or confirms accuracy. "
                        "Your job is to produce the final, corrected, clean analysis. "
                        "Apply every correction flagged in the validation. Keep everything that was confirmed accurate. "
                        "Do not mention the validation process or the word 'corrected' — just write the final clean analysis "
                        "as if you are delivering it directly to the shift manager. "
                        "Follow the exact same output structure as the initial analysis. "
                        "Opportunity Customer (OC) alerts must appear early and be complete — never omit or shorten them."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"=== INITIAL ANALYSIS ===\n{initial_analysis}\n\n"
                        f"=== VALIDATION NOTES (apply these corrections) ===\n{validation_notes}\n\n"
                        f"=== OUTPUT STRUCTURE TO FOLLOW ===\n{output_structure}"
                    ),
                },
            ],
            temperature=0.2,
            max_completion_tokens=2500,
        )
        return final_response.choices[0].message.content

    except Exception as e:
        return f"Board analysis could not be completed: {str(e)}"


def write_board_analysis_to_excel(wb, analysis_text, oc_matches=None):
    sheet_name = "Board Analysis"
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(sheet_name)

    dark_blue = "0F5B78"
    orange = "C55A11"
    white = "FFFFFF"
    light_blue = "D9EAF7"
    light_orange = "FCE4D6"
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "Board Excel Analysis — AI Insights"
    ws["A1"].font = Font(size=16, bold=True, color=white)
    ws["A1"].fill = PatternFill("solid", fgColor=dark_blue)
    ws["A1"].alignment = Alignment(horizontal="center")
    ws.merge_cells("A1:G1")
    ws.row_dimensions[1].height = 28

    ws["A2"] = "Generated by Groq AI — cross-referenced with today's staffing and demand data"
    ws["A2"].font = Font(italic=True, size=10)
    ws["A2"].fill = PatternFill("solid", fgColor=light_blue)
    ws.merge_cells("A2:G2")

    current_row = 4

    if oc_matches:
        ws.cell(current_row, 1).value = "OPPORTUNITY CUSTOMER ALERT — SPECIAL HANDLING REQUIRED"
        ws.cell(current_row, 1).font = Font(size=13, bold=True, color=white)
        ws.cell(current_row, 1).fill = PatternFill("solid", fgColor=orange)
        ws.cell(current_row, 1).alignment = Alignment(horizontal="center")
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        ws.row_dimensions[current_row].height = 22
        current_row += 1

        for match in oc_matches:
            c = match["customer"]
            oc_lines = [
                f"CUSTOMER: {c['name'].upper()}  |  Priority: {c['priority']}",
                f"Issue History: {c['issue']}",
                f"DC Requirements: {c['requirements']}",
            ]
            if c["sign_off"]:
                oc_lines.append("DC Supervisor Sign-Off REQUIRED before this load ships.")
            if c["pictures"]:
                oc_lines.append("Photos REQUIRED: 3 on dock + 3 during loading (6 total). Email to manager.")
            for line in oc_lines:
                cell = ws.cell(current_row, 1, line)
                cell.font = Font(size=10, bold=("CUSTOMER:" in line or "" in line or "" in line))
                cell.fill = PatternFill("solid", fgColor=light_orange)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = border
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                ws.row_dimensions[current_row].height = max(15, min(60, len(line) // 5))
                current_row += 1
            current_row += 1
        current_row += 1

    for line in analysis_text.split("\n"):
        cell = ws.cell(current_row, 1, line)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = border
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
        ws.row_dimensions[current_row].height = max(15, min(60, len(line) // 5))
        current_row += 1

    for col in range(1, 8):
        ws.column_dimensions[get_column_letter(col)].width = 22
    ws.column_dimensions["A"].width = 110


def write_recommendations_to_excel(wb, staff):
    ws_staff = wb["Staffing sheet 1ST Shift"]
    ws_crew = wb["Crew Sheet"]

    for excel_row, task in zip(range(2, len(staff) + 2), staff["Recommended Task"]):
        ws_staff[f"I{excel_row}"] = task

    crew_name_to_row = {}
    for row in range(2, ws_crew.max_row + 1):
        name = ws_crew[f"A{row}"].value
        if name:
            crew_name_to_row[str(name).strip().lower()] = row

    for _, row in staff.iterrows():
        name = str(row["Name"]).strip().lower()
        task = row["Recommended Task"]
        if name in crew_name_to_row:
            crew_row = crew_name_to_row[name]
            ws_crew[f"C{crew_row}"] = task
            ws_crew[f"D{crew_row}"] = task


def build_dashboard(wb, summary_table, present_recommendations, recommendations, oc_matches=None):
    if "Staffing Dashboard" in wb.sheetnames:
        ws_dash = wb["Staffing Dashboard"]
        ws_dash.delete_rows(1, ws_dash.max_row)
    else:
        ws_dash = wb.create_sheet("Staffing Dashboard")

    dark_blue = "0F5B78"
    orange = "C55A11"
    light_blue = "D9EAF7"
    green = "C6EFCE"
    red = "FFC7CE"
    yellow = "FFEB9C"
    white = "FFFFFF"
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws_dash["A1"] = "1st Shift Staffing Dashboard"
    ws_dash["A1"].font = Font(size=18, bold=True, color=white)
    ws_dash["A1"].fill = PatternFill("solid", fgColor=dark_blue)
    ws_dash["A1"].alignment = Alignment(horizontal="center")
    ws_dash.merge_cells("A1:K1")

    total_present = len(present_recommendations)
    total_needed = int(summary_table["Needed"].sum())
    total_assigned = int(summary_table["Assigned"].sum())
    lead_extra = int((present_recommendations["Recommended Task"] == "Lead/Extra").sum())
    overall_gap = total_assigned - total_needed

    kpis = [
        ("Total Present", total_present),
        ("Total Needed", total_needed),
        ("Total Assigned", total_assigned),
        ("Lead/Extra", lead_extra),
        ("Overall Gap", overall_gap),
    ]
    kpi_cols = [1, 3, 5, 7, 9]
    for (label, value), col in zip(kpis, kpi_cols):
        ws_dash.cell(3, col).value = label
        ws_dash.cell(4, col).value = value
        ws_dash.cell(3, col).font = Font(bold=True, color=white)
        ws_dash.cell(3, col).fill = PatternFill("solid", fgColor=dark_blue)
        ws_dash.cell(3, col).alignment = Alignment(horizontal="center")
        ws_dash.cell(4, col).font = Font(bold=True, size=14)
        ws_dash.cell(4, col).fill = PatternFill("solid", fgColor=light_blue)
        ws_dash.cell(4, col).alignment = Alignment(horizontal="center")
        ws_dash.merge_cells(start_row=3, start_column=col, end_row=3, end_column=col + 1)
        ws_dash.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)

    oc_banner_row = 6
    if oc_matches:
        customer_names = ", ".join(m["customer"]["name"].upper() for m in oc_matches)
        ws_dash.cell(oc_banner_row, 1).value = (
            f"OC ALERT: Opportunity Customers on today's board — {customer_names} — See 'Board Analysis' tab for full requirements."
        )
        ws_dash.cell(oc_banner_row, 1).font = Font(bold=True, color=white, size=11)
        ws_dash.cell(oc_banner_row, 1).fill = PatternFill("solid", fgColor=orange)
        ws_dash.cell(oc_banner_row, 1).alignment = Alignment(horizontal="center", wrap_text=True)
        ws_dash.merge_cells(start_row=oc_banner_row, start_column=1, end_row=oc_banner_row, end_column=11)
        ws_dash.row_dimensions[oc_banner_row].height = 22
        summary_label_row = oc_banner_row + 2
    else:
        summary_label_row = oc_banner_row

    ws_dash.cell(summary_label_row, 1).value = "Needed vs Assigned"
    ws_dash.cell(summary_label_row, 1).font = Font(size=14, bold=True)

    header_row = summary_label_row + 1
    headers = ["Task", "Needed", "Assigned", "Difference", "Status"]
    for c, header in enumerate(headers, 1):
        cell = ws_dash.cell(header_row, c)
        cell.value = header
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for r, (task, row) in enumerate(summary_table.iterrows(), header_row + 1):
        values = [task, int(row["Needed"]), int(row["Assigned"]), int(row["Difference"]), row["Status"]]
        for c, value in enumerate(values, 1):
            cell = ws_dash.cell(r, c)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if c == 5:
                if value == "Good":
                    cell.fill = PatternFill("solid", fgColor=green)
                elif value == "Understaffed":
                    cell.fill = PatternFill("solid", fgColor=red)
                else:
                    cell.fill = PatternFill("solid", fgColor=yellow)

    ws_dash.cell(summary_label_row, 7).value = "Written Recommendations / What-Ifs"
    ws_dash.cell(summary_label_row, 7).font = Font(size=14, bold=True)

    rec_row = header_row
    for rec in recommendations:
        ws_dash.cell(rec_row, 7).value = f"• {rec}"
        ws_dash.cell(rec_row, 7).alignment = Alignment(wrap_text=True, vertical="top")
        ws_dash.merge_cells(start_row=rec_row, start_column=7, end_row=rec_row, end_column=11)
        rec_row += 1

    board_start = max(header_row + len(summary_table) + 4, rec_row + 2)
    ws_dash.cell(board_start, 1).value = "Recommended Staffing Board"
    ws_dash.cell(board_start, 1).font = Font(size=14, bold=True)

    board_headers = ["Name", "Skills", "Best Fit", "Recommended Task"]
    for c, header in enumerate(board_headers, 1):
        cell = ws_dash.cell(board_start + 1, c)
        cell.value = header
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for r, (_, row) in enumerate(present_recommendations.iterrows(), board_start + 2):
        values = [row["Name"], row["Skills"], row["Best Fit"], row["Recommended Task"]]
        for c, value in enumerate(values, 1):
            cell = ws_dash.cell(r, c)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(horizontal="center")
            if r % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=light_blue)

    chart_anchor_row = board_start + len(present_recommendations) + 5

    bar = BarChart()
    bar.title = "Needed vs Assigned"
    bar.y_axis.title = "Workers"
    bar.x_axis.title = "Task"
    data = Reference(ws_dash, min_col=2, max_col=3, min_row=header_row, max_row=header_row + len(summary_table))
    cats = Reference(ws_dash, min_col=1, min_row=header_row + 1, max_row=header_row + len(summary_table))
    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 9
    bar.width = 15
    bar.legend.position = "r"
    ws_dash.add_chart(bar, f"E{chart_anchor_row}")

    pie = PieChart()
    pie.title = "Assigned Labor Distribution"
    pie_data = Reference(ws_dash, min_col=3, min_row=header_row, max_row=header_row + len(summary_table))
    pie_cats = Reference(ws_dash, min_col=1, min_row=header_row + 1, max_row=header_row + len(summary_table))
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 9
    pie.width = 13
    pie.legend.position = "r"
    ws_dash.add_chart(pie, f"I{chart_anchor_row}")

    for col in range(1, 12):
        ws_dash.column_dimensions[get_column_letter(col)].width = 18
    ws_dash.column_dimensions["A"].width = 22
    for col in ["G", "H", "I", "J", "K"]:
        ws_dash.column_dimensions[col].width = 35
    ws_dash.freeze_panes = f"A{header_row}"


def build_email_draft(
    day, shift, total_cases, hours_remaining, total_outbound_loads_day,
    summary_table, present_recommendations, recommendations,
    board_analysis_text=None, oc_matches=None,
):
    total_present = len(present_recommendations)
    total_needed = int(summary_table["Needed"].sum())
    total_assigned = int(summary_table["Assigned"].sum())
    overall_gap = total_assigned - total_needed
    subject = f"{day} {shift} Shift Staffing Report"

    staffing_lines = []
    for task, row in summary_table.iterrows():
        staffing_lines.append(
            f"- {task}: Need {int(row['Needed'])}, Assigned {int(row['Assigned'])}, "
            f"Gap {int(row['Difference'])} ({row['Status']})"
        )
    top_recommendations = "\n".join([f"- {rec}" for rec in recommendations[:8]])

    oc_email_block = ""
    if oc_matches:
        oc_lines = ["\n OPPORTUNITY CUSTOMER ALERT:"]
        for match in oc_matches:
            c = match["customer"]
            oc_lines.append(f" - {c['name'].upper()} [{c['priority']}]: {c['requirements']}")
            if c["sign_off"]:
                oc_lines.append(" → Supervisor sign-off REQUIRED before shipping.")
            if c["pictures"]:
                oc_lines.append(" → 6 photos required (3 on dock, 3 loading). Email to manager.")
        oc_email_block = "\n".join(oc_lines)

    body = f"""
Good morning,

Here is the staffing report for {day} {shift} shift.

Daily Inputs:
- Total cases: {total_cases:,}
- Total outbound loads: {total_outbound_loads_day}
- Hours remaining: {hours_remaining}
- Total present: {total_present}
- Total needed: {total_needed}
- Total assigned: {total_assigned}
- Overall labor gap: {overall_gap}

Staffing Summary:
{chr(10).join(staffing_lines)}
{oc_email_block}

Key Recommendations / What-Ifs:
{top_recommendations}
"""

    if board_analysis_text:
        body += f"""

Board Analysis:
{board_analysis_text}
"""
    body += """

The full staffing report is attached.

Thanks,
"""
    return subject, body.strip()


#  STREAMLIT INTERFACE 

st.sidebar.header("Daily Inputs")

day = st.sidebar.selectbox(
    "Day",
    ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
)

shift = st.sidebar.selectbox("Shift", ["1st", "2nd"])

total_cases = st.sidebar.number_input("Total Cases for Today", min_value=0, step=1, value=0)

hours_remaining = st.sidebar.number_input("Hours Remaining in Shift", min_value=0.0, step=0.25, value=8.0)

total_outbound_loads_day = st.sidebar.number_input("Total Outbound Loads for the Day", min_value=0, step=1, value=0)

crossroads_open = st.sidebar.selectbox("Crossroads plant open?", ["YES", "NO"])
deer_creek_open = st.sidebar.selectbox("Deer Creek plant open?", ["YES", "NO"])
msb_open = st.sidebar.selectbox("MSB plant open?", ["YES", "NO"])

present_workers = st.sidebar.multiselect("Who is present?", names)

notes = st.sidebar.text_area("Operations Notes")

st.markdown("---")
st.subheader("Outbound Board Excel / CSV")

board_file = st.file_uploader(
    "Upload the outbound load board Excel or CSV file",
    type=["xlsx", "xls", "csv"],
    help="Cell values and color flags (yellow = load check, light-blue = TT4, red font = Canadian) are read directly from the file.",
)

if board_file:
    st.success("Board file loaded — ready for analysis.")

    with st.expander("Preview: What Python parsed from the board (no AI tokens used)", expanded=False):
        try:
            board_file.seek(0)
            file_name_lower = board_file.name.lower()
            if file_name_lower.endswith(".csv"):
                preview_rows = board_records_from_csv(board_file)
            else:
                preview_rows = board_records_from_excel(board_file)

            if not preview_rows:
                st.warning("No load rows were parsed. Check that the file has day headers (e.g. 'Monday') and 5-9 digit load numbers in column A.")
            else:
                #  Staff counts (live from sidebar selection) 
                total_staff_present = len(present_workers)

                st.metric("Staff Present Today", total_staff_present)

                st.markdown("---")

                #  Board summary counts 
                preview_summary = build_python_board_summary(preview_rows)
                total = preview_summary["loads_read_from_board"]

                col1, col2, col3, col4, col5 = st.columns(5)
                col1.metric("Total Loads", total)
                col2.metric("RTL", preview_summary["rtl_loads"])
                col3.metric("Picking/Short", preview_summary["picking_short_loads"])
                col4.metric("R/S", preview_summary["rs_loads"])
                col5.metric("Loaded Short", preview_summary["loaded_short_loads"])

                col6, col7, col8, col9, col10 = st.columns(5)
                col6.metric("Picking", preview_summary["picking_loads"])
                col7.metric("Blank/Not Started", preview_summary["blank_or_not_started_loads"])
                col8.metric("Live Loads", preview_summary["live_loads"])
                col9.metric("CPU Loads", preview_summary["cpu_loads"])
                col10.metric("Late", preview_summary["late_loads"])

                st.caption(f"Outbound loads by day: {preview_summary['loads_by_day']}")

                # ── Inbound summary ───────────────────────────────────────────
                board_file.seek(0)
                inbound_preview_rows = board_records_from_inbound_sheet(board_file)
                if inbound_preview_rows:
                    inbound_preview_summary = build_python_inbound_summary(inbound_preview_rows)
                    st.markdown("---")
                    st.markdown("**Inbound**")
                    ib1, ib2, ib3, ib4 = st.columns(4)
                    ib1.metric("Total Inbound", inbound_preview_summary["loads_read_from_inbound"])
                    ib2.metric("Live", inbound_preview_summary["live_loads"])
                    ib3.metric("Drop", inbound_preview_summary["drop_loads"])
                    ib4.metric("On Lot / At Door", inbound_preview_summary["on_lot"] + inbound_preview_summary["at_door"])
                    st.caption(f"Inbound loads by day: {inbound_preview_summary['loads_by_day']}")
                    inbound_df = pd.DataFrame([
                        {
                            "Day": r.get("day", ""),
                            "Load #": r.get("load_number", ""),
                            "Carrier": r.get("carrier", ""),
                            "Time": r.get("appt_time", ""),
                            "Type": r.get("type", ""),
                            "Trailer": r.get("trailer", ""),
                            "Status": r.get("status", ""),
                            "Receiver": r.get("receiver", ""),
                            "Origin": r.get("origin", ""),
                            "Notes": r.get("notes", ""),
                        }
                        for r in inbound_preview_rows
                    ])
                    st.dataframe(inbound_df, use_container_width=True, height=250)

                st.markdown("---")
                #  Full outbound parsed table
                st.markdown("**Every outbound load row Python extracted from the file:**")
                preview_df = pd.DataFrame([
                    {
                        "Day": r.get("day", ""),
                        "Date": r.get("date", ""),
                        "Load #": r.get("load_number", ""),
                        "Customer": r.get("customer", ""),
                        "Carrier": r.get("carrier", ""),
                        "Time": r.get("appt_time", ""),
                        "Door": r.get("door", ""),
                        "Trailer": r.get("trailer", ""),
                        "Status": r.get("status", "") or "—",
                        "Type": r.get("type", ""),
                        "TT4": r.get("tt4", ""),
                        "Loader": r.get("loader", ""),
                        "Picks": r.get("picks", 0),
                        "Pulls": r.get("pulls", 0),
                        "Flags": ", ".join(r.get("flags", [])),
                        "Comments": r.get("comments", ""),
                    }
                    for r in preview_rows
                ])
                st.dataframe(preview_df, use_container_width=True, height=400)

                #  Quick sanity checks 
                st.markdown("**Quick sanity checks:**")
                issues = []
                blank_time = [r["load_number"] for r in preview_rows if not r.get("appt_time")]
                if blank_time:
                    issues.append(f" {len(blank_time)} load(s) have no time parsed: {', '.join(blank_time[:5])}{'...' if len(blank_time) > 5 else ''}")
                blank_customer = [r["load_number"] for r in preview_rows if not r.get("customer")]
                if blank_customer:
                    issues.append(f" {len(blank_customer)} load(s) have no customer name: {', '.join(blank_customer[:5])}")
                no_day = [r["load_number"] for r in preview_rows if not r.get("day")]
                if no_day:
                    issues.append(f" {len(no_day)} load(s) have no day context (missing day header row?): {', '.join(no_day[:5])}")
                if issues:
                    for issue in issues:
                        st.warning(issue)
                else:
                    st.success("All loads have time, customer, and day context — parse looks clean.")

        except Exception as e:
            st.error(f"Preview failed: {e}")
            st.exception(e)

with st.expander("View Opportunity Customer List (from Excel file)"):
    oc_list_preview = load_oc_customer_list()
    if oc_list_preview:
        oc_preview_rows = []
        for c in oc_list_preview:
            oc_preview_rows.append({
                "Customer": c["name"].title(),
                "Customer #": c["customer_number"] or "—",
                "Priority": c["priority"],
                "Issue": c["issue"],
                "DC Requirements": c["requirements"],
                "Sign-Off Required": "Yes" if c["sign_off"] else "No",
                "Photos Required": "Yes" if c["pictures"] else "No",
            })
        st.dataframe(pd.DataFrame(oc_preview_rows), use_container_width=True)
        st.caption(f"Loaded {len(oc_list_preview)} customers from '{OC_FILE}'")
    else:
        st.warning(f"No customers loaded. Check that '{OC_FILE}' exists in the app folder.")

st.markdown("---")

if st.button("Generate Staffing Report"):
    working_file = f"working_staffing_file_{day}_{shift}.xlsx"
    shutil.copyfile(TEMPLATE_FILE, working_file)

    wb = load_workbook(working_file)
    ws = wb["Inputs"]

    total_outbound_loads_actual = total_outbound_loads_day * 0.52

    ws["B1"] = day
    ws["B2"] = shift
    ws["B3"] = total_cases
    ws["B4"] = hours_remaining
    ws["B8"] = crossroads_open
    ws["B9"] = deer_creek_open
    ws["B10"] = msb_open

    cases_to_pick, full_pallets = calculate_input_values(day, shift, total_cases)
    ws["B5"] = cases_to_pick
    ws["B6"] = full_pallets
    ws["B7"] = total_outbound_loads_actual

    # Dynamic range — scan all rows that have a name in col E
    _last_name_row = 3
    for _r in range(3, ws.max_row + 1):
        if ws[f"E{_r}"].value and str(ws[f"E{_r}"].value).strip():
            _last_name_row = _r
        elif _r > _last_name_row + 10:
            break
    for row in range(3, _last_name_row + 1):
        ws[f"F{row}"] = ""

    selected = {name.strip().lower() for name in present_workers}
    for row in range(3, _last_name_row + 1):
        worker_name = ws[f"E{row}"].value
        if worker_name and str(worker_name).strip().lower() in selected:
            ws[f"F{row}"] = "x"

    ws["B12"] = notes

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(working_file)

    needed, raw_needed, cases_to_pick, full_pallets, inbound_pallets = calculate_needed(
        day, shift, total_cases, hours_remaining, total_outbound_loads_actual,
        crossroads_open, deer_creek_open, msb_open,
    )

    staff = pd.read_excel(working_file, sheet_name="Staffing sheet 1ST Shift", usecols="A,D,F,I,T")
    staff.columns = ["Name", "Skills", "Best Fit", "Recommended Task", "Present"]
    staff = staff[staff["Name"].notna()].copy()

    selected_names = {name.strip().lower() for name in present_workers}
    staff["Present"] = staff["Name"].astype(str).str.strip().str.lower().apply(
        lambda x: "x" if x in selected_names else ""
    )

    staff = generate_recommendations(staff, needed)
    present_recommendations, summary_table = build_summary(staff, needed)
    recommendations = build_recommendations(summary_table, present_recommendations, raw_needed, hours_remaining, notes)

    wb = load_workbook(working_file)
    write_recommendations_to_excel(wb, staff)

    board_analysis_text = None
    oc_matches = []

    if board_file is not None:
        with st.spinner("Reading board file → scanning for Opportunity Customers → running analysis → validating → finalizing..."):
            board_text = read_board_file_to_text(board_file)

            oc_matches = find_oc_customers_in_board(board_text)
            oc_alert_text = build_oc_alert_text(oc_matches)

            if oc_matches:
                customer_names_found = [m["customer"]["name"].upper() for m in oc_matches]
                st.warning(
                    f" **Opportunity Customer Alert:** "
                    f"The following customers were detected on today's board and require special handling: "
                    f"**{', '.join(customer_names_found)}**. "
                    f"See the OC Alerts section below for full requirements."
                )

            board_analysis_text = analyze_board_with_groq(
                board_text=board_text,
                day=day,
                shift=shift,
                total_cases=total_cases,
                hours_remaining=hours_remaining,
                total_outbound_loads=total_outbound_loads_day,
                crossroads_open=crossroads_open,
                deer_creek_open=deer_creek_open,
                msb_open=msb_open,
                needed=needed,
                summary_table=summary_table,
                cases_to_pick=cases_to_pick,
                inbound_pallets=inbound_pallets,
                notes=notes,
                oc_alert_text=oc_alert_text,
            )

            write_board_analysis_to_excel(wb, board_analysis_text, oc_matches=oc_matches)

    build_dashboard(wb, summary_table, present_recommendations, recommendations, oc_matches=oc_matches)

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    try:
        os.remove(working_file)
    except Exception:
        pass

    st.success("Staffing report generated successfully.")

    if oc_matches:
        st.markdown("---")
        st.subheader("Opportunity Customer Alerts")
        st.error(
            "The following customers on today's board are on the **Opportunity Customer List** "
            "and require special DC actions before their loads ship."
        )
        for match in oc_matches:
            c = match["customer"]
            with st.expander(f" {c['name'].upper()}  —  Priority: {c['priority']}", expanded=True):
                st.markdown(f"**Issue History:** {c['issue']}")
                st.markdown(f"**DC Requirements:** {c['requirements']}")
                if c["sign_off"]:
                    st.markdown(" **DC Supervisor Sign-Off REQUIRED before this load ships.**")
                if c["pictures"]:
                    st.markdown(" **Photos REQUIRED:** 3 on dock + 3 during loading (6 total). Email to manager.")
    elif board_file is not None:
        st.info("No Opportunity Customers detected on today's board.")

    st.subheader("Staffing Summary")
    st.dataframe(summary_table, use_container_width=True)

    st.subheader("Recommended Staffing Board")
    st.dataframe(
        present_recommendations[["Name", "Skills", "Best Fit", "Recommended Task"]].reset_index(drop=True),
        use_container_width=True,
    )

    st.subheader("Written Recommendations / What-Ifs")
    for rec in recommendations:
        st.write(f"• {rec}")

    if board_analysis_text:
        st.markdown("---")
        st.subheader("Board Excel Analysis — AI Insights")
        st.info(
            "The analysis below was generated by Groq AI reading the board Excel/CSV file directly "
            "from cell values, including color flags for load checks, TT4s, and Canadian loads."
        )
        st.markdown(board_analysis_text)

    st.download_button(
        label="Download Staffing Report",
        data=output,
        file_name="Staffing Report Generated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    email_subject, email_body = build_email_draft(
        day=day,
        shift=shift,
        total_cases=total_cases,
        hours_remaining=hours_remaining,
        total_outbound_loads_day=total_outbound_loads_day,
        summary_table=summary_table,
        present_recommendations=present_recommendations,
        recommendations=recommendations,
        board_analysis_text=board_analysis_text,
        oc_matches=oc_matches,
    )

    st.markdown("---")
    st.subheader("Email Ready to Send")
    st.text_input("Email Subject", value=email_subject)
    st.text_area("Email Body", value=email_body, height=500)
