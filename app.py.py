import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from io import BytesIO
import json
import re
import os
import shutil
from openai import OpenAI


st.set_page_config(page_title="Staffing Report Generator", layout="wide")

st.title("Staffing Report Generator")
st.write("Enter daily inputs, select who is present, and generate the staffing report.")

TEMPLATE_FILE = "staffing_template.xlsx"


if not os.path.exists(TEMPLATE_FILE):
    st.error("Template file not found. Put staffing_template.xlsx in the same folder as report.py.")
    st.stop()


# ── OPPORTUNITY CUSTOMER LIST (Excel-driven) ─────────────────────────────────
# Put this Excel file in the same folder as app.py / report.py in GitHub.
# Expected file name matches the file you provided.
OC_FILE = "Resers DCs Opportunity Cusotmer List.xlsx"


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def yes_no_to_bool(value):
    return clean_text(value).upper() in ["Y", "YES", "TRUE", "1", "X"]


def build_aliases(customer_name):
    """
    Builds practical match terms from the customer name.
    This lets the board match shortened customer names like Target, Sobeys, Jewel, etc.
    """
    name = clean_text(customer_name).lower()

    aliases = []

    if name:
        aliases.append(name)

    cleaned = (
        name.replace(" - all loads", "")
        .replace("(olathe)", "")
        .replace("'", "")
        .replace("’", "")
        .strip()
    )

    if cleaned and cleaned not in aliases:
        aliases.append(cleaned)

    words = cleaned.split()

    if len(words) > 0:
        aliases.append(words[0])

    # Extra common aliases / misspellings
    if "target" in cleaned:
        aliases.append("target")

    if "sysco kc" in cleaned or "sysco" in cleaned:
        aliases += ["sysco", "sysco kansas city", "sysco olathe", "sysco kc olathe"]

    if "sobey" in cleaned:
        aliases += ["sobeys", "sobey", "sobey's"]

    if "pfs" in cleaned:
        aliases += ["pfs", "pfs virginia", "pfs virgina", "pfs va"]

    if "metro toronto" in cleaned:
        aliases += ["metro toronto", "metro fresh", "metro toronto fresh dc"]

    if "jewel" in cleaned:
        aliases += ["jewels", "jewel", "jewel's"]

    if "whataburguer" in cleaned or "whataburger" in cleaned:
        aliases += ["whataburger", "whataburguer"]

    if "awg" in cleaned:
        aliases += ["awg", "associated wholesale grocers"]

    return list(dict.fromkeys([alias for alias in aliases if alias]))


def load_oc_customer_list():
    """
    Reads Opportunity Customers from the Excel file instead of hardcoding them in Python.

    File expected:
    - Resers DCs Opportunity Cusotmer List.xlsx
    - Sheet: OC Customer List
    - Headers on Excel row 6, so pandas header=5

    Expected columns:
    - Customer Name
    - Customer #
    - Customer Profile-Why are they an OC?
    - DC Requirements (Summarized)
    - Sign Off \n(Y/N)
    - Pictures  \n(Y/N)
    """
    if not os.path.exists(OC_FILE):
        st.warning(
            f"OC customer list file not found: {OC_FILE}. "
            "Opportunity Customer detection will be skipped."
        )
        return []

    try:
        df = pd.read_excel(
            OC_FILE,
            sheet_name="OC Customer List",
            header=5
        ).fillna("")

        oc_list = []

        for _, row in df.iterrows():
            customer_name = clean_text(row.get("Customer Name", ""))

            if not customer_name:
                continue

            # Skip example/template rows if present
            if customer_name.lower() in ["market x", "example", "customer name"]:
                continue

            customer_number = clean_text(row.get("Customer #", ""))
            issue = clean_text(row.get("Customer Profile-Why are they an OC?", ""))
            requirements = clean_text(row.get("DC Requirements (Summarized)", ""))

            sign_off = yes_no_to_bool(row.get("Sign Off \n(Y/N)", ""))
            pictures = yes_no_to_bool(row.get("Pictures  \n(Y/N)", ""))

            priority = "HIGH" if sign_off or pictures else "MEDIUM"

            oc_list.append(
                {
                    "name": customer_name.lower(),
                    "aliases": build_aliases(customer_name),
                    "customer_number": customer_number,
                    "issue": issue,
                    "requirements": requirements,
                    "sign_off": sign_off,
                    "pictures": pictures,
                    "priority": priority,
                }
            )

        return oc_list

    except Exception as e:
        st.error(f"Could not read OC customer list Excel file: {e}")
        return []


OC_CUSTOMER_LIST = load_oc_customer_list()

def find_oc_customers_in_board(board_text):
    """
    Scan the board text for any Opportunity Customer names or aliases.
    Returns a list of matched OC entries with the context they were found in.
    """
    board_lower = board_text.lower()
    matches = []

    for customer in OC_CUSTOMER_LIST:
        search_terms = [customer["name"]] + customer.get("aliases", [])
        found_terms = [term for term in search_terms if term.lower() in board_lower]

        if found_terms:
            matches.append({
                "customer": customer,
                "matched_on": found_terms,
            })

    return matches


def build_oc_alert_text(oc_matches):
    """
    Build a plain-text OC alert block to inject into the AI prompt and display in the UI.
    """
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
        lines.append(f"▶ CUSTOMER: {c['name'].upper()}")
        lines.append(f"  Matched on: {', '.join(match['matched_on'])}")
        lines.append(f"  Priority: {c['priority']}")
        lines.append(f"  Issue History: {c['issue']}")
        lines.append(f"  DC Requirements: {c['requirements']}")
        if c["sign_off"]:
            lines.append("  ⚠ DC Supervisor Sign-Off REQUIRED before this load ships.")
        if c["pictures"]:
            lines.append("  📷 Photos REQUIRED: 3 on dock + 3 during loading (6 total). Email to manager.")
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

    for row in range(3, 52):
        name = ws[f"E{row}"].value
        if name:
            names.append(str(name).strip())

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
        "Sunday": 0.20,
        "Monday": 0.18,
        "Tuesday": 0.18,
        "Wednesday": 0.19,
        "Thursday": 0.19,
        "Friday": 0.18,
        "Saturday": 0.21,
    }

    second_shift_pick = {
        "Sunday": 0.19,
        "Monday": 0.15,
        "Tuesday": 0.15,
        "Wednesday": 0.17,
        "Thursday": 0.17,
        "Friday": 0.17,
        "Saturday": 0.19,
    }

    first_shift_fp = {
        "Sunday": 0.28,
        "Monday": 0.32,
        "Tuesday": 0.40,
        "Wednesday": 0.35,
        "Thursday": 0.35,
        "Friday": 0.36,
        "Saturday": 0.31,
    }

    second_shift_fp = {
        "Sunday": 0.32,
        "Monday": 0.33,
        "Tuesday": 0.27,
        "Wednesday": 0.29,
        "Thursday": 0.28,
        "Friday": 0.30,
        "Saturday": 0.30,
    }

    if shift == "1st":
        cases_to_pick = total_cases * first_shift_pick.get(day, 0)
        full_pallets = (total_cases * first_shift_fp.get(day, 0)) / 70
    else:
        cases_to_pick = total_cases * second_shift_pick.get(day, 0)
        full_pallets = (total_cases * second_shift_fp.get(day, 0)) / 70

    return cases_to_pick, full_pallets


def calculate_needed(
    day,
    shift,
    total_cases,
    hours_remaining,
    total_outbound_loads_actual,
    crossroads_open,
    deer_creek_open,
    msb_open,
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
            raw_needed["Putaway"]
            + raw_needed["Replenishment"]
            + raw_needed["Full Pallets"]
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

    skill_map = {
        "Unloading": "U",
        "Receiving": "R",
        "Loading": "L",
        "Picking": "P",
        "Tasking": "T",
    }

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

                if best_fit(row, "Task") and (
                    has_skill(row, "T")
                    or has_skill(row, "L")
                    or has_skill(row, "P")
                ):
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
        idx
        for idx in present_indexes
        if any(name in str(staff.at[idx, "Name"]).lower() for name in preferred_extra_names)
    ]

    current_extra_idxs = [
        idx
        for idx in present_indexes
        if staff.at[idx, "Recommended Task"] == "Lead/Extra"
    ]

    for preferred_idx in preferred_idxs:
        if not current_extra_idxs:
            break

        if staff.at[preferred_idx, "Recommended Task"] == "Lead/Extra":
            continue

        swap_idx = None

        for extra_idx in current_extra_idxs:
            if not any(
                name in str(staff.at[extra_idx, "Name"]).lower()
                for name in preferred_extra_names
            ):
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


def build_recommendations(
    summary_table,
    present_recommendations,
    raw_needed,
    hours_remaining,
    notes,
):
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
            recommendations.append(
                f"{task}: approximately {diff * hours_remaining:.1f} labor-hours ahead / available capacity."
            )
        else:
            recommendations.append(f"{task}: Staffing is balanced.")

    picking_gap = int(summary_table.loc["Picking", "Difference"]) if "Picking" in summary_table.index else 0
    tasking_gap = int(summary_table.loc["Tasking", "Difference"]) if "Tasking" in summary_table.index else 0
    receiving_gap = int(summary_table.loc["Receiving", "Difference"]) if "Receiving" in summary_table.index else 0
    unloading_gap = int(summary_table.loc["Unloading", "Difference"]) if "Unloading" in summary_table.index else 0
    loading_gap = int(summary_table.loc["Loading", "Difference"]) if "Loading" in summary_table.index else 0
    lead_gap = int(summary_table.loc["Lead/Extra", "Difference"]) if "Lead/Extra" in summary_table.index else 0

    if picking_gap < 0:
        recommendations.append(
            "High picking short risk detected. Consider moving tasking labor into replenishment to protect pickers."
        )
        recommendations.append(
            "Avoid pulling pickers into unloading or loading unless outbound service is critical."
        )

        if tasking_gap > 0:
            recommendations.append(
                f"Tasking currently has {tasking_gap} extra worker(s). Consider temporarily assigning them to replenishment."
            )

        if lead_gap > 0:
            recommendations.append(
                "Lead/Extra capacity available. Consider flexing extra labor into replenishment or picking support."
            )

    if unloading_gap < 0 or receiving_gap < 0:
        recommendations.append(
            "Inbound flow risk detected. Falling behind may create dock congestion and delayed putaway."
        )
        recommendations.append(
            "Consider moving flexible tasking labor into unloading or receiving temporarily."
        )

        if tasking_gap > 1:
            recommendations.append(
                "Tasking has available labor that can support inbound operations."
            )

    if loading_gap < 0:
        recommendations.append(
            "Outbound loading risk detected. Late departures and service failures may increase."
        )
        recommendations.append(
            "Protect loading labor before reallocating to non-critical work."
        )

        if lead_gap > 0:
            recommendations.append(
                "Use Lead/Extra labor to support outbound staging or trailer cleanup."
            )

    if total_labor_gap > 1:
        recommendations.append("Operation currently has excess labor capacity.")
        recommendations.append(
            "Consider deep cleaning, trailer audits, replenishment cleanup, or cross-training."
        )
        recommendations.append(
            "Extra labor could be used proactively to prevent later picking shortages."
        )

    inbound_pressure = raw_needed["Unloading"] + raw_needed["Receiving"] + raw_needed["Putaway"]
    outbound_pressure = raw_needed["Picking"] + raw_needed["Loading"]

    if inbound_pressure > outbound_pressure * 1.3:
        recommendations.append("Inbound workload is significantly heavier than outbound.")
        recommendations.append(
            "Focus on unloading, receiving, and putaway to avoid congestion."
        )
    elif outbound_pressure > inbound_pressure * 1.3:
        recommendations.append("Outbound workload is significantly heavier than inbound.")
        recommendations.append(
            "Prioritize replenishment and picking continuity to avoid shorts."
        )

    if hours_remaining <= 4:
        recommendations.append(
            "Shift is entering final hours. Prioritize completion work and outbound execution."
        )
    elif hours_remaining >= 8:
        recommendations.append(
            "Enough shift time remains to strategically rebalance labor before bottlenecks form."
        )

    lower_notes = notes.lower()

    if "late" in lower_notes:
        recommendations.append(
            "Manager notes mention late loads. Prioritize outbound execution and trailer readiness."
        )

    if "short" in lower_notes:
        recommendations.append(
            "Manager notes indicate short risk. Protect replenishment and picking flow."
        )

    if "live" in lower_notes:
        recommendations.append(
            "Live loads detected in notes. Prioritize those doors before drop trailers."
        )

    if "cpu" in lower_notes:
        recommendations.append(
            "CPU loads referenced. Ensure loading labor is protected."
        )

    return recommen# ── BOARD EXCEL READING ───────────────────────────────────────────────────────
# This section makes Python parse the board first, then sends verified JSON to AI.
# The AI should interpret the data, not guess the data.

BOARD_STATUS_KEYWORDS = [
    "Loaded Short",
    "Picking/Short",
    "Picking Short",
    "Ready/Short",
    "R/S",
    "RTL",
    "Ready To Load",
    "Picking",
    "Completed",
    "Complete",
    "Loaded",
    "Late",
    "No Driver",
]

BOARD_DAY_NAMES = [
    "Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"
]


def board_cell_flags(cell):
    flags = []

    fill = cell.fill
    fill_color = ""
    if fill and fill.fgColor and fill.fgColor.type == "rgb":
        fill_color = str(fill.fgColor.rgb).upper()

    font = cell.font
    font_color = ""
    if font and font.color and font.color.type == "rgb":
        font_color = str(font.color.rgb).upper()

    if fill_color in ("FFFFFF00", "00FFFF00", "FFFF00"):
        flags.append("LOAD-CHECK")
    elif fill_color in ("FFADD8E6", "FF87CEEB", "FFADD8FF", "FFB0E0E6", "FF00BFFF"):
        flags.append("TT4-NEEDED")

    if font_color in ("FFFF0000", "00FF0000"):
        flags.append("CANADIAN")

    return flags


def normalize_board_text(value):
    if value is None:
        return ""
    text = str(value).replace("\n", " ").strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def looks_like_load_number(value):
    """
    Real board load numbers are usually 5-9 digits.
    This intentionally rejects dates, times, door numbers, and appointment times.
    """
    text = normalize_board_text(value)

    if not text:
        return ""

    if re.search(r"\d{1,2}/\d{1,2}/\d{2,4}", text):
        return ""

    if re.fullmatch(r"\d{1,2}:\d{2}", text):
        return ""

    if re.fullmatch(r"\d{1,2}", text):
        return ""

    text = re.sub(r"^LD", "", text, flags=re.IGNORECASE)
    digits = re.sub(r"[^0-9]", "", text)

    if 5 <= len(digits) <= 9:
        return digits

    return ""


def detect_board_day(row_text):
    for day_name in BOARD_DAY_NAMES:
        if re.search(rf"\b{day_name}\b", row_text, flags=re.IGNORECASE):
            return day_name
    return ""


def detect_board_date(row_text):
    match = re.search(r"\b\d{1,2}/\d{1,2}/\d{2,4}\b", row_text)
    if match:
        return normalize_date(match.group(0)) or match.group(0)
    return ""


def detect_board_time_from_text(text):
    text = normalize_board_text(text)

    if not text:
        return ""

    if re.fullmatch(r"\d{1,2}:\d{2}", text):
        return normalize_time(text)

    return ""


def detect_board_time(row_values):
    for value in row_values:
        candidate = detect_board_time_from_text(value)
        if candidate:
            return candidate
    return ""


def detect_board_status(row_text):
    text = normalize_board_text(row_text).upper()

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


def clean_board_header_key(value):
    return normalize_header(value)


def build_header_map_from_row(row_values):
    """
    Builds a header map when the board has a real header row.
    Uses careful matching so 'Load Type' does not accidentally become the load-number column.
    """
    header_map = {}

    for idx, value in enumerate(row_values):
        header = clean_board_header_key(value)

        if not header:
            continue

        if header in ["LOAD", "LOADNUMBER", "LOADNO", "LOADREF", "LOADREFERENCE", "ORDER", "ORDERNUMBER"]:
            header_map["load"] = idx
        elif header in ["DESTINATION", "CUSTOMER", "CUSTOMERNAME", "CUST", "SHIPTO", "CONSIGNEE", "CONSIGNEENAME"]:
            header_map["customer"] = idx
        elif "CARRIER" in header:
            header_map["carrier"] = idx
        elif header in ["TIME", "APPT", "APPTTIME", "APPOINTMENTTIME", "PUAPPTTIME"]:
            header_map["time"] = idx
        elif header in ["DOOR", "DOCK"]:
            header_map["door"] = idx
        elif header in ["TRAILER", "TRLR", "TRAILERNUMBER"]:
            header_map["trailer"] = idx
        elif header in ["STATUS", "STAT"]:
            header_map["status"] = idx
        elif header in ["TYPE", "LOADTYPE"]:
            header_map["type"] = idx
        elif header == "TT4":
            header_map["tt4"] = idx
        elif header in ["LOADER", "EMPLOYEE", "ASSIGNED"]:
            header_map["loader"] = idx
        elif header in ["COMMENTS", "COMMENT", "NOTES", "NOTE"]:
            header_map["comments"] = idx
        elif "PICK" in header:
            header_map["picks"] = idx
        elif "PULL" in header:
            header_map["pulls"] = idx

    return header_map


def row_looks_like_header(row_values):
    text = " ".join(str(v).upper() for v in row_values if str(v).strip())
    hits = 0

    for word in ["LOAD", "DESTINATION", "CUSTOMER", "CARRIER", "TIME", "DOOR", "STATUS", "TRAILER", "LOADER", "COMMENTS"]:
        if word in text:
            hits += 1

    return hits >= 2


def get_by_header(row_values, header_map, key):
    idx = header_map.get(key)
    if idx is None or idx >= len(row_values):
        return ""
    return normalize_board_text(row_values[idx])


def detect_ticket_count(row_values, header_map, key):
    value = get_by_header(row_values, header_map, key)
    if value:
        return parse_number(value)
    return 0


def first_nonempty_index(row_values):
    for idx, value in enumerate(row_values):
        if normalize_board_text(value):
            return idx
    return 0


def find_status_column_value(row_values):
    for value in row_values:
        status = detect_board_status(value)
        if status:
            return status
    return ""


def parse_board_rows_from_records(records, source_name):
    """
    Parses your board layout:
    Load | Destination | Carrier | Time | Door | Trailer | Status | TT4 | Loader | Comments

    It also works when the first column/header is blank, because it searches for the first real load number
    and then reads the columns relative to that load number.
    """
    board_rows = []
    current_day = ""
    current_date = ""
    header_map = {}

    for record in records:
        row_number = record["row_number"]
        row_values = [normalize_board_text(v) for v in record["values"]]
        row_flags = record.get("flags", [])
        row_text = " ".join(v for v in row_values if v)

        if not row_text.strip():
            continue

        detected_day = detect_board_day(row_text)
        detected_date = detect_board_date(row_text)

        load_candidates = []
        for v in row_values:
            candidate = looks_like_load_number(v)
            if candidate:
                load_candidates.append(candidate)

        has_load = bool(load_candidates)

        if row_looks_like_header(row_values) and not has_load:
            header_map = build_header_map_from_row(row_values)
            continue

        if (detected_day or detected_date) and not has_load:
            if detected_day:
                current_day = detected_day
            if detected_date:
                current_date = detected_date
            continue

        if not has_load:
            continue

        load_number = get_by_header(row_values, header_map, "load")
        load_number = looks_like_load_number(load_number) or load_candidates[0]

        try:
            load_idx = next(
                i for i, v in enumerate(row_values)
                if looks_like_load_number(v) == load_number
            )
        except StopIteration:
            load_idx = first_nonempty_index(row_values)

        def get_pos(offset):
            idx = load_idx + offset
            if 0 <= idx < len(row_values):
                return normalize_board_text(row_values[idx])
            return ""

        customer = get_by_header(row_values, header_map, "customer") or get_pos(1)
        carrier = get_by_header(row_values, header_map, "carrier") or get_pos(2)

        appt_time = (
            get_by_header(row_values, header_map, "time")
            or detect_board_time_from_text(get_pos(3))
            or detect_board_time(row_values)
        )

        door = get_by_header(row_values, header_map, "door") or get_pos(4)
        trailer = get_by_header(row_values, header_map, "trailer") or get_pos(5)

        status_raw = get_by_header(row_values, header_map, "status") or get_pos(6)
        status = detect_board_status(status_raw) or find_status_column_value(row_values)

        load_type = get_by_header(row_values, header_map, "type")
        tt4 = get_by_header(row_values, header_map, "tt4") or get_pos(7)
        loader = get_by_header(row_values, header_map, "loader") or get_pos(8)
        comments = get_by_header(row_values, header_map, "comments") or get_pos(9)

        picks = detect_ticket_count(row_values, header_map, "picks")
        pulls = detect_ticket_count(row_values, header_map, "pulls")

        if not comments:
            comments = row_text

        flags = sorted(set(row_flags))

        board_rows.append(
            {
                "source": source_name,
                "row_number": row_number,
                "day": current_day,
                "date": current_date,
                "load_number": load_number,
                "customer": customer,
                "carrier": carrier,
                "appt_time": appt_time,
                "door": door,
                "trailer": trailer,
                "status": status,
                "type": load_type,
                "tt4": tt4,
                "loader": loader,
                "comments": comments,
                "picks": picks,
                "pulls": pulls,
                "flags": flags,
                "raw_row": row_text,
            }
        )

    return board_rows


def board_records_from_excel(board_file):
    board_file.seek(0)
    file_name = board_file.name.lower()
    all_board_rows = []

    if file_name.endswith(".xls"):
        sheets = pd.read_excel(board_file, sheet_name=None, header=None, engine="xlrd")

        for sheet_name, df in sheets.items():
            df = df.fillna("")
            records = []

            for idx, row in df.iterrows():
                records.append(
                    {
                        "row_number": int(idx) + 1,
                        "values": [normalize_board_text(v) for v in row.tolist()],
                        "flags": [],
                    }
                )

            all_board_rows.extend(parse_board_rows_from_records(records, sheet_name))

        return all_board_rows

    wb = load_workbook(board_file, data_only=True)

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        records = []

        for row_idx, row in enumerate(ws.iter_rows(), start=1):
            values = []
            flags = []

            for cell in row:
                val_str = normalize_board_text(cell.value)
                values.append(val_str)

                for flag in board_cell_flags(cell):
                    flags.append(flag)

            records.append(
                {
                    "row_number": row_idx,
                    "values": values,
                    "flags": sorted(set(flags)),
                }
            )

        all_board_rows.extend(parse_board_rows_from_records(records, sheet_name))

    return all_board_rows


def board_records_from_csv(board_file):
    board_file.seek(0)
    df = pd.read_csv(board_file, header=None).fillna("")
    records = []

    for idx, row in df.iterrows():
        records.append(
            {
                "row_number": int(idx) + 1,
                "values": [normalize_board_text(v) for v in row.tolist()],
                "flags": [],
            }
        )

    return parse_board_rows_from_records(records, "CSV Board")


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

        if "LATE" in status_upper or " LATE " in f" {raw_upper} ":
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

    # Count blank/not-started only after known statuses.
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

    # Remove exact duplicate priority rows.
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
    compact_rows = []

    for row in board_rows:
        compact_rows.append(
            {
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
            }
        )

    return compact_rows


def read_board_file_to_text(board_file):
    board_file.seek(0)
    file_name = board_file.name.lower()

    try:
        if file_name.endswith(".csv"):
            board_rows = board_records_from_csv(board_file)
        else:
            board_rows = board_records_from_excel(board_file)

        board_summary = build_python_board_summary(board_rows)
        compact_rows = compact_board_rows_for_ai(board_rows)

        payload = {
            "python_verified_summary": board_summary,
            "structured_load_rows": compact_rows,
            "debug": {
                "file_name": board_file.name,
                "rows_parsed_by_python": len(board_rows),
                "message": "If rows_parsed_by_python is 0, the board layout did not match parser rules or the file could not be read.",
            },
            "instructions_for_ai": [
                "Use python_verified_summary as the source of truth for counts.",
                "Use structured_load_rows for load-level details, load numbers, customers, status, doors, times, flags, and comments.",
                "Do not invent load numbers or counts that are not present in this JSON.",
                "If a field is blank, say unclear instead of guessing.",
            ],
        }

        return json.dumps(payload, indent=2, ensure_ascii=False)

    except Exception as e:
        return json.dumps(
            {
                "error": f"Could not read board file: {str(e)}",
                "python_verified_summary": {},
                "structured_load_rows": [],
                "debug": {
                    "file_name": getattr(board_file, "name", "Unknown"),
                    "rows_parsed_by_python": 0,
                    "error": str(e),
                },
            },
            indent=2,
            ensure_ascii=False,
        )


    

def analyze_board_with_groq(
    board_text,
    day,
    shift,
    total_cases,
    hours_remaining,
    total_outbound_loads,
    crossroads_open,
    deer_creek_open,
    msb_open,
    needed,
    summary_table,
    cases_to_pick,
    inbound_pallets,
    notes,
    oc_alert_text=None,
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
            f"  {task}: Need {int(row['Needed'])}, Have {int(row['Assigned'])}, "
            f"Gap {int(row['Difference'])} ({row['Status']})"
        )

    staffing_summary = "\n".join(staffing_lines)

    plants_open = [
        p
        for p, status in [
            ("Crossroads", crossroads_open),
            ("Deer Creek", deer_creek_open),
            ("MSB", msb_open),
        ]
        if status == "YES"
    ]

    oc_section = ""
    if oc_alert_text:
        oc_section = f"\n\n{oc_alert_text}\n"

    base_context = f"""
You are an experienced warehouse operations shift manager analyzing an outbound load board that was read directly from an Excel file (cell values, not a screenshot or image). Python has already parsed the board into structured JSON. Treat the Python-verified counts and structured load rows as the source of truth.
Use short bullet points. don't over explain.
Use the structured load rows and Python summary instead of guessing from raw Excel text. 
The idea is to get as ahead as possible with the current resources. 

When reading: separate loads and their data by day, focus on today but still mention when there are still loads on the board from days before, from what day and what is happening with them.
Don't assume the load number for the day. Look for today's day and then count how many load number it has under it, until the next day appears. 

Additional warehouse operation context:
This is a high-volume outbound grocery distribution center operation. This is the first shift and it starts from 6 am to 4:30 pm with 9.5 workable hours. Setting up the second shift for success can vary, but if my morning shift has all loads RTL and the appointments are until 4pm that is still success, not behind. 
The outbound board represents live warehouse execution, not future planning. The board uses 24 hour clock instead of 12.

The manager using this system is focused on:
- Avoiding late departures
- Protecting dock flow
- Prioritizing live loads correctly
- Reducing congestion
- Getting ahead instead of reacting late
- Preventing shorts
- Keeping pickers productive


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
- the operation is behind schedule when we are don't have ready to load the loads for the time of the day it is. 
- If outbound workload is heavier than staffing, recommend labor moves immediately.

Management philosophy:
The goal is not only to survive the shift. The goal is to get ahead early enough that later appointments are protected. We only send people to manufacturing if we are overstaffed.

The manager needs:
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
- Give achievable operational goals for the next 30 minutes and next 2 hours and the end of the shift
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
- The data below was extracted and summarized by Python before being sent to you. Use the JSON counts as the source of truth.
- Cells annotated with [LOAD-CHECK] had a yellow fill in Excel, meaning that load needs a load check.
- Cells annotated with [CANADIAN] had red font in Excel, meaning it is a Canadian load.
- If a color annotation is absent, the cell had no special flag — do not guess.
- Blank status on the board means the load is not currently being worked.
- R/S means Ready to load but still short on full pallets.
- Picking is measured in tickets on the board, but analyze everything in cases. Our average is 60 cases per picking ticket.
- If a column or value is unclear or missing, say "unclear" — do not guess information.
- Read the structured JSON rows and give insights based only on those rows and verified counts. 

Here is the Python-verified structured board data. Use this JSON as the source of truth for counts and load-level details:
{board_text}
"""

    output_structure = """
Read the board carefully row by row.

Give me a clear, practical warehouse manager analysis in plain English covering:

1. Board Summary:
- Break loads down by status and day: RTL, R/S, Late, Picking, Picking/Short, Loaded Short, Completed, blank/not started, etc.
- Specify how many loads are completed today out of the total for the day.
- Specify any late loads, from when, if they are occupying a door, and which door.

2. Opportunity Customer (OC) Alerts:
- List every load on the board that belongs to a customer on the Opportunity Customer List.
- For each OC load: state the load number, customer name, current status, appointment time, and EXACTLY what special actions are required before this load ships.
- If pictures are required, state when they should be taken and who should own it.
- If supervisor sign-off is required, state who should sign off and when.
- If no OC customers are on the board today, state clearly: "No Opportunity Customers detected on today's board."

3. Picking & Short Risk:
- How many loads have not been started?
- Given cases-to-pick and current staffing, are we at risk of falling further behind? In easy words, yes or no and why.
- How big is the risk? Explain what are the risk factors.
- Specify people from what areas we can move from and to where. Should we consider sending people to manufacturing to reduce short risks? Specify people from what areas we can move staff from and to where.

4. Prioritization:
- Are there any loads we should prioritize? Be specific, add load numbers.
- How and why should we prioritize them?

5. Cross-Analysis with Staffing:
- Given staffing gaps or surpluses, which problems can we actually fix right now?
- Where should labor move first?
- Based on staffing and demand, what should be an achievable goal for this shift?
- Can we get ahead? Yes or no and why. What is an achievable goal for the end of shift?
- Given all this information, how far ahead or behind are we forecasted to finish this shift?
- Give me the load appointment times we should be picking and have RTL by the end of this shift based on the above stated goal.

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
            messages=[
                {
                    "role": "user",
                    "content": base_context + output_structure,
                }
            ],
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
                        "Make sure to count the loads per day from every cell. "
                        "Follow the exact same output structure as the initial analysis. "
                        "Opportunity Customer (OC) alerts must appear early and be complete — never omit or shorten them."
                    ),
                },
                {
                    "role": "user",
                    "content": (
                        f"=== PYTHON-VERIFIED BOARD DATA AND OPERATIONAL CONTEXT ===\n{base_context}\n\n"
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

    # ── OC Alert block at top of sheet if matches found ──────────────────────
    if oc_matches:
        ws.cell(current_row, 1).value = "⚠ OPPORTUNITY CUSTOMER ALERT — SPECIAL HANDLING REQUIRED"
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
                oc_lines.append("⚠ DC Supervisor Sign-Off REQUIRED before this load ships.")
            if c["pictures"]:
                oc_lines.append("📷 Photos REQUIRED: 3 on dock + 3 during loading (6 total). Email to manager.")

            for line in oc_lines:
                cell = ws.cell(current_row, 1, line)
                cell.font = Font(size=10, bold=("CUSTOMER:" in line or "⚠" in line or "📷" in line))
                cell.fill = PatternFill("solid", fgColor=light_orange)
                cell.alignment = Alignment(wrap_text=True, vertical="top")
                cell.border = border
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=7)
                ws.row_dimensions[current_row].height = max(15, min(60, len(line) // 5))
                current_row += 1

            current_row += 1

        current_row += 1

    # ── AI analysis lines ─────────────────────────────────────────────────────
    for line in analysis_text.split("\n"):
        cell = ws.cell(current_row, 1, line)
        cell.alignment = Alignment(wrap_text=True, vertical="top")
        cell.border = border
        ws.merge_cells(
            start_row=current_row,
            start_column=1,
            end_row=current_row,
            end_column=7,
        )
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
    light_orange = "FCE4D6"
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

    # ── OC Alert banner on dashboard if matches found ─────────────────────────
    oc_banner_row = 6

    if oc_matches:
        customer_names = ", ".join(m["customer"]["name"].upper() for m in oc_matches)
        ws_dash.cell(oc_banner_row, 1).value = (
            f"⚠ OC ALERT: Opportunity Customers on today's board — {customer_names} — See 'Board Analysis' tab for full requirements."
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
        values = [
            task,
            int(row["Needed"]),
            int(row["Assigned"]),
            int(row["Difference"]),
            row["Status"],
        ]

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
        ws_dash.merge_cells(
            start_row=rec_row,
            start_column=7,
            end_row=rec_row,
            end_column=11,
        )
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
        values = [
            row["Name"],
            row["Skills"],
            row["Best Fit"],
            row["Recommended Task"],
        ]

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

    data = Reference(
        ws_dash,
        min_col=2,
        max_col=3,
        min_row=header_row,
        max_row=header_row + len(summary_table),
    )

    cats = Reference(
        ws_dash,
        min_col=1,
        min_row=header_row + 1,
        max_row=header_row + len(summary_table),
    )

    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 9
    bar.width = 15
    bar.legend.position = "r"

    ws_dash.add_chart(bar, f"E{chart_anchor_row}")

    pie = PieChart()
    pie.title = "Assigned Labor Distribution"

    pie_data = Reference(
        ws_dash,
        min_col=3,
        min_row=header_row,
        max_row=header_row + len(summary_table),
    )

    pie_cats = Reference(
        ws_dash,
        min_col=1,
        min_row=header_row + 1,
        max_row=header_row + len(summary_table),
    )

    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 9
    pie.width = 13
    pie.legend.position = "r"

    ws_dash.add_chart(pie, f"I{chart_anchor_row}")

    for col in range(1, 12):
        ws_dash.column_dimensions[get_column_letter(col)].width = 18

    ws_dash.column_dimensions["A"].width = 22
    ws_dash.column_dimensions["G"].width = 35
    ws_dash.column_dimensions["H"].width = 35
    ws_dash.column_dimensions["I"].width = 35
    ws_dash.column_dimensions["J"].width = 35
    ws_dash.column_dimensions["K"].width = 35

    ws_dash.freeze_panes = f"A{header_row}"


def build_email_draft(
    day,
    shift,
    total_cases,
    hours_remaining,
    total_outbound_loads_day,
    summary_table,
    present_recommendations,
    recommendations,
    board_analysis_text=None,
    oc_matches=None,
):
    total_present = len(present_recommendations)
    total_needed = int(summary_table["Needed"].sum())
    total_assigned = int(summary_table["Assigned"].sum())
    overall_gap = total_assigned - total_needed

    subject = f"{day} {shift} Shift Staffing Report"

    staffing_lines = []

    for task, row in summary_table.iterrows():
        staffing_lines.append(
            f"- {task}: "
            f"Need {int(row['Needed'])}, "
            f"Assigned {int(row['Assigned'])}, "
            f"Gap {int(row['Difference'])} "
            f"({row['Status']})"
        )

    top_recommendations = "\n".join(
        [f"- {rec}" for rec in recommendations[:8]]
    )

    oc_email_block = ""
    if oc_matches:
        oc_lines = ["\n⚠ OPPORTUNITY CUSTOMER ALERT:"]
        for match in oc_matches:
            c = match["customer"]
            oc_lines.append(f"  - {c['name'].upper()} [{c['priority']}]: {c['requirements']}")
            if c["sign_off"]:
                oc_lines.append("    → Supervisor sign-off REQUIRED before shipping.")
            if c["pictures"]:
                oc_lines.append("    → 6 photos required (3 on dock, 3 loading). Email to manager.")
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


# ── STREAMLIT INTERFACE ───────────────────────────────────────────────────────

st.sidebar.header("Daily Inputs")

day = st.sidebar.selectbox(
    "Day",
    ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"],
)

shift = st.sidebar.selectbox("Shift", ["1st", "2nd"])

total_cases = st.sidebar.number_input(
    "Total Cases for Today",
    min_value=0,
    step=1,
    value=0,
)

hours_remaining = st.sidebar.number_input(
    "Hours Remaining in Shift",
    min_value=0.0,
    step=0.25,
    value=8.0,
)

total_outbound_loads_day = st.sidebar.number_input(
    "Total Outbound Loads for the Day",
    min_value=0,
    step=1,
    value=0,
)

crossroads_open = st.sidebar.selectbox("Crossroads plant open?", ["YES", "NO"])
deer_creek_open = st.sidebar.selectbox("Deer Creek plant open?", ["YES", "NO"])
msb_open = st.sidebar.selectbox("MSB plant open?", ["YES", "NO"])

present_workers = st.sidebar.multiselect("Who is present?", names)

notes = st.sidebar.text_area("Operations Notes")

st.markdown("---")
st.subheader("Outbound Board Excel")

board_file = st.file_uploader(
    "Upload the outbound load board Excel",
    type=["xlsx", "xls", "csv"],
    help="Cell values and color flags (yellow = load check, light-blue = TT4, red font = Canadian) are read directly from the file.",
)

if board_file:
    st.success("Board file loaded — ready for analysis.")

# ── OC List preview (expandable) ─────────────────────────────────────────────
with st.expander("View Opportunity Customer List from Excel"):
    oc_preview_rows = []
    for c in OC_CUSTOMER_LIST:
        oc_preview_rows.append({
            "Customer": c["name"].title(),
            "Priority": c["priority"],
            "Issue": c["issue"],
            "Sign-Off Required": "Yes" if c["sign_off"] else "No",
            "Photos Required": "Yes" if c["pictures"] else "No",
        })
    st.dataframe(pd.DataFrame(oc_preview_rows), use_container_width=True)

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

    for row in range(3, 52):
        ws[f"F{row}"] = ""

    selected = {name.strip().lower() for name in present_workers}

    for row in range(3, 52):
        worker_name = ws[f"E{row}"].value

        if worker_name and str(worker_name).strip().lower() in selected:
            ws[f"F{row}"] = "x"

    ws["B12"] = notes

    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(working_file)

    needed, raw_needed, cases_to_pick, full_pallets, inbound_pallets = calculate_needed(
        day,
        shift,
        total_cases,
        hours_remaining,
        total_outbound_loads_actual,
        crossroads_open,
        deer_creek_open,
        msb_open,
    )

    staff = pd.read_excel(
        working_file,
        sheet_name="Staffing sheet 1ST Shift",
        usecols="A,D,F,I,T",
    )

    staff.columns = ["Name", "Skills", "Best Fit", "Recommended Task", "Present"]
    staff = staff[staff["Name"].notna()].copy()

    selected_names = {name.strip().lower() for name in present_workers}

    staff["Present"] = staff["Name"].astype(str).str.strip().str.lower().apply(
        lambda x: "x" if x in selected_names else ""
    )

    staff = generate_recommendations(staff, needed)

    present_recommendations, summary_table = build_summary(staff, needed)

    recommendations = build_recommendations(
        summary_table,
        present_recommendations,
        raw_needed,
        hours_remaining,
        notes,
    )

    wb = load_workbook(working_file)

    write_recommendations_to_excel(wb, staff)

    board_analysis_text = None
    oc_matches = []

    if board_file is not None:
        with st.spinner("Reading board file → scanning for Opportunity Customers → running analysis → validating → finalizing..."):
            board_text = read_board_file_to_text(board_file)

            # ── OC detection ─────────────────────────────────────────────────
            oc_matches = find_oc_customers_in_board(board_text)
            oc_alert_text = build_oc_alert_text(oc_matches)

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

    st.subheader("Staffing Summary")
    st.dataframe(summary_table, use_container_width=True)

    st.subheader("Recommended Staffing Board")
    st.dataframe(
        present_recommendations[
            ["Name", "Skills", "Best Fit", "Recommended Task"]
        ].reset_index(drop=True),
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

    # ── OC Alerts UI block ────────────────────────────────────────────────────
    if oc_matches:
        st.markdown("---")
        st.subheader(" Opportunity Customer Alerts")
        st.error(
            "The following customers on today's board are on the **Opportunity Customer List** "
            "and require special DC actions before their loads ship."
        )

        for match in oc_matches:
            c = match["customer"]
            with st.expander(f"{c['name'].upper()}  —  Priority: {c['priority']}", expanded=True):
                st.markdown(f"**Issue History:** {c['issue']}")
                st.markdown(f"**DC Requirements:** {c['requirements']}")
                if c["sign_off"]:
                    st.markdown(" **DC Supervisor Sign-Off REQUIRED before this load ships.**")
                if c["pictures"]:
                    st.markdown("**Photos REQUIRED:** 3 on dock + 3 during loading (6 total). Email to manager.")
    elif board_file is not None:
        st.info(" No Opportunity Customers detected on today's board.")

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

    st.text_input(
        "Email Subject",
        value=email_subject
    )

    st.text_area(
        "Email Body",
        value=email_body,
        height=500,
    )
