import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from io import BytesIO
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
        "Unloading": whole_workers(raw_needed["Unloading"]),
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

    return recommendations


# BOARD EXCEL READING 
# Reads directly from Excel cell values — no OCR, no image processing.
# openpyxl is used so we can also capture cell fill colors for status flags.

def read_board_file_to_text(board_file):
    """
    Read the board Excel/CSV directly from cell values.
    For Excel files we use openpyxl so we can detect fill colors on cells
    (yellow = load check needed, light-blue = TT4 needed, red font = Canadian).
    Returns a plain-text string the AI can reason over.
    """
    board_file.seek(0)
    file_name = board_file.name.lower()

    # CSV path 
    if file_name.endswith(".csv"):
        try:
            df = pd.read_csv(board_file)
            df = df.dropna(how="all").dropna(axis=1, how="all").fillna("")
            return df.to_csv(index=False)
        except Exception as e:
            return f"Could not read CSV board file: {e}"

    # Excel path — cell-level read with color detection 
    try:
        board_file.seek(0)
        wb = load_workbook(board_file, data_only=True)   # data_only=True → formula results
        sections = []

        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]

            # Collect rows, skipping completely empty ones
            rows_data = []
            headers = None

            for row_idx, row in enumerate(ws.iter_rows(), start=1):
                row_values = []
                color_flags = []

                for cell in row:
                    val = cell.value
                    val_str = "" if val is None else str(val).strip()

                    # Detect fill color flags
                    fill = cell.fill
                    fill_color = ""
                    if fill and fill.fgColor and fill.fgColor.type == "rgb":
                        fill_color = fill.fgColor.rgb.upper()  # e.g. "FFFFFF00" = yellow

                    # Detect font color flags
                    font = cell.font
                    font_color = ""
                    if font and font.color and font.color.type == "rgb":
                        font_color = font.color.rgb.upper()

                    # Annotate special cells inline so the AI sees them clearly
                    flags = []
                    if fill_color in ("FFFFFF00", "00FFFF00", "FFFF00"):        # yellow variants
                        flags.append("[LOAD-CHECK]")
                    elif fill_color in ("FFADD8E6", "FF87CEEB", "FFADD8FF",    # light-blue variants
                                        "FFB0E0E6", "FF00BFFF"):
                        flags.append("[TT4-NEEDED]")

                    if font_color in ("FFFF0000", "00FF0000"):                  # red font
                        flags.append("[CANADIAN]")

                    annotated = val_str + (" " + " ".join(flags) if flags else "")
                    row_values.append(annotated)

                # Skip rows that are entirely blank
                if all(v.strip() == "" for v in row_values):
                    continue

                if row_idx == 1:
                    headers = row_values
                else:
                    rows_data.append(row_values)

            if not rows_data:
                continue

            # Build a simple CSV-style block for this sheet
            section_lines = [f"--- SHEET: {sheet_name} ---"]
            if headers:
                section_lines.append(",".join(headers))
            for r in rows_data:
                section_lines.append(",".join(r))

            sections.append("\n".join(section_lines))

        if not sections:
            return "No readable board data found in the uploaded Excel file."

        return "\n\n".join(sections)

    except Exception as e:
        return f"Could not read Excel board file: {e}"


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

    prompt = f"""
You are an experienced warehouse operations shift manager analyzing an outbound load board that was read directly from an Excel file (cell values, not a screenshot or image). All data is clean and structured — treat every field as accurate cell content.
Use short bullet points. don't over explain. 

When reading: separate loads and their data by day, focus on today but still mention when they are still loads on the board from days before, from what day and what is happening with them. 


Additional warehouse operation context:
This is a high-volume outbound grocery distribution center operation. This is the first shift and it starts from 5 am to 4 pm with 9.5 workable hours. Setting up the second shift for success can varies, but if my morning shift have all loads RTL and the appointments are until 3pm that is still success, not behind. 
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

- RTL = Ready To Load
  Product is staged and ready. Loader can execute.

- R/S = Ready/Short
  Load is mostly ready but missing full pallets or replenishment inventory.
  This is a major operational risk and can quickly become late.

- Picking = Order currently being picked.

- Picking/Short = Picking in progress but inventory shortages are occurring.
  This usually means replenishment or manufacturing support is needed.

- Loaded Short = Trailer loaded but missing product.
  This is a severe service risk.

- Live = Trailer physically waiting at the dock.
  Live loads always have higher priority than drop trailers.

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

The goal is not only to survive the shift.
The goal is to get ahead early enough that later appointments are protected.
We only sending people to manufacturing if it's going to benefit us.

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

Here is today's operational context:
- Day: {day}, Shift: {shift}
- Total cases forecast for today: {total_cases:,}
- Cases to pick this shift: {cases_to_pick:,.0f}
- Hours remaining in shift: {hours_remaining}
- Total outbound loads scheduled today: {total_outbound_loads}
- Inbound pallets expected: {inbound_pallets:,} (Plants open: {", ".join(plants_open) if plants_open else "None"})
- Manager notes: {notes if notes.strip() else "None"}

Current staffing vs. what we need:
{staffing_summary}

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

Read the board carefully row by row.

Give me a clear, practical warehouse manager analysis in plain English covering:

1. Board Summary:
- Break loads down by status and day: RTL, R/S, Late, Picking, Picking/Short, Loaded Short, Completed, blank/not started, etc.
- Specify how many loads are completed today out of the total for the day.
- Specify if any late loads are occupying a door, and which door.

2. Picking & Short Risk:
- How many loads have not been started?
- Given cases-to-pick and current staffing, are we at risk of falling further behind? In easy words, yes or no and why. 
- How big is the risk? Explain what are the risk factors. 
- Can we get ahead? Yes or no and why
- Given all this information, how far ahead can we finish this shift?
- Give me the load appointment times we should be picking by the end of this shift.
- Specify people from what areas we can move from and to where. Should we consider sending people to manufacturing to reduce short risks Specify people from what areas we can move staff from and to where.

3. Prioritization:
- Are there any loads we should prioritize? Be specific add load numbers
- How and why should we prioritize them?

4. Cross-Analysis with Staffing:
- Given staffing gaps or surpluses, which problems can we actually fix right now?
- Where should labor move first?
- Based on staffing and demand, what should be an achievable goal for this shift?
- How ahead or behind should we finish this shift?

5. Top 3 Action Items:
- What are the 3 most important things the manager should do in the next 30 minutes?
- What are the 3 most important things the manager should do in the next 2 hours to achieve today's goal?


Make sure every recommendation and suggested action is achievable and following the same direction. 
Have somewhere where you clearly set the expectations for the shift and explain why. I want this easy to identify. 
Keep the tone like a smart, experienced ops manager talking to another manager.
No corporate fluff.
Be clear, practical, and actionable.
Add times and case/pallet numbers to every goal so progress is measurable.
Include what-if scenarios: if X happens, here is what to do.
Only use data, do not guess
Talk about how you are heading the second shift for success. 
When suggesting to think about moving staff specify from where to where. 
Remember even though we have 62 loads for the day it is separated in 2 shifts. We load approximately 52% of loads in the first shift. Take that into consideration, we still can have the loads ready to load for second shift. Read the board and check the times. 
When making suggestions that we should be ready to load up to a specific hour do not use a range, be specific. 
"""

    try:
        response = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[
                {
                    "role": "user",
                    "content": prompt,
                }
            ],
            temperature=0.2,
            max_completion_tokens=2500,
        )

        return response.choices[0].message.content

    except Exception as e:
        return f"Board analysis could not be completed: {str(e)}"


def write_board_analysis_to_excel(wb, analysis_text):
    sheet_name = "Board Analysis"

    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        ws.delete_rows(1, ws.max_row)
    else:
        ws = wb.create_sheet(sheet_name)

    dark_blue = "0F5B78"
    white = "FFFFFF"
    light_blue = "D9EAF7"
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


# ── WRITTEN RECOMMENDATIONS ──────────────────────────────────────────────────

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


def build_dashboard(wb, summary_table, present_recommendations, recommendations):
    if "Staffing Dashboard" in wb.sheetnames:
        ws_dash = wb["Staffing Dashboard"]
        ws_dash.delete_rows(1, ws_dash.max_row)
    else:
        ws_dash = wb.create_sheet("Staffing Dashboard")

    dark_blue = "0F5B78"
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

    ws_dash["A6"] = "Needed vs Assigned"
    ws_dash["A6"].font = Font(size=14, bold=True)

    headers = ["Task", "Needed", "Assigned", "Difference", "Status"]

    for c, header in enumerate(headers, 1):
        cell = ws_dash.cell(7, c)
        cell.value = header
        cell.font = Font(bold=True, color=white)
        cell.fill = PatternFill("solid", fgColor=dark_blue)
        cell.border = border
        cell.alignment = Alignment(horizontal="center")

    for r, (task, row) in enumerate(summary_table.iterrows(), 8):
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

    ws_dash["G6"] = "Written Recommendations / What-Ifs"
    ws_dash["G6"].font = Font(size=14, bold=True)

    rec_row = 7

    for rec in recommendations:
        ws_dash[f"G{rec_row}"] = f"• {rec}"
        ws_dash[f"G{rec_row}"].alignment = Alignment(wrap_text=True, vertical="top")

        ws_dash.merge_cells(
            start_row=rec_row,
            start_column=7,
            end_row=rec_row,
            end_column=11,
        )

        rec_row += 1

    board_start = max(16, rec_row + 2)

    ws_dash[f"A{board_start}"] = "Recommended Staffing Board"
    ws_dash[f"A{board_start}"].font = Font(size=14, bold=True)

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

    bar = BarChart()
    bar.title = "Needed vs Assigned"
    bar.y_axis.title = "Workers"
    bar.x_axis.title = "Task"

    data = Reference(
        ws_dash,
        min_col=2,
        max_col=3,
        min_row=7,
        max_row=7 + len(summary_table),
    )

    cats = Reference(
        ws_dash,
        min_col=1,
        min_row=8,
        max_row=7 + len(summary_table),
    )

    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 9
    bar.width = 15
    bar.legend.position = "r"

    ws_dash.add_chart(bar, "E28")

    pie = PieChart()
    pie.title = "Assigned Labor Distribution"

    pie_data = Reference(
        ws_dash,
        min_col=3,
        min_row=7,
        max_row=7 + len(summary_table),
    )

    pie_cats = Reference(
        ws_dash,
        min_col=1,
        min_row=8,
        max_row=7 + len(summary_table),
    )

    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)
    pie.height = 9
    pie.width = 13
    pie.legend.position = "r"

    ws_dash.add_chart(pie, "I28")

    for col in range(1, 12):
        ws_dash.column_dimensions[get_column_letter(col)].width = 18

    ws_dash.column_dimensions["A"].width = 22
    ws_dash.column_dimensions["G"].width = 35
    ws_dash.column_dimensions["H"].width = 35
    ws_dash.column_dimensions["I"].width = 35
    ws_dash.column_dimensions["J"].width = 35
    ws_dash.column_dimensions["K"].width = 35

    ws_dash.freeze_panes = "A7"


# ── STREAMLIT INTERFACE ──────────────────────────────────────────────────────

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
st.subheader("📋 Outbound Board Excel / CSV")

board_file = st.file_uploader(
    "Upload the outbound load board Excel or CSV file",
    type=["xlsx", "xls", "csv"],
    help="Cell values and color flags (yellow = load check, light-blue = TT4, red font = Canadian) are read directly from the file.",
)

if board_file:
    st.success("Board file loaded — ready for analysis.")

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

    build_dashboard(wb, summary_table, present_recommendations, recommendations)

    board_analysis_text = None

    if board_file is not None:
        with st.spinner("Reading board Excel file and analyzing with Groq AI..."):
            board_text = read_board_file_to_text(board_file)

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
            )

            write_board_analysis_to_excel(wb, board_analysis_text)

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

    st.download_button(
        label="Download Staffing Report",
        data=output,
        file_name="Staffing Report Generated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
