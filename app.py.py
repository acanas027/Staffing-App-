import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.utils import get_column_letter
from io import BytesIO
import os
import shutil

st.set_page_config(page_title="Staffing Report Generator", layout="wide")

st.title("Staffing Report Generator")
st.write("Enter daily inputs, select who is present, and generate the staffing report.")

TEMPLATE_FILE = "staffing_template.xlsx"

if not os.path.exists(TEMPLATE_FILE):
    st.error("Template file not found. Put staffing_template.xlsx in the same folder as App.py.")
    st.stop()


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
        "Sunday": 0.20, "Monday": 0.18, "Tuesday": 0.18,
        "Wednesday": 0.19, "Thursday": 0.19, "Friday": 0.18,
        "Saturday": 0.21,
    }

    second_shift_pick = {
        "Sunday": 0.19, "Monday": 0.15, "Tuesday": 0.15,
        "Wednesday": 0.17, "Thursday": 0.17, "Friday": 0.17,
        "Saturday": 0.19,
    }

    first_shift_fp = {
        "Sunday": 0.28, "Monday": 0.32, "Tuesday": 0.40,
        "Wednesday": 0.35, "Thursday": 0.35, "Friday": 0.36,
        "Saturday": 0.31,
    }

    second_shift_fp = {
        "Sunday": 0.32, "Monday": 0.33, "Tuesday": 0.27,
        "Wednesday": 0.29, "Thursday": 0.28, "Friday": 0.30,
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
    msb_open
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
                    has_skill(row, "T") or has_skill(row, "L") or has_skill(row, "P")
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

    # FINAL OVERRIDE:
    # If Lead/Extra exists, Will Perkins and Antonio D. become Lead/Extra first.
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
            recommendations.append(
                f"{task}: approximately {diff * hours_remaining:.1f} labor-hours ahead / available capacity."
            )
        else:
            recommendations.append(f"{task}: Staffing is balanced.")

    if notes.strip():
        recommendations.append(f"Manager notes: {notes.strip()}")

    return recommendations


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

    ws_dash["G6"] = "Written Recommendations / What-Ifs"
    ws_dash["G6"].font = Font(size=14, bold=True)

    rec_row = 7
    for rec in recommendations:
        ws_dash[f"G{rec_row}"] = f"• {rec}"
        ws_dash[f"G{rec_row}"].alignment = Alignment(wrap_text=True, vertical="top")
        ws_dash.merge_cells(start_row=rec_row, start_column=7, end_row=rec_row, end_column=11)
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
        values = [row["Name"], row["Skills"], row["Best Fit"], row["Recommended Task"]]

        for c, value in enumerate(values, 1):
            cell = ws_dash.cell(r, c)
            cell.value = value
            cell.border = border
            cell.alignment = Alignment(horizontal="center")

            if r % 2 == 0:
                cell.fill = PatternFill("solid", fgColor=light_blue)

    chart_start = 28

    bar = BarChart()
    bar.title = "Needed vs Assigned"
    bar.y_axis.title = "Workers"
    bar.x_axis.title = "Task"

    data = Reference(ws_dash, min_col=2, max_col=3, min_row=7, max_row=7 + len(summary_table))
    cats = Reference(ws_dash, min_col=1, min_row=8, max_row=7 + len(summary_table))

    bar.add_data(data, titles_from_data=True)
    bar.set_categories(cats)
    bar.height = 9
    bar.width = 15
    bar.legend.position = "r"

    ws_dash.add_chart(bar, "E28")

    pie = PieChart()
    pie.title = "Assigned Labor Distribution"

    pie_data = Reference(ws_dash, min_col=3, min_row=7, max_row=7 + len(summary_table))
    pie_cats = Reference(ws_dash, min_col=1, min_row=8, max_row=7 + len(summary_table))

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


st.sidebar.header("Daily Inputs")

day = st.sidebar.selectbox(
    "Day",
    ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
)

shift = st.sidebar.selectbox("Shift", ["1st", "2nd"])

total_cases = st.sidebar.number_input(
    "Total Cases for Today",
    min_value=0,
    step=1,
    value=0
)

hours_remaining = st.sidebar.number_input(
    "Hours Remaining in Shift",
    min_value=0.0,
    step=0.25,
    value=8.0
)

total_outbound_loads_day = st.sidebar.number_input(
    "Total Outbound Loads for the Day",
    min_value=0,
    step=1,
    value=0
)

crossroads_open = st.sidebar.selectbox("Crossroads plant open?", ["YES", "NO"])
deer_creek_open = st.sidebar.selectbox("Deer Creek plant open?", ["YES", "NO"])
msb_open = st.sidebar.selectbox("MSB plant open?", ["YES", "NO"])

present_workers = st.sidebar.multiselect(
    "Who is present?",
    names
)

notes = st.sidebar.text_area("Operations Notes")

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
        msb_open
    )

    staff = pd.read_excel(
        working_file,
        sheet_name="Staffing sheet 1ST Shift",
        usecols="A,D,F,I,T"
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
        notes
    )

    wb = load_workbook(working_file)

    write_recommendations_to_excel(wb, staff)

    build_dashboard(
        wb,
        summary_table,
        present_recommendations,
        recommendations
    )

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("Staffing report generated successfully.")

    st.subheader("Staffing Summary")
    st.dataframe(summary_table, use_container_width=True)

    st.subheader("Recommended Staffing Board")
    st.dataframe(
        present_recommendations[["Name", "Skills", "Best Fit", "Recommended Task"]].reset_index(drop=True),
        use_container_width=True
    )

    st.subheader("Written Recommendations / What-Ifs")
    for rec in recommendations:
        st.write(f"• {rec}")

    st.download_button(
        label="Download Staffing Report",
        data=output,
        file_name="Staffing Report Generated.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
