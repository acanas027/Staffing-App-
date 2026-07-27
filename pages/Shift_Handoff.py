from datetime import date
from html import escape
from io import BytesIO
from urllib.parse import quote

import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


REPORT_STATE_KEY = "shift_handoff_report_v4"

CHECKLIST_ITEMS = [
    "All shorts up to the next shift start time are cut or accounted for.",
    "Every trailer with short product is on a door or documented.",
    "Every short that was cut is documented.",
    "Revision emails are checked and addressed.",
    "Inbound and outbound board are updated.",
    "Every drop load has every necessary piece of information and is updated in Yardview.",
    "Check UKG punches and fix any missed punch.",
    "Send attendance to HR.",
    "Received inbound pallets are put away.",
    "Pickers are correctly logged out of the system.", 
    "Sanitation tasks are completed.",
    "Equipment Checklist is reviewed and signed.", 
]

EMAIL_RECIPIENTS = [
    "JuanB@resers.com",
    "CCameron@Resers.com",
    "MarkeithF@Resers.com",
    "CarmenD@Resers.com",
    "BrianM@resers.com",
]


st.set_page_config(
    page_title="Shift Handoff",
    layout="wide",
    initial_sidebar_state="collapsed",
)


st.markdown(
    """
    <style>
        .stApp {
            background:
                radial-gradient(circle at 8% 8%, rgba(62, 207, 142, .13), transparent 27%),
                radial-gradient(circle at 92% 4%, rgba(76, 125, 255, .14), transparent 30%),
                #f7f9fc;
        }

        .block-container {
            max-width: 1180px;
            padding-top: 2rem;
            padding-bottom: 4rem;
        }

        .hero {
            padding: 1.65rem 1.8rem;
            border-radius: 22px;
            color: white;
            background: linear-gradient(120deg, #12355b 0%, #167b75 58%, #39a96b 100%);
            box-shadow: 0 14px 34px rgba(18, 53, 91, .18);
            margin-bottom: 1.2rem;
        }

        .hero h1 {
            margin: 0;
            font-size: clamp(1.8rem, 4vw, 2.8rem);
            letter-spacing: -.04em;
        }

        .hero p {
            margin: .45rem 0 0;
            color: rgba(255, 255, 255, .88);
            font-size: 1.02rem;
        }

        .section-banner {
            display: flex;
            align-items: center;
            gap: .7rem;
            margin: 1.4rem 0 .8rem;
            font-size: 1.16rem;
            font-weight: 800;
            color: #12355b;
        }

        .section-number {
            display: inline-grid;
            place-items: center;
            width: 30px;
            height: 30px;
            border-radius: 10px;
            color: white;
            background: linear-gradient(135deg, #3976e8, #21a179);
            font-size: .9rem;
        }

        div[data-testid="stForm"] {
            background: rgba(255, 255, 255, .85);
            border: 1px solid rgba(18, 53, 91, .09);
            border-radius: 22px;
            padding: 1.1rem 1.35rem 1.35rem;
            box-shadow: 0 12px 32px rgba(18, 53, 91, .07);
        }

        div[data-testid="stNumberInput"],
        div[data-testid="stTextInput"],
        div[data-testid="stDateInput"],
        div[data-testid="stSelectbox"],
        div[data-testid="stTextArea"] {
            border-radius: 14px;
        }

        div[data-testid="stCheckbox"] {
            min-height: 3.15rem;
            padding: .55rem .7rem;
            border: 1px solid #e2e8f0;
            border-radius: 12px;
            background: #fbfcfe;
        }

        div.stButton > button,
        div[data-testid="stFormSubmitButton"] > button,
        div[data-testid="stDownloadButton"] > button {
            border: 0;
            border-radius: 12px;
            font-weight: 750;
            min-height: 2.8rem;
        }

        div[data-testid="stFormSubmitButton"] > button {
            color: white;
            background: linear-gradient(100deg, #236bd6, #16a36d);
            box-shadow: 0 8px 18px rgba(35, 107, 214, .22);
        }

        .result-head {
            display: flex;
            justify-content: space-between;
            align-items: center;
            gap: 1rem;
            padding: 1.15rem 1.3rem;
            margin: 1.8rem 0 .8rem;
            background: #12355b;
            color: white;
            border-radius: 18px;
        }

        .result-head h2 {
            margin: 0;
            font-size: 1.45rem;
        }

        .status-pill {
            display: inline-block;
            padding: .36rem .72rem;
            border-radius: 999px;
            background: #dff8e9;
            color: #09643c;
            font-size: .82rem;
            font-weight: 850;
            white-space: nowrap;
        }

        .status-pill.attention {
            background: #fff1cc;
            color: #815400;
        }

        .status-summary {
            display: flex;
            flex-direction: column;
            align-items: flex-end;
            gap: .3rem;
            text-align: right;
        }

        .status-reason {
            max-width: 300px;
            color: #fff1cc;
            font-size: .78rem;
            line-height: 1.25;
        }

        .issue-card {
            min-height: 125px;
            padding: 1rem 1.05rem;
            border: 1px solid #dfe6ef;
            border-left: 6px solid #24a26f;
            border-radius: 15px;
            background: white;
            box-shadow: 0 6px 16px rgba(18, 53, 91, .05);
        }

        .issue-card.attention {
            border-left-color: #e5a11b;
            background: #fffdf7;
        }

        .issue-card h4 {
            color: #12355b;
            margin: 0 0 .45rem;
        }

        .issue-card p {
            color: #485b70;
            margin: 0;
            line-height: 1.5;
        }

        [data-testid="stMetric"] {
            background: white;
            border: 1px solid #e2e8f0;
            border-radius: 16px;
            padding: .85rem 1rem;
            box-shadow: 0 6px 16px rgba(18, 53, 91, .05);
        }

        .small-note {
            color: #65758a;
            font-size: .88rem;
            margin-top: -.3rem;
        }

        .operation-group-title {
            display: flex;
            align-items: center;
            gap: .55rem;
            margin: 0 0 .2rem;
            color: #12355b;
            font-size: .98rem;
            font-weight: 800;
        }

        .operation-group-title::before {
            content: "";
            width: 8px;
            height: 8px;
            border-radius: 999px;
            background: linear-gradient(135deg, #3976e8, #21a179);
            box-shadow: 0 0 0 4px rgba(57, 118, 232, .10);
        }

        .operation-group-note {
            color: #718096;
            font-size: .78rem;
            line-height: 1.35;
            margin: 0 0 .75rem;
        }

        div[data-testid="stVerticalBlockBorderWrapper"] {
            border-color: #e2e8f0 !important;
            border-radius: 16px !important;
            background: rgba(251, 252, 254, .72);
            box-shadow: 0 4px 12px rgba(18, 53, 91, .035);
        }

        .checklist-summary {
            margin-top: 1rem;
            padding: 1rem 1.05rem;
            border: 1px solid #dfe6ef;
            border-left: 6px solid #3976e8;
            border-radius: 15px;
            background: white;
            box-shadow: 0 6px 16px rgba(18, 53, 91, .05);
        }

        .checklist-summary h4 {
            color: #12355b;
            margin: 0 0 .65rem;
        }

        .checklist-summary p {
            color: #485b70;
            margin: .24rem 0;
            line-height: 1.4;
        }

        .supervisor-notes {
            margin-top: 1rem;
            padding: 1rem 1.05rem;
            border: 1px solid #dfe6ef;
            border-left: 6px solid #167b75;
            border-radius: 15px;
            background: white;
            box-shadow: 0 6px 16px rgba(18, 53, 91, .05);
        }

        .supervisor-notes h4 {
            color: #12355b;
            margin: 0 0 .45rem;
        }

        .supervisor-notes p {
            color: #485b70;
            margin: 0;
            line-height: 1.5;
            white-space: pre-wrap;
        }
    </style>
    """,
    unsafe_allow_html=True,
)


def clean_detail(status: str, details: str, clear_message: str) -> str:
    """Return a consistent description for the submitted handoff."""
    return details.strip() if status == "Issue to hand off" else clear_message


def get_attention_reasons(report: dict) -> list[str]:
    """Return concise reasons the receiving supervisor should review."""
    checklist = report.get("checklist", [])
    incomplete_count = sum(not item["completed"] for item in checklist)
    reasons = []

    loads_waiting = report.get("loads_waiting", 0)
    if loads_waiting > 0:
        load_word = "load" if loads_waiting == 1 else "loads"
        reasons.append(f"{loads_waiting:,} {load_word} waiting on product")
    if report["safety_status"] == "Issue to hand off":
        reasons.append("Safety issue to hand off")
    if report["equipment_status"] == "Issue to hand off":
        reasons.append("Equipment issue to hand off")
    if incomplete_count:
        item_word = "item" if incomplete_count == 1 else "items"
        reasons.append(f"{incomplete_count} checklist {item_word} incomplete")

    return reasons


def determine_status(report: dict) -> tuple[str, str]:
    """Flag operational follow-up, explicit issues, or unfinished checklist items."""
    return (
        ("ATTENTION ITEMS", "attention")
        if get_attention_reasons(report)
        else ("CLEAR HANDOFF", "")
    )


def make_result_header_html(report: dict) -> str:
    """Build continuous HTML so Streamlit never renders a closing tag as text."""
    status_label, status_class = determine_status(report)
    attention_reasons = get_attention_reasons(report)
    status_reason_html = (
        f'<div class="status-reason">{escape(" · ".join(attention_reasons))}</div>'
        if attention_reasons
        else ""
    )
    return (
        '<div class="result-head">'
        "<div>"
        "<h2>Ready for the next supervisor</h2>"
        f"<div>{escape(report['shift'])} · "
        f"{escape(report['report_date'])} · "
        f"{escape(report['supervisor'])}</div>"
        "</div>"
        '<div class="status-summary">'
        f'<span class="status-pill {status_class}">{escape(status_label)}</span>'
        f"{status_reason_html}"
        "</div>"
        "</div>"
    )


def make_checklist_text(report: dict) -> str:
    """Format the checklist consistently for copied text and email."""
    checklist = report.get("checklist", [])
    completed_count = sum(item["completed"] for item in checklist)
    lines = [
        f"Completed: {completed_count} of {len(checklist)}",
    ]
    lines.extend(
        f"[{'X' if item['completed'] else ' '}] {item['item']}"
        for item in checklist
    )
    return "\n".join(lines)


def make_text_report(report: dict) -> str:
    supervisor_notes = (
        report.get("supervisor_notes", "").strip()
        or "No additional supervisor notes."
    )
    staffing_change_details = (
        report.get("staffing_change_details", "").strip()
        or "No staffing changes reported."
    )
    status_label, _ = determine_status(report)
    attention_reasons = get_attention_reasons(report)
    status_text = status_label
    if attention_reasons:
        status_text += "\n" + "\n".join(
            f"- {reason}" for reason in attention_reasons
        )
    return f"""END-OF-SHIFT SUPERVISOR HANDOFF
Date: {report['report_date']}
Shift: {report['shift']}
Supervisor: {report['supervisor']}

HANDOFF STATUS
{status_text}

OPERATION SNAPSHOT
Outbound loads completed: {report['loads_completed']:,}
Loads waiting on product: {report['loads_waiting']:,}
Open stages: {report['open_stages']:,}
Outbound drivers checked in and currently in lot: {report['outbound_drivers_in_lot']:,}
Inbound drivers checked in and currently in lot: {report['inbound_drivers_in_lot']:,}
Cases picked: {report['cases_picked']:,}
Full Pallet Pull Cases: {report['full_pallet_pull_cases']:,}
Staffing at beginning of shift: {report['staffing_beginning']:,}
Staffing at end of shift: {report['staffing_end']:,}

STAFFING CHANGES
{staffing_change_details}

SAFETY
{report['safety_detail']}

EQUIPMENT
{report['equipment_detail']}

SUPERVISOR NOTES
{supervisor_notes}

SHIFT COMPLETION CHECKLIST
{make_checklist_text(report)}
"""


def make_email_url(text_report: str) -> str:
    """Open an Outlook Web draft with resolved, separate To recipients."""
    recipients = quote(",".join(EMAIL_RECIPIENTS), safe="@")
    return (
        "https://outlook.office.com/mail/deeplink/compose?to="
        + recipients
        + "&subject="
        + quote("Shift Handoff", safe="")
        + "&body="
        + quote(text_report, safe="")
    )


def make_pdf_report(report: dict) -> bytes:
    """Build a clean, compact PDF version of the handoff."""
    buffer = BytesIO()
    document = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.55 * inch,
        leftMargin=0.55 * inch,
        topMargin=0.38 * inch,
        bottomMargin=0.38 * inch,
    )

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        "HandoffTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=22,
        leading=26,
        textColor=colors.white,
        spaceAfter=5,
    )
    meta_style = ParagraphStyle(
        "HandoffMeta",
        parent=styles["BodyText"],
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#E6F2F2"),
    )
    section_style = ParagraphStyle(
        "SectionTitle",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=13,
        leading=16,
        textColor=colors.HexColor("#12355B"),
        spaceAfter=8,
    )
    label_style = ParagraphStyle(
        "MetricLabel",
        parent=styles["BodyText"],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#607086"),
    )
    value_style = ParagraphStyle(
        "MetricValue",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=18,
        textColor=colors.HexColor("#12355B"),
    )
    body_style = ParagraphStyle(
        "IssueBody",
        parent=styles["BodyText"],
        fontSize=9.5,
        leading=13,
        textColor=colors.HexColor("#42576D"),
    )
    checklist_style = ParagraphStyle(
        "ChecklistBody",
        parent=styles["BodyText"],
        fontSize=8.2,
        leading=9.8,
        textColor=colors.HexColor("#42576D"),
    )
    checklist_status_style = ParagraphStyle(
        "ChecklistStatus",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=8.2,
        leading=9.8,
        textColor=colors.HexColor("#12355B"),
        alignment=1,
    )
    checklist_count_style = ParagraphStyle(
        "ChecklistCount",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#607086"),
        spaceAfter=6,
    )
    status_reason_style = ParagraphStyle(
        "StatusReason",
        parent=styles["BodyText"],
        fontSize=7.3,
        leading=9,
        textColor=colors.HexColor("#FFF1CC"),
        alignment=2,
    )

    status_label, status_class = determine_status(report)
    attention_reasons = get_attention_reasons(report)
    status_reason = " | ".join(attention_reasons)
    status_color = colors.HexColor("#E5A11B" if status_class else "#24A26F")

    header = Table(
        [[
            Paragraph("End-of-Shift Handoff", title_style),
            Paragraph(f"<b>{escape(status_label)}</b>", meta_style),
        ], [
            Paragraph(
                f"{escape(report['shift'])} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"{escape(report['report_date'])} &nbsp;&nbsp;|&nbsp;&nbsp; "
                f"{escape(report['supervisor'])}",
                meta_style,
            ),
            Paragraph(escape(status_reason), status_reason_style),
        ]],
        colWidths=[4.95 * inch, 1.95 * inch],
    )
    header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#12355B")),
        ("BOX", (0, 0), (-1, -1), 0, colors.HexColor("#12355B")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "RIGHT"),
        ("TOPPADDING", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 16),
        ("RIGHTPADDING", (0, 0), (-1, -1), 16),
        ("LINEBELOW", (1, 0), (1, 0), 3, status_color),
    ]))

    metrics = (
        ("Outbound loads completed", report["loads_completed"]),
        ("Loads waiting on product", report["loads_waiting"]),
        ("Open stages", report["open_stages"]),
        ("Outbound drivers currently in lot", report["outbound_drivers_in_lot"]),
        ("Inbound drivers currently in lot", report["inbound_drivers_in_lot"]),
        ("Cases picked", report["cases_picked"]),
        ("Full Pallet Pull Cases", report["full_pallet_pull_cases"]),
        ("Staffing at beginning of shift", report["staffing_beginning"]),
        ("Staffing at end of shift", report["staffing_end"]),
    )
    metric_cells = [
        [Paragraph(escape(label), label_style), Paragraph(f"{value:,}", value_style)]
        for label, value in metrics
    ]
    metric_rows = []
    for index in range(0, len(metric_cells), 2):
        left_metric = metric_cells[index]
        right_metric = metric_cells[index + 1] if index + 1 < len(metric_cells) else ["", ""]
        metric_rows.append(left_metric + right_metric)
    metric_table = Table(metric_rows, colWidths=[2.45 * inch, 0.65 * inch] * 2)
    metric_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F7F9FC")),
        ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#DFE6EF")),
        ("INNERGRID", (0, 0), (-1, -1), 0.6, colors.HexColor("#DFE6EF")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 9),
        ("RIGHTPADDING", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("ALIGN", (3, 0), (3, -1), "RIGHT"),
    ]))

    def issue_cell(title: str, detail: str) -> list:
        safe_detail = escape(detail).replace("\n", "<br/>")
        return [
            Paragraph(f"<b>{escape(title)}</b>", section_style),
            Paragraph(safe_detail, body_style),
        ]

    issues = Table(
        [[issue_cell("Safety", report["safety_detail"]), issue_cell("Equipment", report["equipment_detail"])]],
        colWidths=[3.45 * inch, 3.45 * inch],
    )
    issues.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FBFCFE")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DFE6EF")),
        ("INNERGRID", (0, 0), (-1, -1), 0.7, colors.HexColor("#DFE6EF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))

    supervisor_notes = (
        report.get("supervisor_notes", "").strip()
        or "No additional supervisor notes."
    )
    staffing_change_details = (
        report.get("staffing_change_details", "").strip()
        or "No staffing changes reported."
    )
    safe_staffing_change_details = escape(staffing_change_details).replace("\n", "<br/>")
    staffing_details_table = Table(
        [[
            Paragraph("<b>Staffing changes</b>", section_style),
            Paragraph(safe_staffing_change_details, body_style),
        ]],
        colWidths=[1.45 * inch, 5.45 * inch],
    )
    staffing_details_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FBFCFE")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DFE6EF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))

    safe_supervisor_notes = escape(supervisor_notes).replace("\n", "<br/>")
    notes_table = Table(
        [[
            Paragraph("<b>Supervisor notes</b>", section_style),
            Paragraph(safe_supervisor_notes, body_style),
        ]],
        colWidths=[1.45 * inch, 5.45 * inch],
    )
    notes_table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FBFCFE")),
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DFE6EF")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("RIGHTPADDING", (0, 0), (-1, -1), 12),
    ]))

    checklist = report.get("checklist", [])
    completed_count = sum(item["completed"] for item in checklist)
    checklist_rows = [
        [
            Paragraph("[X]" if item["completed"] else "[ ]", checklist_status_style),
            Paragraph(escape(item["item"]), checklist_style),
        ]
        for item in checklist
    ]
    checklist_table = Table(
        checklist_rows,
        colWidths=[0.48 * inch, 6.42 * inch],
    )
    checklist_table_style = [
        ("BOX", (0, 0), (-1, -1), 0.7, colors.HexColor("#DFE6EF")),
        ("INNERGRID", (0, 0), (-1, -1), 0.45, colors.HexColor("#E8EDF3")),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
    ]
    for row_index, item in enumerate(checklist):
        row_fill = "#F0FAF5" if item["completed"] else "#FFF9E8"
        checklist_table_style.append(
            ("BACKGROUND", (0, row_index), (-1, row_index), colors.HexColor(row_fill))
        )
    checklist_table.setStyle(TableStyle(checklist_table_style))

    story = [
        header,
        Spacer(1, 0.12 * inch),
        Paragraph("Operation snapshot", section_style),
        metric_table,
        Spacer(1, 0.12 * inch),
        staffing_details_table,
        Spacer(1, 0.10 * inch),
        issues,
        Spacer(1, 0.10 * inch),
        notes_table,
        Spacer(1, 0.12 * inch),
        Paragraph("Shift completion checklist", section_style),
        Paragraph(
            f"{completed_count} of {len(checklist)} items completed",
            checklist_count_style,
        ),
        checklist_table,
    ]
    document.build(story)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


st.markdown(
    """
    <div class="hero">
        <h1>Shift Handoff</h1>
        <p>A five-minute operation snapshot for the supervisor taking over.</p>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.form("shift_handoff_form", border=False):
    st.markdown(
        '<div class="section-banner"><span class="section-number">1</span> Who is handing off?</div>',
        unsafe_allow_html=True,
    )

    identity_1, identity_2, identity_3 = st.columns([1, 1, 1.6], gap="large")
    with identity_1:
        report_date = st.date_input("Report date *", value=date.today())
    with identity_2:
        shift = st.selectbox(
            "Shift *",
            ["Select shift", "1st Shift", "2nd Shift"],
        )
    with identity_3:
        supervisor = st.text_input(
            "Supervisor name *",
            placeholder="Enter your name",
        )

    st.markdown(
        '<div class="section-banner"><span class="section-number">2</span> Operation pulse</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="small-note">Enter the current totals. Counts only.</p>',
        unsafe_allow_html=True,
    )

    with st.container(border=True):
        st.markdown(
            """
            <div class="operation-group-title">Load Activity</div>
            <p class="operation-group-note">
                Completed loads, work still waiting, open stages, and ready-to-load status.
            </p>
            """,
            unsafe_allow_html=True,
        )
        load_1, load_2, load_3, load_4, load_5 = st.columns(5, gap="medium")
        with load_1:
            loads_completed = st.number_input(
                "Outbound loads completed",
                min_value=0,
                step=1,
            )
        with load_2:
            Inbounds_completed = st.number_input(
                "Inbound loads completed",
                min_value=0,
                step=1,
            )
        with load_3:
            loads_waiting = st.number_input(
                "Loads waiting on product",
                min_value=0,
                step=1,
            )
        with load_4:
            open_stages = st.number_input(
                "Open stages",
                min_value=0,
                step=1,
            )
        with load_5:
            RTL_up_to = st.text_input(
                "Ready to load up to",
            )

    yard_group, picking_group, staffing_group = st.columns(3, gap="large")

    with yard_group:
        with st.container(border=True):
            st.markdown(
                """
                <div class="operation-group-title">Central lot activity</div>
                <p class="operation-group-note">Drivers currently checked in and in the lot.</p>
                """,
                unsafe_allow_html=True,
            )
            outbound_drivers_in_lot = st.number_input(
                "Outbound drivers checked in and currently in lot",
                min_value=0,
                step=1,
            )
            inbound_drivers_in_lot = st.number_input(
                "Inbound drivers checked in and currently in lot",
                min_value=0,
                step=1,
            )

    with picking_group:
        with st.container(border=True):
            st.markdown(
                """
                <div class="operation-group-title">Floor productivity</div>
                <p class="operation-group-note">Current case-picking and full-pallet volume.</p>
                """,
                unsafe_allow_html=True,
            )
            cases_picked = st.number_input(
                "Cases picked",
                min_value=0,
                step=1,
            )
            full_pallet_pull_cases = st.number_input(
                "Full Pallet Pull Cases",
                min_value=0,
                step=1,
            )

    with staffing_group:
        with st.container(border=True):
            st.markdown(
                """
                <div class="operation-group-title">Staffing</div>
                <p class="operation-group-note">Headcount at the beginning and end of the shift.</p>
                """,
                unsafe_allow_html=True,
            )
            staffing_beginning = st.number_input(
                "Staffing at beginning of shift",
                min_value=0,
                step=1,
            )
            staffing_end = st.number_input(
                "Staffing at end of shift",
                min_value=0,
                step=1,
            )
            staffing_change_details = st.text_area(
                "Who left and why? (optional)",
                placeholder="Specify who left and why, if applicable.",
                height=90,
            )
            

    st.markdown(
        '<div class="section-banner"><span class="section-number">3</span> Safety & equipment</div>',
        unsafe_allow_html=True,
    )

    safety_col, equipment_col = st.columns(2, gap="large")
    with safety_col:
        safety_status = st.selectbox(
            "Safety status *",
            ["No safety issues", "Issue to hand off"],
        )
        safety_details = st.text_area(
            "Safety details",
            placeholder="Required only if there is an issue",
            height=100,
        )
    with equipment_col:
        equipment_status = st.selectbox(
            "Equipment status *",
            ["No equipment issues", "Issue to hand off"],
        )
        equipment_details = st.text_area(
            "Equipment details",
            placeholder="Required only if there is an issue",
            height=100,
        )

    st.markdown(
        '<div class="section-banner"><span class="section-number">4</span> Shift completion checklist</div>',
        unsafe_allow_html=True,
    )
    st.markdown(
        '<p class="small-note">Check each item that has been completed. Unchecked items will remain visible in the handoff report.</p>',
        unsafe_allow_html=True,
    )

    checklist_answers = {}
    checklist_columns = st.columns(2, gap="large")
    for item_index, item_text in enumerate(CHECKLIST_ITEMS):
        with checklist_columns[item_index % 2]:
            checklist_answers[item_text] = st.checkbox(
                item_text,
                key=f"handoff_checklist_{item_index}",
            )

    supervisor_notes = st.text_area(
        "Supervisor notes",
        placeholder="Add any additional context or information for the next supervisor",
        height=110,
    )

    submitted = st.form_submit_button(
        "Create shift handoff",
        use_container_width=True,
    )

if submitted:
    errors = []
    if shift == "Select shift":
        errors.append("Select a shift.")
    if not supervisor.strip():
        errors.append("Enter the supervisor name.")
    if safety_status == "Issue to hand off" and not safety_details.strip():
        errors.append("Add the safety issue details.")
    if equipment_status == "Issue to hand off" and not equipment_details.strip():
        errors.append("Add the equipment issue details.")

    if errors:
        st.error("Please finish these items:\n\n- " + "\n- ".join(errors))
    else:
        st.session_state[REPORT_STATE_KEY] = {
            "report_date": report_date.strftime("%B %d, %Y"),
            "shift": shift,
            "supervisor": supervisor.strip(),
            "loads_completed": int(loads_completed),
            "loads_waiting": int(loads_waiting),
            "open_stages": int(open_stages),
            "outbound_drivers_in_lot": int(outbound_drivers_in_lot),
            "inbound_drivers_in_lot": int(inbound_drivers_in_lot),
            "cases_picked": int(cases_picked),
            "full_pallet_pull_cases": int(full_pallet_pull_cases),
            "staffing_beginning": int(staffing_beginning),
            "staffing_end": int(staffing_end),
            "staffing_change_details": staffing_change_details.strip(),
            "safety_status": safety_status,
            "safety_detail": clean_detail(
                safety_status,
                safety_details,
                "No safety issues reported.",
            ),
            "equipment_status": equipment_status,
            "equipment_detail": clean_detail(
                equipment_status,
                equipment_details,
                "No equipment issues reported.",
            ),
            "supervisor_notes": supervisor_notes.strip(),
            "checklist": [
                {
                    "item": item_text,
                    "completed": bool(checklist_answers[item_text]),
                }
                for item_text in CHECKLIST_ITEMS
            ],
        }
        st.success("Handoff created. Review it below, then copy, email, or download it.")

if REPORT_STATE_KEY in st.session_state:
    report = st.session_state[REPORT_STATE_KEY]

    st.markdown(
        make_result_header_html(report),
        unsafe_allow_html=True,
    )

    metric_row_1 = st.columns(5, gap="medium")
    metric_row_1[0].metric("Outbound loads completed", f"{report['loads_completed']:,}")
    metric_row_1[1].metric("Loads waiting on product", f"{report['loads_waiting']:,}")
    metric_row_1[2].metric("Open stages", f"{report['open_stages']:,}")
    metric_row_1[3].metric(
        "Outbound drivers in lot",
        f"{report['outbound_drivers_in_lot']:,}",
    )
    metric_row_1[4].metric(
        "Inbound drivers in lot",
        f"{report['inbound_drivers_in_lot']:,}",
    )

    metric_row_2 = st.columns(4, gap="medium")
    metric_row_2[0].metric("Cases picked", f"{report['cases_picked']:,}")
    metric_row_2[1].metric("Full Pallet Pull Cases", f"{report['full_pallet_pull_cases']:,}")
    metric_row_2[2].metric("Staffing at beginning", f"{report['staffing_beginning']:,}")
    metric_row_2[3].metric("Staffing at end", f"{report['staffing_end']:,}")

    staffing_change_details_display = (
        report.get("staffing_change_details", "").strip()
        or "No staffing changes reported."
    )
    st.markdown(
        f"""
        <div class="supervisor-notes">
            <h4>Staffing changes</h4>
            <p>{escape(staffing_change_details_display)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    safety_class = "attention" if report["safety_status"] == "Issue to hand off" else ""
    equipment_class = "attention" if report["equipment_status"] == "Issue to hand off" else ""
    status_cols = st.columns(2, gap="large")
    with status_cols[0]:
        st.markdown(
            f"""
            <div class="issue-card {safety_class}">
                <h4>Safety</h4>
                <p>{escape(report['safety_detail'])}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with status_cols[1]:
        st.markdown(
            f"""
            <div class="issue-card {equipment_class}">
                <h4>Equipment</h4>
                <p>{escape(report['equipment_detail'])}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    supervisor_notes_display = (
        report.get("supervisor_notes", "").strip()
        or "No additional supervisor notes."
    )
    st.markdown(
        f"""
        <div class="supervisor-notes">
            <h4>Supervisor notes</h4>
            <p>{escape(supervisor_notes_display)}</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    checklist = report.get("checklist", [])
    completed_count = sum(item["completed"] for item in checklist)
    checklist_html = "".join(
        (
            f"<p><strong>[{'X' if item['completed'] else ' '}]</strong> "
            f"{escape(item['item'])}</p>"
        )
        for item in checklist
    )
    st.markdown(
        f"""
        <div class="checklist-summary">
            <h4>Shift completion checklist - {completed_count} of {len(checklist)} completed</h4>
            {checklist_html}
        </div>
        """,
        unsafe_allow_html=True,
    )

    text_report = make_text_report(report)
    pdf_report = make_pdf_report(report)
    email_url = make_email_url(text_report)

    action_1, action_2 = st.columns(2, gap="medium")
    safe_date = report["report_date"].replace(",", "").replace(" ", "-")
    with action_1:
        st.link_button(
            "Open email in Outlook",
            email_url,
            use_container_width=True,
        )
    with action_2:
        st.download_button(
            "Download PDF report",
            data=pdf_report,
            file_name=f"shift-handoff-{safe_date}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
