from datetime import date
from html import escape

import streamlit as st


st.set_page_config(
    page_title="Shift Handoff",
    page_icon="🔄",
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
    </style>
    """,
    unsafe_allow_html=True,
)


def clean_detail(status: str, details: str, clear_message: str) -> str:
    """Return a consistent description for the submitted handoff."""
    return details.strip() if status == "Issue to hand off" else clear_message


def determine_status(report: dict) -> tuple[str, str]:
    """Create a simple visual status without asking for another input."""
    has_issue = (
        report["safety_status"] == "Issue to hand off"
        or report["equipment_status"] == "Issue to hand off"
        or report["loads_waiting"] > 0
        or report["open_stages"] > 0
    )
    return ("ATTENTION ITEMS", "attention") if has_issue else ("CLEAR HANDOFF", "")


def make_text_report(report: dict) -> str:
    return f"""END-OF-SHIFT SUPERVISOR HANDOFF
Date: {report['report_date']}
Shift: {report['shift']}
Supervisor: {report['supervisor']}

OPERATION SNAPSHOT
Outbound loads completed: {report['loads_completed']:,}
Loads waiting on product: {report['loads_waiting']:,}
Open stages: {report['open_stages']:,}
Drivers checked in and currently in lot: {report['drivers_in_lot']:,}
Cases picked: {report['cases_picked']:,}
Pallet pulls: {report['pallet_pulls']:,}

SAFETY
{report['safety_detail']}

EQUIPMENT
{report['equipment_detail']}
"""


def make_html_report(report: dict) -> str:
    status_label, status_class = determine_status(report)
    attention_color = "#e5a11b" if status_class else "#24a26f"

    metric_rows = "".join(
        f"<div class='metric'><span>{escape(label)}</span><strong>{value:,}</strong></div>"
        for label, value in (
            ("Outbound loads completed", report["loads_completed"]),
            ("Loads waiting on product", report["loads_waiting"]),
            ("Open stages", report["open_stages"]),
            ("Drivers currently in lot", report["drivers_in_lot"]),
            ("Cases picked", report["cases_picked"]),
            ("Pallet pulls", report["pallet_pulls"]),
        )
    )

    return f"""<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Shift Handoff - {escape(report['report_date'])}</title>
  <style>
    * {{ box-sizing: border-box; }}
    body {{ margin: 0; padding: 36px 18px; background: #f4f7fb; color: #17324d;
            font-family: Arial, Helvetica, sans-serif; }}
    main {{ max-width: 920px; margin: auto; background: white; border-radius: 22px;
            overflow: hidden; box-shadow: 0 16px 45px rgba(18,53,91,.12); }}
    header {{ padding: 30px; color: white;
              background: linear-gradient(120deg, #12355b, #167b75 65%, #39a96b); }}
    h1 {{ margin: 0 0 8px; }}
    header p {{ margin: 0; opacity: .88; }}
    .pill {{ display: inline-block; margin-top: 18px; padding: 7px 12px; border-radius: 999px;
             background: white; color: {attention_color}; font-weight: 800; font-size: 12px; }}
    section {{ padding: 26px 30px; }}
    h2 {{ margin: 0 0 15px; font-size: 19px; }}
    .metrics {{ display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; }}
    .metric {{ padding: 16px; border-radius: 14px; background: #f7f9fc; border: 1px solid #e3e9f0; }}
    .metric span {{ display: block; min-height: 36px; color: #607086; font-size: 13px; }}
    .metric strong {{ font-size: 28px; color: #12355b; }}
    .issues {{ display: grid; grid-template-columns: 1fr 1fr; gap: 14px; }}
    .card {{ padding: 18px; border-radius: 14px; background: #fbfcfe;
             border-left: 6px solid {attention_color}; }}
    .card h3 {{ margin: 0 0 8px; font-size: 16px; }}
    .card p {{ margin: 0; color: #506276; line-height: 1.5; white-space: pre-wrap; }}
    @media (max-width: 650px) {{
      .metrics {{ grid-template-columns: repeat(2, 1fr); }}
      .issues {{ grid-template-columns: 1fr; }}
    }}
    @media print {{ body {{ padding: 0; background: white; }} main {{ box-shadow: none; }} }}
  </style>
</head>
<body>
  <main>
    <header>
      <h1>End-of-Shift Handoff</h1>
      <p>{escape(report['shift'])} · {escape(report['report_date'])} · {escape(report['supervisor'])}</p>
      <div class="pill">{status_label}</div>
    </header>
    <section>
      <h2>Operation snapshot</h2>
      <div class="metrics">{metric_rows}</div>
    </section>
    <section class="issues">
      <div class="card"><h3>Safety</h3><p>{escape(report['safety_detail'])}</p></div>
      <div class="card"><h3>Equipment</h3><p>{escape(report['equipment_detail'])}</p></div>
    </section>
  </main>
</body>
</html>"""


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
            ["Select shift", "1st Shift", "2nd Shift", "3rd Shift"],
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
        '<p class="small-note">Enter the current totals. Counts only—fast and clean.</p>',
        unsafe_allow_html=True,
    )

    pulse_1, pulse_2, pulse_3 = st.columns(3, gap="large")
    with pulse_1:
        loads_completed = st.number_input(
            "✅ Outbound loads completed",
            min_value=0,
            step=1,
        )
        drivers_in_lot = st.number_input(
            "🚚 Drivers checked in and currently in lot",
            min_value=0,
            step=1,
        )
    with pulse_2:
        loads_waiting = st.number_input(
            "⏳ Loads waiting on product",
            min_value=0,
            step=1,
        )
        cases_picked = st.number_input(
            "📦 Cases picked",
            min_value=0,
            step=1,
        )
    with pulse_3:
        open_stages = st.number_input(
            "🏷️ Open stages",
            min_value=0,
            step=1,
        )
        pallet_pulls = st.number_input(
            "🛞 Pallet pulls",
            min_value=0,
            step=1,
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

    submitted = st.form_submit_button(
        "Create shift handoff →",
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
        st.session_state["handoff_report"] = {
            "report_date": report_date.strftime("%B %d, %Y"),
            "shift": shift,
            "supervisor": supervisor.strip(),
            "loads_completed": int(loads_completed),
            "loads_waiting": int(loads_waiting),
            "open_stages": int(open_stages),
            "drivers_in_lot": int(drivers_in_lot),
            "cases_picked": int(cases_picked),
            "pallet_pulls": int(pallet_pulls),
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
        }
        st.success("Handoff created. Review it below, then copy or download it.")

if "handoff_report" in st.session_state:
    report = st.session_state["handoff_report"]
    status_label, status_class = determine_status(report)

    st.markdown(
        f"""
        <div class="result-head">
            <div>
                <h2>Ready for the next supervisor</h2>
                <div>{escape(report['shift'])} · {escape(report['report_date'])} · {escape(report['supervisor'])}</div>
            </div>
            <span class="status-pill {status_class}">{status_label}</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    metric_row_1 = st.columns(3, gap="medium")
    metric_row_1[0].metric("Outbound loads completed", f"{report['loads_completed']:,}")
    metric_row_1[1].metric("Loads waiting on product", f"{report['loads_waiting']:,}")
    metric_row_1[2].metric("Open stages", f"{report['open_stages']:,}")

    metric_row_2 = st.columns(3, gap="medium")
    metric_row_2[0].metric("Drivers currently in lot", f"{report['drivers_in_lot']:,}")
    metric_row_2[1].metric("Cases picked", f"{report['cases_picked']:,}")
    metric_row_2[2].metric("Pallet pulls", f"{report['pallet_pulls']:,}")

    safety_class = "attention" if report["safety_status"] == "Issue to hand off" else ""
    equipment_class = "attention" if report["equipment_status"] == "Issue to hand off" else ""
    status_cols = st.columns(2, gap="large")
    with status_cols[0]:
        st.markdown(
            f"""
            <div class="issue-card {safety_class}">
                <h4>🦺 Safety</h4>
                <p>{escape(report['safety_detail'])}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )
    with status_cols[1]:
        st.markdown(
            f"""
            <div class="issue-card {equipment_class}">
                <h4>🛠️ Equipment</h4>
                <p>{escape(report['equipment_detail'])}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

    text_report = make_text_report(report)
    html_report = make_html_report(report)

    st.markdown("#### Copy or download")
    st.caption("Use the copy icon in the top-right corner of the text box, or download a report.")
    st.code(text_report, language=None)

    download_1, download_2, download_space = st.columns([1, 1, 2], gap="medium")
    safe_date = report["report_date"].replace(",", "").replace(" ", "-")
    with download_1:
        st.download_button(
            "Download text report",
            data=text_report,
            file_name=f"shift-handoff-{safe_date}.txt",
            mime="text/plain",
            use_container_width=True,
        )
    with download_2:
        st.download_button(
            "Download HTML report",
            data=html_report,
            file_name=f"shift-handoff-{safe_date}.html",
            mime="text/html",
            use_container_width=True,
        )
