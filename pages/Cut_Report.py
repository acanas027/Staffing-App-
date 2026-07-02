"""
Cuts / Shorts From Loads — Rep Email Generator
------------------------------------------------
Upload the weekly "SHORTS FROM LOADS" workbook and this app will:
  1. Read the 1ST SHIFT CUTS and 2ND SHIFT CUTS tabs
  2. Match each customer to their CS Rep + email using the
     CUSTOMER SERVICE MASTER LIST tab
  3. Build ONE email per rep containing every affected line item
  4. Let you download ready-to-send .eml files (opens directly in
     Outlook / Mail, addressed and formatted, just hit Send)

Run with:  streamlit run cuts_email_generator.py
"""

import io
import zipfile
from html import escape

import openpyxl
import pandas as pd
import streamlit as st
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# --------------------------------------------------------------------------
# CONFIG — adjust here if sheet names / sender address ever change
# --------------------------------------------------------------------------
SHIFT_SHEETS = ["1ST SHIFT CUTS", "2ND SHIFT CUTS"]
MASTER_SHEET = "CUSTOMER SERVICE MASTER LIST"
FROM_EMAIL = "customerservice@resers.com"  # <-- change to your real sender

st.set_page_config(page_title="Cuts / Shorts Rep Email Generator", layout="wide")
st.title("📧 Cuts / Shorts From Loads — Rep Email Generator")
st.caption(
    "Upload the workbook → one ready-to-send email is built per Customer "
    "Service rep, with a table of every affected item and a Thank you note."
)

uploaded = st.file_uploader("Upload the workbook (.xlsx)", type=["xlsx"])


# --------------------------------------------------------------------------
# PARSING HELPERS
# --------------------------------------------------------------------------
def build_rep_directory(wb):
    """
    Reads the CUSTOMER SERVICE MASTER LIST sheet.
    Layout: col A = rep name (section header, sparse), col B = rep name
    (repeated per row), col C = customer name, col D = email (only on the
    first row of each rep's color block).
    Returns (customer_to_rep dict, rep_to_email dict).
    """
    ws = wb[MASTER_SHEET]
    customer_to_rep = {}
    rep_to_email = {}
    current_rep = None

    for row in ws.iter_rows(min_row=1, max_row=ws.max_row):
        rep_cell = row[1].value if len(row) > 1 else None   # col B
        cust_cell = row[2].value if len(row) > 2 else None  # col C
        email_cell = row[3].value if len(row) > 3 else None  # col D

        if rep_cell:
            current_rep = str(rep_cell).strip()

        if cust_cell and current_rep:
            customer_to_rep[str(cust_cell).strip().upper()] = current_rep

        if email_cell and current_rep and current_rep not in rep_to_email:
            rep_to_email[current_rep] = str(email_cell).strip()

    return customer_to_rep, rep_to_email


def extract_shift_rows(wb, sheet_name):
    """
    Reads one shift-cuts sheet and returns a list of line-item dicts.

    Column layout (1-indexed): A=CS REP, B=CUSTOMER, C=TRIP//LOAD#,
    D=ORDER NUMBER, E=ITEM NUMBER, F=DESCRIPTION, G=QUANTITY CASES CUT,
    H=REASON CODE, I=REASON DESCRIPTION.

    LOAD# forward-fills down (a single physical load can carry several
    orders, and the load number is only typed once). CUSTOMER forward-
    fills ONLY while the ORDER NUMBER stays the same as the row above
    (true multi-line continuation of one order) -- it is NOT carried
    across a change in order number, since that would be guessing a
    customer that was never actually entered. Rows where a new order
    number appears with no customer typed in are labelled "UNKNOWN"
    and surfaced separately so they can be fixed in the source file.
    """
    ws = wb[sheet_name]

    header_row_idx = None
    for r in range(1, ws.max_row + 1):
        val = ws.cell(row=r, column=1).value
        if val and "CS REP" in str(val).upper():
            header_row_idx = r
            break
    if header_row_idx is None:
        return []

    rows = []
    last_load = None
    last_customer = None
    last_order = None

    for r in range(header_row_idx + 1, ws.max_row + 1):
        customer = ws.cell(row=r, column=2).value
        load = ws.cell(row=r, column=3).value
        order_no = ws.cell(row=r, column=4).value
        item_no = ws.cell(row=r, column=5).value
        desc = ws.cell(row=r, column=6).value
        qty = ws.cell(row=r, column=7).value
        reason_code = ws.cell(row=r, column=8).value
        reason_desc = ws.cell(row=r, column=9).value

        # Skip fully blank / padding rows
        if order_no in (None, "") and item_no in (None, "") and not customer and not load:
            continue

        if load:
            last_load = load

        if customer:
            last_customer = str(customer).strip()
        elif order_no != last_order:
            # New order number with no customer typed in -> genuinely unknown
            last_customer = None

        last_order = order_no

        if order_no in (None, ""):
            continue  # not a real line item

        rows.append(
            {
                "Shift": sheet_name.replace(" CUTS", "").title(),
                "CUSTOMER": last_customer or "UNKNOWN",
                "LOAD": last_load,
                "ORDER NUMBER": order_no,
                "ITEM NUMBER": item_no,
                "DESCRIPTION": desc if desc not in (None, "") else "",
                "QUANTITY CASES CUT": qty,
                "REASON CODE": reason_code,
                "REASON DESCRIPTION": reason_desc,
            }
        )

    return rows


def build_html_table(df):
    cols = [
        "CUSTOMER", "LOAD", "ORDER NUMBER", "ITEM NUMBER", "DESCRIPTION",
        "QUANTITY CASES CUT", "REASON CODE", "REASON DESCRIPTION",
    ]
    html = [
        '<table cellspacing="0" cellpadding="0" style="border-collapse:collapse;'
        'font-family:Calibri,Arial,sans-serif;font-size:13px;">'
    ]
    html.append("<tr>")
    for c in cols:
        html.append(
            f'<th style="border:1px solid #999999;background:#4472C4;color:#ffffff;'
            f'padding:4px 8px;text-align:left;">{escape(c)}</th>'
        )
    html.append("</tr>")

    for _, row in df.iterrows():
        html.append("<tr>")
        for c in cols:
            val = row[c]
            val = "" if pd.isna(val) else val
            html.append(
                f'<td style="border:1px solid #999999;padding:4px 8px;">{escape(str(val))}</td>'
            )
        html.append("</tr>")
    html.append("</table>")
    return "".join(html)


def build_eml(to_email, subject, html_table, from_email=FROM_EMAIL):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = subject
    msg["From"] = from_email
    msg["To"] = to_email if to_email else ""
    full_html = (
        f"<html><body style='font-family:Calibri,Arial,sans-serif;font-size:13px;'>"
        f"{html_table}<p style='margin-top:16px;'>Thank you.</p></body></html>"
    )
    msg.attach(MIMEText(full_html, "html"))
    return msg.as_bytes()


# --------------------------------------------------------------------------
# MAIN APP
# --------------------------------------------------------------------------
if uploaded:
    wb = openpyxl.load_workbook(uploaded, data_only=True)

    missing = [s for s in SHIFT_SHEETS + [MASTER_SHEET] if s not in wb.sheetnames]
    if missing:
        st.error(f"Missing expected sheet(s) in this workbook: {missing}")
        st.stop()

    customer_to_rep, rep_to_email = build_rep_directory(wb)

    all_rows = []
    for sheet in SHIFT_SHEETS:
        all_rows.extend(extract_shift_rows(wb, sheet))

    if not all_rows:
        st.warning("No line items found in the shift cuts sheets.")
        st.stop()

    df = pd.DataFrame(all_rows)
    df["REP"] = df["CUSTOMER"].apply(
        lambda c: customer_to_rep.get(str(c).strip().upper())
    )
    df["EMAIL"] = df["REP"].map(rep_to_email)

    matched = df[df["REP"].notna()].copy()
    unmatched = df[df["REP"].isna()].copy()

    st.subheader(
        f"Found {len(df)} affected line item(s) — "
        f"{len(matched)} matched to a rep, {len(unmatched)} need review"
    )

    if not unmatched.empty:
        st.warning(
            f"{len(unmatched)} line item(s) couldn't be matched to a rep "
            "(customer is UNKNOWN or not found in the master list). "
            "No email is generated for these — fix the CUSTOMER column in "
            "the source file and re-upload, or review manually below."
        )
        st.dataframe(
            unmatched[
                ["Shift", "CUSTOMER", "LOAD", "ORDER NUMBER", "ITEM NUMBER", "DESCRIPTION"]
            ],
            use_container_width=True,
        )

    st.markdown("---")
    st.subheader("Generated emails")

    if matched.empty:
        st.info("No rows matched a rep yet — nothing to generate.")
    else:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            for rep, group in matched.groupby("REP"):
                email_addr = rep_to_email.get(rep, "")
                order_numbers = sorted({str(x) for x in group["ORDER NUMBER"]})
                customers = sorted({str(x) for x in group["CUSTOMER"]})
                loads = sorted({str(x) for x in group["LOAD"] if x not in (None, "")})

                subject = (
                    f"{', '.join(order_numbers)} - CUT - {', '.join(customers)} "
                    f"- Load# {', '.join(loads) if loads else 'N/A'}"
                )
                html_table = build_html_table(group)
                safe_name = rep.replace(" ", "_").replace("/", "-")

                with st.expander(
                    f"✉️  {rep}  —  {email_addr or '⚠️ NO EMAIL ON FILE'}  "
                    f"({len(group)} item{'s' if len(group) != 1 else ''})"
                ):
                    st.text_input("To", value=email_addr, disabled=True, key=f"to_{rep}")
                    st.text_input("Subject", value=subject, key=f"subj_{rep}")
                    st.markdown(html_table, unsafe_allow_html=True)
                    st.write("Thank you.")

                    eml_bytes = build_eml(email_addr, subject, html_table)
                    st.download_button(
                        "Download .eml (ready to send)",
                        data=eml_bytes,
                        file_name=f"{safe_name}_cuts_email.eml",
                        mime="message/rfc822",
                        key=f"dl_{rep}",
                    )
                    zf.writestr(f"{safe_name}_cuts_email.eml", eml_bytes)

        st.markdown("---")
        st.download_button(
            "⬇️ Download ALL emails as .zip",
            data=zip_buffer.getvalue(),
            file_name="cuts_emails.zip",
            mime="application/zip",
        )
else:
    st.info("Upload the workbook above to get started.")
