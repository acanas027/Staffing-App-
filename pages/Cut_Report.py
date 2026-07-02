"""
Cuts / Shorts From Loads — Rep Email Generator
------------------------------------------------
Same pattern as the "Email this report" button on the Shift Closeout page:
a mailto: link that opens your default mail app (Outlook) with the message
pre-filled. No setup, no server-side sending, no download step — you just
click "Open email", review it in Outlook, and press Send yourself.

Upload the weekly "SHORTS FROM LOADS" workbook and this app will:
  1. Read the 1ST SHIFT CUTS and 2ND SHIFT CUTS tabs
  2. Match each customer to their CS Rep + email using the
     CUSTOMER SERVICE MASTER LIST tab
  3. Build ONE email per rep with a table of every affected item
  4. Give you an "Open email" button per rep that opens it in Outlook,
     addressed, subject filled in, table in the body — ready to press Send.

Note: mailto: links only support plain text (no HTML/colors), so the table
is a monospace, column-aligned text table rather than a styled grid — it
will still line up cleanly in Outlook. mailto links also have a practical
length limit (~2000 characters); reps with a very large number of items
may get a link that's too long for some mail clients to open reliably.

Run with:  streamlit run cuts_email_generator.py
"""

import urllib.parse

import openpyxl
import pandas as pd
import streamlit as st

# --------------------------------------------------------------------------
# CONFIG
# --------------------------------------------------------------------------
SHIFT_SHEETS = ["1ST SHIFT CUTS", "2ND SHIFT CUTS"]
MASTER_SHEET = "CUSTOMER SERVICE MASTER LIST"
MAILTO_SAFE_LENGTH = 1800  # links longer than this may fail to open in some clients

st.set_page_config(page_title="Cuts / Shorts Rep Email Generator", layout="wide")
st.title("📧 Cuts / Shorts From Loads — Rep Email Generator")
st.caption(
    "Upload the workbook → one email per Customer Service rep is built. "
    "Click Open email, it opens in Outlook ready to send — just press Send."
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


# --------------------------------------------------------------------------
# EMAIL BUILDING
# --------------------------------------------------------------------------
TABLE_COLS = [
    "CUSTOMER", "LOAD", "ORDER NUMBER", "ITEM NUMBER", "DESCRIPTION",
    "QUANTITY CASES CUT", "REASON CODE", "REASON DESCRIPTION",
]


def build_plain_text_table(df):
    """
    Plain-text item list for the mailto body. Customer/Load/Order repeat across
    multiple line items in the same shipment, so instead of repeating that on
    every item (harder to scan), each distinct Customer/Load/Order combination
    gets one header, with its items listed underneath. Outlook's compose window
    uses a proportional font, so this avoids relying on column-width alignment
    (which drifts) and stays compact enough to fit more items under the mailto
    link's length limit.
    """
    def val(row, col):
        v = row[col]
        return "" if pd.isna(v) else str(v)

    lines = []
    grouped = df.groupby(["CUSTOMER", "LOAD", "ORDER NUMBER"], sort=False, dropna=False)
    for (customer, load, order), sub in grouped:
        header = f"{customer}   |   Load #: {load}   |   Order #: {order}"
        rule = "=" * min(len(header), 60)
        lines.append(rule)
        lines.append(header)
        lines.append(rule)

        for _, row in sub.iterrows():
            reason_code = val(row, "REASON CODE")
            reason_desc = val(row, "REASON DESCRIPTION")
            reason = f"{reason_code} - {reason_desc}" if reason_code and reason_desc else (reason_code or reason_desc)

            lines.append(f"  Item #:      {val(row, 'ITEM NUMBER')}")
            lines.append(f"  Description: {val(row, 'DESCRIPTION')}")
            lines.append(f"  Qty Cut:     {val(row, 'QUANTITY CASES CUT')}")
            lines.append(f"  Reason:      {reason}")
            lines.append("")

    return "\n".join(lines).rstrip()


def build_subject(group):
    order_numbers = sorted({str(x) for x in group["ORDER NUMBER"]})
    customers = sorted({str(x) for x in group["CUSTOMER"]})
    loads = sorted({str(x) for x in group["LOAD"] if x not in (None, "")})
    return (
        f"{', '.join(order_numbers)} - CUT - {', '.join(customers)} "
        f"- Load# {', '.join(loads) if loads else 'N/A'}"
    )


def build_mailto(to_addr, subject, body):
    """Same helper as the Shift Closeout page: builds a mailto: link."""
    params = urllib.parse.urlencode(
        {"subject": subject, "body": body},
        quote_via=urllib.parse.quote,
    )
    return f"mailto:{(to_addr or '').strip()}?{params}"


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
    st.subheader("Emails")

    if matched.empty:
        st.info("No rows matched a rep yet — nothing to generate.")
    else:
        for rep, group in matched.groupby("REP"):
            email_addr = rep_to_email.get(rep, "")
            subject = build_subject(group)
            plain_table = build_plain_text_table(group)
            body = f"{plain_table}\n\nThank you."
            mailto_link = build_mailto(email_addr, subject, body)
            too_long = len(mailto_link) > MAILTO_SAFE_LENGTH

            with st.expander(
                f"✉️  {rep}  —  {email_addr or '⚠️ NO EMAIL ON FILE'}  "
                f"({len(group)} item{'s' if len(group) != 1 else ''})"
            ):
                st.text_input("To", value=email_addr, disabled=True, key=f"to_{rep}")
                st.text_input("Subject", value=subject, disabled=True, key=f"subj_{rep}")
                st.code(body, language=None)

                if not email_addr:
                    st.error("No email on file for this rep — add one to the master list.")
                else:
                    if too_long:
                        st.warning(
                            f"This email has {len(group)} items — the link is "
                            f"{len(mailto_link)} characters, above the ~1800-character "
                            "range some mail clients handle reliably. Try opening it "
                            "anyway; if Outlook doesn't open or the body looks cut off, "
                            "let me know and I'll add a way to split large reps into "
                            "multiple emails."
                        )
                    st.link_button("📤 Open email (ready to send in Outlook)", mailto_link)
