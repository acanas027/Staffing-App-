"""
Warehouse Case-Picking Leaderboard
==================================
Paste a screenshot OR paste raw text. The app extracts, per row:
    1. name
    2. cases picked
    3. % of total
…then you enter the hours each person actually picked, and it ranks everyone by
CASES PER HOUR (the fair metric). Top 3 (at/above standard) show green; anyone
below the 185 cases/hr standard shows red.

Deploy on Streamlit Community Cloud (files at REPO ROOT):
  - app.py            (this file)
  - requirements.txt  (python deps)
  - packages.txt      (optional; only needed for screenshot OCR -> tesseract-ocr)
"""

import io
import re
from datetime import datetime

# Central Time clock for report timestamps. America/Chicago auto-handles
# CST (winter) vs CDT (summer); a fixed -6 offset would be an hour off half the year.
# Fallback to fixed CST only if the tz database isn't available on the host.
try:
    from zoneinfo import ZoneInfo
    _CENTRAL = ZoneInfo("America/Chicago")
except Exception:
    from datetime import timezone, timedelta
    _CENTRAL = timezone(timedelta(hours=-6))  # CST (no DST) fallback


def central_now() -> datetime:
    return datetime.now(_CENTRAL)

import pandas as pd
import streamlit as st

# ---- Optional deps (app still runs if these are missing) --------------------
try:
    from PIL import Image, ImageOps
    PIL_OK = True
except Exception:
    PIL_OK = False

try:
    import pytesseract
    TESS_OK = True
except Exception:
    TESS_OK = False

try:
    from streamlit_paste_button import paste_image_button
    PASTE_OK = True
except Exception:
    PASTE_OK = False

# ---- PDF (reportlab: pure-python, no system packages needed) -----------------
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.graphics.shapes import Drawing, Rect, String, Line


def tesseract_engine_ready() -> bool:
    if not TESS_OK:
        return False
    try:
        pytesseract.get_tesseract_version()
        return True
    except Exception:
        return False


# =============================================================================
# Parsing  (UNCHANGED)
# =============================================================================
NUMERIC_RE = re.compile(r"^[\$]?[\d.,]+%?$")


def _to_number(tok: str):
    cleaned = tok.replace(",", "").replace("$", "").replace("%", "").strip()
    if cleaned in ("", ".", "-"):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_rows(text: str) -> pd.DataFrame:
    """
    Per line: name = tokens before the first pure-number token (keeps
    digit-names like 'TKTEMP3'); first number = cases, last number = %.
    The throwaway middle column is ignored.
    """
    rows = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue
        name_parts, numbers, seen = [], [], False
        for tok in line.split():
            is_num = bool(NUMERIC_RE.match(tok)) and _to_number(tok) is not None
            if is_num:
                seen = True
                numbers.append(_to_number(tok))
            elif not seen:
                name_parts.append(tok)
        name = " ".join(name_parts).strip()
        if not name or not numbers or not any(c.isalpha() for c in name):
            continue
        rows.append({
            "name": name,
            "cases": numbers[0],
            "pct": numbers[-1] if len(numbers) > 1 else None,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    df["cases"] = df["cases"].fillna(0).astype(int)
    if df["pct"].isna().all():
        total = df["cases"].sum()
        df["pct"] = (df["cases"] / total * 100) if total else 0.0
    df = df.sort_values("cases", ascending=False).reset_index(drop=True)
    df.insert(0, "rank", df.index + 1)
    return df


def ocr_image(file_bytes: bytes) -> str:
    img = Image.open(io.BytesIO(file_bytes))
    img = ImageOps.exif_transpose(img).convert("L")
    if max(img.size) < 1600:
        scale = 1600 / max(img.size)
        img = img.resize((int(img.width * scale), int(img.height * scale)))
    img = ImageOps.autocontrast(img)
    return pytesseract.image_to_string(img, config="--psm 6")


# =============================================================================
# Rate math  (NEW — single source of truth for screen + PDF)
# =============================================================================
STANDARD_RATE = 185   # cases/hr; at/above = OK, below = red flag

# Predetermined hours worked at each point in the shift, by day type.
# SUPERVISORS: change these numbers if the schedule changes — nothing else needs editing.
# Picking these in the app pre-fills every worker's hours; you can still adjust
# individuals (e.g. someone moved off picking early or came in late).
DEFAULT_HOURS = {
    "Weekday": {
        "1st break":   3.0,
        "Lunch":       5.75,
        "2nd break":   7.75,
        "End of day":  9.5,
    },
    "Weekend": {
        "1st break":   3.25,
        "Lunch":       6.5,
        "2nd break":   9.25,
        "End of day":  11.25,
    },
}
DAY_TYPES = ["Weekday", "Weekend"]
BREAK_POINTS = ["1st break", "Lunch", "2nd break", "End of day"]

# Per-worker hours dropdown: every 15 minutes from 0 up to MAX_HOURS.
# All DEFAULT_HOURS values are quarter-hours, so they're guaranteed to be in this list.
MAX_HOURS = 13.0
HOURS_OPTIONS = [round(i * 0.25, 2) for i in range(int(MAX_HOURS / 0.25) + 1)]


def compute_rates(df: pd.DataFrame) -> pd.DataFrame:
    """
    Add cases/hr, re-rank by rate, and tag each row's color group.
    Rules:
      - no hours entered  -> rate = NaN, group 'neutral'
      - rate < STANDARD    -> 'red'  (below standard wins, even if top 3)
      - top 3 by rate (>=standard) -> 'green'
      - everyone else      -> 'neutral'
    Rows without hours sort to the bottom (by cases) so the ranking stays clean.
    """
    d = df.copy()
    if "hours" not in d.columns:
        d["hours"] = 0.0
    d["hours"] = pd.to_numeric(d["hours"], errors="coerce").fillna(0.0)
    d["rate"] = [
        (c / h) if (h and h > 0) else float("nan")
        for c, h in zip(d["cases"], d["hours"])
    ]

    has = d[d["rate"].notna()].copy()
    has["rate"] = has["rate"].astype(float)
    has = has.sort_values("rate", ascending=False)
    no = d[d["rate"].isna()].sort_values("cases", ascending=False)
    d = pd.concat([has, no], ignore_index=True)
    d["rank"] = d.index + 1

    groups = []
    for i, rate in enumerate(d["rate"].tolist()):
        if pd.isna(rate):
            groups.append("neutral")
        elif rate < STANDARD_RATE:
            groups.append("red")
        elif i < 3:
            groups.append("green")
        else:
            groups.append("neutral")
    d["color_group"] = groups
    return d


# =============================================================================
# Print-ready PDF (single combined chart, one page)
# =============================================================================
INK     = colors.HexColor("#1f2933")
STEEL   = colors.HexColor("#3e4c59")
LINE    = colors.HexColor("#cbd2d9")
FAINT   = colors.HexColor("#eef1f4")
HEADER  = colors.HexColor("#243b53")
GREEN   = colors.HexColor("#2f9e44")   # top 3 at/above standard
RED     = colors.HexColor("#e03131")   # below 185/hr standard
BASEBAR = colors.HexColor("#52606d")   # everyone else
ZEBRA   = colors.HexColor("#f5f7fa")


def _bar_color(group):
    return GREEN if group == "green" else (RED if group == "red" else BASEBAR)


def labeled_bar_chart(df: pd.DataFrame, width: float) -> Drawing:
    """One row per picker: rank + name | bar = cases/hr | cases | % of total."""
    n = len(df)
    row_h = 28
    head_h = 18
    height = head_h + n * row_h + 4
    d = Drawing(width, height)

    name_w = 132          # rank + name zone
    cases_w = 56          # right CASES column
    pct_w = 52            # right % OF TOTAL column
    right_w = cases_w + pct_w
    gap = 10
    bar_x = name_w + gap
    bar_max = width - name_w - right_w - 2 * gap
    cases_x = width - pct_w - 4   # right edge of the CASES number
    pct_x = width - 2             # right edge of the % number

    rate_vals = [r for r in df["rate"].tolist() if not pd.isna(r)]
    maxv = max(rate_vals + [STANDARD_RATE]) if rate_vals else STANDARD_RATE
    maxv = max(maxv, 1)

    # ---- column headers ----
    hy = height - 11
    d.add(String(2, hy, "RANK / PICKER", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL))
    d.add(String(bar_x, hy, "CASES / HR", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL))
    d.add(String(cases_x, hy, "CASES", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL, textAnchor="end"))
    d.add(String(pct_x, hy, "% OF TOTAL", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL, textAnchor="end"))
    d.add(Line(0, height - head_h, width, height - head_h,
               strokeColor=LINE, strokeWidth=0.8))

    # ---- rows ----
    for i, (_, r) in enumerate(df.iterrows()):
        row_top = height - head_h - i * row_h
        cy = row_top - row_h / 2
        grp = r["color_group"]
        is_top = grp == "green"

        if i % 2 == 1:
            d.add(Rect(0, row_top - row_h, width, row_h,
                       fillColor=ZEBRA, strokeColor=None))

        bar_color = _bar_color(grp)
        name_font = "Helvetica-Bold" if is_top else "Helvetica"
        text_y = cy - 3.2

        # rank + name
        d.add(String(2, text_y, str(int(r["rank"])), fontName="Helvetica-Bold",
                     fontSize=9, fillColor=(bar_color if grp != "neutral" else STEEL)))
        d.add(String(22, text_y, str(r["name"]), fontName=name_font,
                     fontSize=9.5, fillColor=INK))

        rate = r["rate"]
        bar_h = 15
        by = cy - bar_h / 2
        if pd.isna(rate):
            d.add(String(bar_x, text_y, "— enter hours", fontName="Helvetica-Oblique",
                         fontSize=8, fillColor=STEEL))
        else:
            w = max(bar_max * (rate / maxv), 2)
            d.add(Rect(bar_x, by, w, bar_h, fillColor=bar_color, strokeColor=None))
            rate_str = f"{rate:.0f}/hr"
            if w > 46:
                d.add(String(bar_x + w - 6, text_y, rate_str, fontName="Helvetica-Bold",
                             fontSize=8.5, fillColor=colors.white, textAnchor="end"))
            else:
                d.add(String(bar_x + w + 5, text_y, rate_str, fontName="Helvetica-Bold",
                             fontSize=8.5, fillColor=STEEL, textAnchor="start"))

        # cases + % of total (two right columns)
        d.add(String(cases_x, text_y, f"{int(r['cases']):,}", fontName=name_font,
                     fontSize=9.5, fillColor=INK if is_top else STEEL, textAnchor="end"))
        pct_val = r.get("pct")
        pct_str = f"{float(pct_val):.2f}%" if pct_val is not None and not pd.isna(pct_val) else "—"
        d.add(String(pct_x, text_y, pct_str, fontName=name_font,
                     fontSize=9.5, fillColor=INK if is_top else STEEL, textAnchor="end"))

        # row separator
        d.add(Line(0, row_top - row_h, width, row_top - row_h,
                   strokeColor=FAINT, strokeWidth=0.5))

    # ---- 185 standard marker (drawn on top of the bars) ----
    std_x = bar_x + bar_max * (STANDARD_RATE / maxv)
    d.add(Line(std_x, height - head_h, std_x, 2,
               strokeColor=RED, strokeWidth=0.7, strokeDashArray=[3, 3]))
    d.add(String(std_x, height - head_h + 1, "185", fontName="Helvetica-Bold",
                 fontSize=6, fillColor=RED, textAnchor="middle"))

    return d


def generate_pdf(df: pd.DataFrame, title: str, subtitle: str) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        leftMargin=0.6 * inch, rightMargin=0.6 * inch,
        topMargin=0.55 * inch, bottomMargin=0.5 * inch,
        title=title,
    )
    content_w = doc.width
    styles = getSampleStyleSheet()

    h_title = ParagraphStyle("t", parent=styles["Title"], fontName="Helvetica-Bold",
                             fontSize=20, textColor=colors.white, leading=24, alignment=0)
    h_sub = ParagraphStyle("s", parent=styles["Normal"], fontName="Helvetica",
                           fontSize=10, textColor=colors.HexColor("#bcccdc"), leading=13)
    h_sec = ParagraphStyle("sec", parent=styles["Heading2"], fontName="Helvetica-Bold",
                           fontSize=12, textColor=INK, spaceBefore=12, spaceAfter=6)
    lbl = ParagraphStyle("lbl", fontName="Helvetica", fontSize=7.5, textColor=STEEL, leading=9)
    val = ParagraphStyle("val", fontName="Helvetica-Bold", fontSize=15, textColor=INK, leading=18)

    story = []

    # ---- header band ----
    gen = central_now().strftime("%B %d, %Y  %I:%M %p %Z")
    head = Table([[Paragraph(title, h_title)],
                  [Paragraph(subtitle, h_sub)],
                  [Paragraph(f"Generated {gen}", h_sub)]], colWidths=[content_w])
    head.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HEADER),
        ("LEFTPADDING", (0, 0), (-1, -1), 16),
        ("RIGHTPADDING", (0, 0), (-1, -1), 16),
        ("TOPPADDING", (0, 0), (0, 0), 14),
        ("BOTTOMPADDING", (-1, -1), (-1, -1), 14),
        ("LINEBELOW", (0, -1), (-1, -1), 3, GREEN),
    ]))
    story.append(head)
    story.append(Spacer(1, 12))

    # ---- summary boxes (rate-focused) ----
    rated = df[df["rate"].notna()]
    total = int(df["cases"].sum())
    avg_rate = rated["rate"].mean() if not rated.empty else None
    top_rate = rated["rate"].max() if not rated.empty else None
    below = int((rated["rate"] < STANDARD_RATE).sum()) if not rated.empty else 0

    def rfmt(v):
        return f"{v:,.0f}/hr" if v is not None else "—"

    stats = [("PICKERS", f"{len(df)}"),
             ("AVG CASES/HR", rfmt(avg_rate)),
             ("TOP CASES/HR", rfmt(top_rate)),
             ("BELOW 185", f"{below}")]
    box = Table([[Table([[Paragraph(l, lbl)], [Paragraph(v, val)]]) for l, v in stats]],
                colWidths=[content_w / 4.0] * 4)
    box.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), ZEBRA),
        ("BOX", (0, 0), (-1, -1), 0.5, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 0.5, LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 12),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    story.append(box)

    # ---- the single combined chart ----
    story.append(Paragraph("Leaderboard — ranked by cases / hour", h_sec))
    story.append(labeled_bar_chart(df, content_w))

    foot = ParagraphStyle("foot", fontName="Helvetica", fontSize=7.5,
                          textColor=STEEL, alignment=1, spaceBefore=10)
    story.append(Spacer(1, 4))
    story.append(Paragraph(
        f"Ranked by cases/hr. Green = top 3 at/above the {STANDARD_RATE}/hr standard. "
        f"Red = below {STANDARD_RATE}/hr. Total cases: {total:,}.", foot))

    doc.build(story)
    return buf.getvalue()


# =============================================================================
# On-screen leaderboard (matches the PDF: one labelled chart, rate-first)
# =============================================================================
def render_leaderboard(df: pd.DataFrame, title: str, subtitle: str):
    rated = df[df["rate"].notna()]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pickers", len(df))
    c2.metric("Avg cases/hr", f"{rated['rate'].mean():,.0f}" if not rated.empty else "—")
    c3.metric("Top cases/hr", f"{rated['rate'].max():,.0f}" if not rated.empty else "—")
    c4.metric("Below 185", int((rated["rate"] < STANDARD_RATE).sum()) if not rated.empty else 0)

    st.markdown("##### Leaderboard — ranked by cases / hour")
    plot = df.copy()
    plot["rate_plot"] = plot["rate"].fillna(0)
    plot["label"] = plot.apply(
        lambda r: (f"{r['rate']:.0f}/hr  •  {int(r['cases']):,} cs  •  {float(r['pct']):.2f}%"
                   if not pd.isna(r["rate"]) else "enter hours"),
        axis=1,
    )
    try:
        import altair as alt
        base = alt.Chart(plot).encode(
            y=alt.Y("name:N", sort=alt.SortField("rate_plot", order="descending"),
                    title=None, axis=alt.Axis(labelFontWeight="bold", labelFontSize=12)),
        )
        bars = base.mark_bar(cornerRadiusEnd=2).encode(
            x=alt.X("rate_plot:Q", title="Cases per hour"),
            color=alt.Color("color_group:N",
                            scale=alt.Scale(domain=["green", "red", "neutral"],
                                            range=["#2f9e44", "#e03131", "#52606d"]),
                            legend=None),
            tooltip=["rank", "name", "cases", "pct", "hours", "rate"],
        )
        text = base.mark_text(align="left", dx=4, fontSize=11,
                              fontWeight="bold", color="#1f2933").encode(
            x="rate_plot:Q", text="label:N")
        rule = alt.Chart(pd.DataFrame({"x": [STANDARD_RATE]})).mark_rule(
            color="#e03131", strokeDash=[4, 4], size=1.5).encode(x="x:Q")
        st.altair_chart((bars + text + rule).properties(height=max(240, len(df) * 30)),
                        use_container_width=True)
    except Exception:
        st.bar_chart(df.set_index("name")["rate"].fillna(0), use_container_width=True)

    st.caption(f"Green = top 3 at/above {STANDARD_RATE}/hr  ·  Red = below {STANDARD_RATE}/hr  "
               "·  dashed line marks the standard.")

    pdf_bytes = generate_pdf(df, title or "Case-Picking Leaderboard", subtitle or "")
    st.download_button("Download print-ready PDF (1 page)", data=pdf_bytes,
                       file_name="leaderboard.pdf", mime="application/pdf",
                       type="primary", use_container_width=True)
    st.download_button("Download CSV", data=df.to_csv(index=False).encode("utf-8"),
                       file_name="leaderboard.csv", mime="text/csv",
                       use_container_width=True)


# =============================================================================
# UI
# =============================================================================
st.set_page_config(page_title="Warehouse Leaderboard", layout="centered")
st.title("Warehouse Case-Picking Leaderboard")
st.caption("Paste a screenshot or raw text. The app extracts each picker's name and "
           "cases picked; you enter the hours they actually picked, and it ranks "
           "everyone by cases per hour.")

with st.expander("Report details (shown on the PDF)", expanded=False):
    report_title = st.text_input("Title", value="Case-Picking Leaderboard")
    report_subtitle = st.text_input("Subtitle / shift / date", value="Warehouse Operations")

if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""

tab_text, tab_image = st.tabs(["Paste text", "Paste / upload screenshot"])

with tab_text:
    st.text_area(
        "Paste rows here (one picker per line):",
        height=240,
        placeholder="AIRELH    1,760    0    11.74\nDANIELV   1,659    0    11.07\n...",
        key="raw_text",
    )

with tab_image:
    engine_ready = PIL_OK and tesseract_engine_ready()
    if not engine_ready:
        st.warning(
            "Image OCR isn't available because the tesseract engine isn't installed "
            "on this server. To enable it, add a file named **packages.txt** "
            "containing the single line `tesseract-ocr` at your **repo root** (not "
            "inside `pages/`), commit, then **Manage app -> Reboot**. The Paste text "
            "tab works regardless."
        )
    else:
        img_bytes = None
        if PASTE_OK:
            res = paste_image_button("Paste image from clipboard", errors="ignore")
            if res is not None and getattr(res, "image_data", None) is not None:
                b = io.BytesIO()
                res.image_data.save(b, format="PNG")
                img_bytes = b.getvalue()
        uploaded = st.file_uploader("…or upload / drag a screenshot",
                                    type=["png", "jpg", "jpeg", "webp"])
        if uploaded is not None:
            img_bytes = uploaded.read()
        if img_bytes:
            st.image(img_bytes, caption="Source image", use_container_width=True)
            try:
                with st.spinner("Reading text from image…"):
                    extracted = ocr_image(img_bytes)
                st.text_area("OCR result (edit if needed, then build below):",
                             value=extracted, height=200, key="ocr_box")
            except Exception as e:
                st.error(f"Couldn't read the image ({type(e).__name__}). "
                         "Use the Paste text tab instead.")

st.divider()

# ---- Build: parse names + cases, then collect hours --------------------------
if st.button("Build leaderboard", type="primary", use_container_width=True):
    # Read from the widgets' own session keys (stable across the rerun the click
    # triggers). OCR box wins if it has content, otherwise the pasted text.
    ocr_text = (st.session_state.get("ocr_box") or "").strip()
    paste_text = (st.session_state.get("raw_text") or "").strip()
    text = ocr_text or paste_text

    if not text:
        st.warning("Paste your picker rows in the **Paste text** tab (or OCR a "
                   "screenshot) first, then click Build leaderboard.")
    else:
        parsed = parse_rows(text)
        if parsed.empty:
            st.error("Couldn't find any picker rows in that text. Each line should be "
                     "a name followed by numbers, e.g. `AIRELH  1,760  0  11.74`.")
        else:
            # Store only the parsed identity + cases; hours are entered below.
            st.session_state["picker_df"] = parsed[["rank", "name", "cases", "pct"]].copy()
            # New roster -> let the hours section re-apply the predetermined default.
            st.session_state.pop("applied_combo", None)

# ---- Hours entry + rate leaderboard -----------------------------------------
if "picker_df" in st.session_state:
    base_df = st.session_state["picker_df"]
    cases_map = dict(zip(base_df["name"], base_df["cases"]))
    names = base_df["name"].tolist()

    st.subheader("Enter hours each person picked")

    # --- Shift settings -> predetermined hours for everyone ---
    sc1, sc2 = st.columns(2)
    day_type = sc1.selectbox("Day type", DAY_TYPES, key="day_type")
    break_point = sc2.selectbox("Point in shift", BREAK_POINTS, key="break_point")
    default_hours = float(DEFAULT_HOURS[day_type][break_point])

    st.caption(
        f"Predetermined hours for **{day_type} · {break_point} = {default_hours} h** are filled in "
        "automatically. Lower anyone who was moved off picking or came in late. "
        "Leave at 0 to show no rate for that person."
    )

    # Make sure every worker's hours key exists before the form is built.
    for nm in names:
        st.session_state.setdefault(f"hrs::{nm}", 0.0)

    # Auto-apply: fill everyone with the predetermined hours, but ONLY when the
    # day-type/break-point selection changes (or on first load / new roster).
    # This must run BEFORE the form widgets are instantiated so it sets their values.
    # Re-applying on every rerun would wipe per-worker edits, so we gate on the combo.
    combo = (day_type, break_point)
    if st.session_state.get("applied_combo") != combo:
        for nm in names:
            st.session_state[f"hrs::{nm}"] = default_hours
        st.session_state["applied_combo"] = combo

    with st.form("hours_form"):
        ncols = 3
        for i in range(0, len(names), ncols):
            chunk = names[i:i + ncols]
            cols = st.columns(ncols)
            for col, nm in zip(cols, chunk):
                col.selectbox(
                    f"{nm}  ({int(cases_map[nm]):,} cs)",
                    options=HOURS_OPTIONS,
                    format_func=lambda v: f"{v:.2f} h",
                    key=f"hrs::{nm}",
                )
        st.form_submit_button("Calculate rates", type="primary", use_container_width=True)

    # Build the working frame from whatever hours have been submitted so far.
    work = base_df.copy()
    work["hours"] = work["name"].apply(
        lambda n: float(st.session_state.get(f"hrs::{n}", 0.0) or 0.0)
    )
    work = compute_rates(work)
    render_leaderboard(work, report_title, report_subtitle)
