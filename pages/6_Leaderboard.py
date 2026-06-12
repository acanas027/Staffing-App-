"""
Warehouse Case-Picking Leaderboard
==================================
Paste a screenshot OR paste raw text. The app extracts, per row:
    1. name
    2. cases picked
    3. % of total
…and renders a single-page, print-ready leaderboard where every picker's bar
is labelled with their case count and their % of the total.

Deploy on Streamlit Community Cloud (files at REPO ROOT):
  - app.py            (this file)
  - requirements.txt  (python deps)
  - packages.txt      (optional; only needed for screenshot OCR -> tesseract-ocr)
"""

import io
import re
from datetime import datetime

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
# Parsing
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
# Print-ready PDF (single combined chart, one page)
# =============================================================================
INK    = colors.HexColor("#1f2933")
STEEL  = colors.HexColor("#3e4c59")
LINE   = colors.HexColor("#cbd2d9")
FAINT  = colors.HexColor("#eef1f4")
HEADER = colors.HexColor("#243b53")
ACCENT = colors.HexColor("#b45309")   # amber: top 3 bars
BASEBAR = colors.HexColor("#52606d")  # steel: everyone else
ZEBRA  = colors.HexColor("#f5f7fa")


def labeled_bar_chart(df: pd.DataFrame, width: float) -> Drawing:
    """One row per picker: rank + name | bar with case count | % of total."""
    n = len(df)
    row_h = 28
    head_h = 18
    height = head_h + n * row_h + 4
    d = Drawing(width, height)

    name_w = 132          # rank + name zone
    pct_w = 58            # right % column
    gap = 10
    bar_x = name_w + gap
    bar_max = width - name_w - pct_w - 2 * gap
    maxv = max(df["cases"].max(), 1)

    # ---- column headers ----
    hy = height - 11
    d.add(String(2, hy, "RANK / PICKER", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL))
    d.add(String(bar_x, hy, "CASES PICKED", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL))
    d.add(String(width - 2, hy, "% OF TOTAL", fontName="Helvetica-Bold",
                 fontSize=7.5, fillColor=STEEL, textAnchor="end"))
    d.add(Line(0, height - head_h, width, height - head_h,
               strokeColor=LINE, strokeWidth=0.8))

    # ---- rows ----
    for i, (_, r) in enumerate(df.iterrows()):
        row_top = height - head_h - i * row_h
        cy = row_top - row_h / 2
        is_top = r["rank"] <= 3

        # zebra background
        if i % 2 == 1:
            d.add(Rect(0, row_top - row_h, width, row_h,
                       fillColor=ZEBRA, strokeColor=None))

        bar_color = ACCENT if is_top else BASEBAR
        name_font = "Helvetica-Bold" if is_top else "Helvetica"
        text_y = cy - 3.2

        # rank + name
        d.add(String(2, text_y, str(int(r["rank"])), fontName="Helvetica-Bold",
                     fontSize=9, fillColor=ACCENT if is_top else STEEL))
        d.add(String(22, text_y, str(r["name"]), fontName=name_font,
                     fontSize=9.5, fillColor=INK))

        # bar
        bar_h = 15
        by = cy - bar_h / 2
        w = max(bar_max * (r["cases"] / maxv), 2)
        d.add(Rect(bar_x, by, w, bar_h, fillColor=bar_color, strokeColor=None))

        # case count: inside the bar if it fits, otherwise just past the end
        cases_str = f"{int(r['cases']):,}"
        if w > 50:
            d.add(String(bar_x + w - 6, text_y, cases_str, fontName="Helvetica-Bold",
                         fontSize=8.5, fillColor=colors.white, textAnchor="end"))
        else:
            d.add(String(bar_x + w + 5, text_y, cases_str, fontName="Helvetica-Bold",
                         fontSize=8.5, fillColor=STEEL, textAnchor="start"))

        # % of total (right column)
        d.add(String(width - 2, text_y, f"{r['pct']:.2f}%", fontName=name_font,
                     fontSize=9.5, fillColor=INK if is_top else STEEL,
                     textAnchor="end"))

        # row separator
        d.add(Line(0, row_top - row_h, width, row_top - row_h,
                   strokeColor=FAINT, strokeWidth=0.5))

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
    gen = datetime.now().strftime("%B %d, %Y  %I:%M %p")
    head = Table([[Paragraph(title, h_title)],
                  [Paragraph(subtitle, h_sub)],
                  [Paragraph(f"Generated {gen}", h_sub)]], colWidths=[content_w])
    head.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), HEADER),
        ("LEFTPADDING", (0, 0), (-1, -1), 16),
        ("RIGHTPADDING", (0, 0), (-1, -1), 16),
        ("TOPPADDING", (0, 0), (0, 0), 14),
        ("BOTTOMPADDING", (-1, -1), (-1, -1), 14),
        ("LINEBELOW", (0, -1), (-1, -1), 3, ACCENT),
    ]))
    story.append(head)
    story.append(Spacer(1, 12))

    # ---- summary boxes ----
    total = int(df["cases"].sum())
    avg = int(round(df["cases"].mean()))
    stats = [("PICKERS", f"{len(df)}"),
             ("TOTAL CASES", f"{total:,}"),
             ("AVG PER PICKER", f"{avg:,}"),
             ("TOP PICKER", df.iloc[0]["name"])]
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
    story.append(Paragraph("Leaderboard", h_sec))
    story.append(labeled_bar_chart(df, content_w))

    foot = ParagraphStyle("foot", fontName="Helvetica", fontSize=7.5,
                          textColor=STEEL, alignment=1, spaceBefore=10)
    story.append(Spacer(1, 4))
    story.append(Paragraph("Ranked by cases picked. Top three shown in amber.", foot))

    doc.build(story)
    return buf.getvalue()


# =============================================================================
# On-screen leaderboard (matches the PDF: one labelled chart)
# =============================================================================
def render_leaderboard(df: pd.DataFrame, title: str, subtitle: str):
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Pickers", len(df))
    c2.metric("Total cases", f"{int(df['cases'].sum()):,}")
    c3.metric("Avg per picker", f"{int(round(df['cases'].mean())):,}")
    c4.metric("Top picker", df.iloc[0]["name"])

    st.markdown("##### Leaderboard")
    plot = df.copy()
    plot["label"] = plot.apply(
        lambda r: f"{int(r['cases']):,}  •  {r['pct']:.2f}%", axis=1)
    try:
        import altair as alt
        base = alt.Chart(plot).encode(
            y=alt.Y("name:N", sort="-x", title=None,
                    axis=alt.Axis(labelFontWeight="bold", labelFontSize=12)),
        )
        bars = base.mark_bar(cornerRadiusEnd=2).encode(
            x=alt.X("cases:Q", title="Cases picked"),
            color=alt.condition(alt.datum.rank <= 3,
                                alt.value("#b45309"), alt.value("#52606d")),
            tooltip=["rank", "name", "cases", "pct"],
        )
        text = base.mark_text(align="left", dx=4, fontSize=11,
                              fontWeight="bold", color="#1f2933").encode(
            x="cases:Q", text="label:N")
        st.altair_chart((bars + text).properties(height=max(240, len(df) * 30)),
                        use_container_width=True)
    except Exception:
        st.bar_chart(df.set_index("name")["cases"], use_container_width=True)

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
st.caption("Paste a screenshot or raw text. The app extracts each picker's name, "
           "cases picked, and share of the total, then builds a one-page report.")

with st.expander("Report details (shown on the PDF)", expanded=False):
    report_title = st.text_input("Title", value="Case-Picking Leaderboard")
    report_subtitle = st.text_input("Subtitle / shift / date", value="Warehouse Operations")

if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""

tab_text, tab_image = st.tabs(["Paste text", "Paste / upload screenshot"])

with tab_text:
    st.session_state.raw_text = st.text_area(
        "Paste rows here (one picker per line):",
        value=st.session_state.raw_text, height=240,
        placeholder="AIRELH    1,760    0    11.74\nDANIELV   1,659    0    11.07\n...",
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
                st.session_state.raw_text = st.session_state.get("ocr_box", extracted)
            except Exception as e:
                st.error(f"Couldn't read the image ({type(e).__name__}). "
                         "Use the Paste text tab instead.")

st.divider()

if st.button("Build leaderboard", type="primary", use_container_width=True):
    text = st.session_state.get("ocr_box") or st.session_state.raw_text
    df = parse_rows(text or "")
    if df.empty:
        st.error("No valid rows found. Check the text/OCR output above — each line "
                 "should be a name followed by numbers.")
    else:
        render_leaderboard(df, report_title, report_subtitle)
