"""
Case-Picking Leaderboard
=========================
Paste a screenshot OR paste raw text. The app extracts three things per row:
    1. name
    2. cases picked
    3. % of total
…and renders a ranked leaderboard with the top 3 highlighted.

Deploy on Streamlit Community Cloud:
  - app.py            (this file)
  - requirements.txt  (python deps)
  - packages.txt      (system deps -> installs the tesseract OCR engine)
"""

import io
import re

import pandas as pd
import streamlit as st

# ---- Optional dependencies (app still runs if they're missing) --------------
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

# A clipboard "Paste from clipboard" button (nice for screenshots).
# Falls back to the normal uploader if the component isn't installed.
try:
    from streamlit_paste_button import paste_image_button
    PASTE_OK = True
except Exception:
    PASTE_OK = False


# =============================================================================
# Parsing
# =============================================================================
NUMERIC_RE = re.compile(r"^[\$]?[\d.,]+%?$")  # a "pure number" token


def _to_number(tok: str):
    """'1,760' -> 1760.0, '.41' -> 0.41, '11.74%' -> 11.74. None if not a number."""
    cleaned = tok.replace(",", "").replace("$", "").replace("%", "").strip()
    if cleaned in ("", ".", "-"):
        return None
    try:
        return float(cleaned)
    except ValueError:
        return None


def parse_rows(text: str) -> pd.DataFrame:
    """
    Turn raw text into a tidy DataFrame: name, cases, pct.

    Strategy per line:
      - tokens before the first PURELY-numeric token = the name
        (so digit-containing names like 'TKTEMP3' stay intact)
      - of the numeric tokens, the FIRST is cases and the LAST is the %
        (the throwaway middle '0' column is simply ignored)
    """
    rows = []
    for raw in text.splitlines():
        line = raw.strip()
        if not line:
            continue

        tokens = line.split()
        name_parts, numbers = [], []
        seen_number = False

        for tok in tokens:
            is_num = bool(NUMERIC_RE.match(tok)) and _to_number(tok) is not None
            if is_num:
                seen_number = True
                numbers.append(_to_number(tok))
            elif not seen_number:
                name_parts.append(tok)
            else:
                # stray text after numbers started — ignore it
                continue

        name = " ".join(name_parts).strip()
        # need a name and at least one number to be a real row
        if not name or not numbers:
            continue
        if not any(c.isalpha() for c in name):  # skip header/junk lines
            continue

        cases = numbers[0]
        pct = numbers[-1] if len(numbers) > 1 else None
        rows.append({"name": name, "cases": cases, "pct": pct})

    df = pd.DataFrame(rows)
    if df.empty:
        return df

    df["cases"] = df["cases"].fillna(0).astype(int)
    # If a % column wasn't present, compute share of the total instead.
    if df["pct"].isna().all():
        total = df["cases"].sum()
        df["pct"] = (df["cases"] / total * 100) if total else 0.0

    df = df.sort_values("cases", ascending=False).reset_index(drop=True)
    df.insert(0, "rank", df.index + 1)
    return df


def ocr_image(file_bytes: bytes) -> str:
    """Run OCR on an image, with light preprocessing to help accuracy."""
    img = Image.open(io.BytesIO(file_bytes))
    img = ImageOps.exif_transpose(img).convert("L")  # grayscale
    # upscale small screenshots so the engine has more pixels to work with
    if max(img.size) < 1600:
        scale = 1600 / max(img.size)
        img = img.resize((int(img.width * scale), int(img.height * scale)))
    img = ImageOps.autocontrast(img)
    # psm 6 = assume a uniform block of text (good for tables)
    return pytesseract.image_to_string(img, config="--psm 6")


# =============================================================================
# Leaderboard rendering
# =============================================================================
MEDALS = {1: "🥇", 2: "🥈", 3: "🥉"}


def render_leaderboard(df: pd.DataFrame):
    total_cases = int(df["cases"].sum())
    c1, c2, c3 = st.columns(3)
    c1.metric("Players", len(df))
    c2.metric("Total cases", f"{total_cases:,}")
    c3.metric("Top picker", df.iloc[0]["name"])

    st.subheader("🏆 Podium")
    top3 = df.head(3)
    cols = st.columns(len(top3))
    for col, (_, r) in zip(cols, top3.iterrows()):
        col.markdown(
            f"""
            <div style="text-align:center; padding:14px; border-radius:14px;
                        background:linear-gradient(160deg,#1f2937,#111827);
                        border:1px solid #374151;">
              <div style="font-size:42px;">{MEDALS.get(int(r['rank']),'')}</div>
              <div style="font-size:20px; font-weight:700; color:#f9fafb;">{r['name']}</div>
              <div style="font-size:26px; font-weight:800; color:#fbbf24;">{int(r['cases']):,}</div>
              <div style="color:#9ca3af;">{r['pct']:.2f}% of total</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    st.subheader("📋 Full standings")
    show = df.copy()
    show["medal"] = show["rank"].map(lambda x: MEDALS.get(x, ""))
    show["cases"] = show["cases"].map(lambda x: f"{x:,}")
    show["pct"] = show["pct"].map(lambda x: f"{x:.2f}%")
    show = show[["rank", "medal", "name", "cases", "pct"]]
    show.columns = ["#", "", "Name", "Cases picked", "% of total"]

    def highlight_top(row):
        if row["#"] in (1, 2, 3):
            return ["background-color: rgba(251,191,36,0.14)"] * len(row)
        return [""] * len(row)

    st.dataframe(
        show.style.apply(highlight_top, axis=1),
        use_container_width=True,
        hide_index=True,
    )

    st.subheader("📊 Cases picked")
    st.bar_chart(df.set_index("name")["cases"], use_container_width=True)

    st.download_button(
        "⬇️ Download as CSV",
        df.to_csv(index=False).encode("utf-8"),
        file_name="leaderboard.csv",
        mime="text/csv",
    )


# =============================================================================
# UI
# =============================================================================
st.set_page_config(page_title="Case-Picking Leaderboard", page_icon="🏆", layout="centered")
st.title("🏆 Case-Picking Leaderboard")
st.caption("Paste a screenshot or paste raw text. I'll find each name, how many cases "
           "they picked, and their % of the total — then rank everyone.")

if "raw_text" not in st.session_state:
    st.session_state.raw_text = ""

tab_text, tab_image = st.tabs(["📝 Paste text", "🖼️ Paste / upload screenshot"])

with tab_text:
    st.session_state.raw_text = st.text_area(
        "Paste the rows here (one player per line):",
        value=st.session_state.raw_text,
        height=240,
        placeholder="AIRELH    1,760    0    11.74\nDANIELV   1,659    0    11.07\n...",
    )

with tab_image:
    if not (PIL_OK and TESS_OK):
        st.warning(
            "Image OCR needs Pillow + pytesseract, and the tesseract engine via "
            "`packages.txt`. Until those are installed you can still use the "
            "**Paste text** tab."
        )
    else:
        img_bytes = None

        if PASTE_OK:
            paste_result = paste_image_button(
                "📋 Paste image from clipboard", errors="ignore"
            )
            if paste_result is not None and getattr(paste_result, "image_data", None) is not None:
                buf = io.BytesIO()
                paste_result.image_data.save(buf, format="PNG")
                img_bytes = buf.getvalue()
        else:
            st.info("Tip: add `streamlit-paste-button` to requirements.txt to paste "
                    "directly from the clipboard. For now, upload the file below.")

        uploaded = st.file_uploader(
            "…or upload / drag a screenshot", type=["png", "jpg", "jpeg", "webp"]
        )
        if uploaded is not None:
            img_bytes = uploaded.read()

        if img_bytes:
            st.image(img_bytes, caption="Source image", use_container_width=True)
            with st.spinner("Reading text from image…"):
                extracted = ocr_image(img_bytes)
            st.text_area("OCR result (edit if needed, then build below):",
                         value=extracted, height=200, key="ocr_box")
            st.session_state.raw_text = st.session_state.get("ocr_box", extracted)

st.divider()

if st.button("🚀 Build leaderboard", type="primary", use_container_width=True):
    text = st.session_state.get("ocr_box") or st.session_state.raw_text
    df = parse_rows(text or "")
    if df.empty:
        st.error("Couldn't find any valid rows. Check the text/OCR output above — "
                 "each line should be: name, then numbers.")
    else:
        render_leaderboard(df)
