import io
import re
import zipfile
import streamlit as st
import pandas as pd
import fitz  # PyMuPDF

# =============================================================
# CONFIG (highlight defaults + layout heuristics)
# =============================================================
DEFAULT_HIGHLIGHT_HEX = "#FFD400"    # yellow
DEFAULT_ALPHA = 0.50                 # 50% opacity
HEADER_Y_TOP_RATIO = 0.18            # ignore top header area (18% height)
DEFAULT_LEFT_BAND_WIDTH = 140.0      # auto left-column width
BAND_WIDEN_PT = 24.0                 # optional widen (not used in fast path)
UTR_PATTERN = re.compile(r"[A-Za-z0-9]{6,}")  # tokens like Q1435783

# =============================================================
# STREAMLIT PAGE (theme handled via .streamlit/config.toml if present)
# =============================================================
st.set_page_config(
    page_title="Reference Highlighter Web",
    page_icon="🖍️",
    layout="wide",
)

# --------- MODERN UI / CSS ---------
st.markdown(
    """
    <style>
    /* container + background */
    body { background: linear-gradient(135deg, #EEF3FF 0%, #FFFFFF 100%); }
    div.block-container { padding-top: 2rem; padding-bottom: 3rem; }

    /* hero */
    .hero-box {
        background: rgba(255,255,255,0.70);
        padding: 24px 22px;
        border-radius: 16px;
        border: 1px solid rgba(0,0,0,0.07);
        backdrop-filter: blur(10px);
        -webkit-backdrop-filter: blur(10px);
        box-shadow: 0 12px 30px rgba(0,0,0,0.06);
        animation: fadeIn 700ms ease;
    }
    .hero-title { margin: 0 0 6px 0; font-weight: 800; }
    .hero-sub  { margin: 0; color: #475569; }

    /* cards */
    .ui-card {
        background: #FFFFFF;
        padding: 18px 16px;
        border-radius: 14px;
        border: 1px solid rgba(0,0,0,0.07);
        box-shadow: 0 8px 22px rgba(0,0,0,0.08);
        transition: transform .18s ease, box-shadow .18s ease;
        animation: fadeInUp 600ms ease;
    }
    .ui-card:hover { transform: translateY(-4px); box-shadow: 0 14px 30px rgba(0,0,0,0.12); }

    /* chips */
    .chip {
        display:inline-flex; align-items:center; gap:8px;
        border-radius:999px; padding:6px 12px;
        background:#EEF2FF; color:#3730A3; font-size:12.5px;
        border:1px solid rgba(55,48,163,0.18);
        margin-right:6px;
    }

    /* primary button */
    .stButton>button {
        background:#0066FF !important; color:#fff !important; border:none !important;
        border-radius:10px !important; padding:10px 22px !important; font-size:16px !important;
        box-shadow: 0 10px 24px rgba(0,102,255,0.25);
        transition: transform .16s ease, box-shadow .16s ease, background .16s ease;
    }
    .stButton>button:hover { background:#0053CC !important; transform:translateY(-1px) scale(1.02); }

    .tiny { font-size:12.5px; color:#6B7280; }
    .muted { color:#64748B; }
    .sep { border:0; height:1px; background:rgba(0,0,0,0.07); margin:10px 0 16px 0; }

    @keyframes fadeIn   { from {opacity:0} to {opacity:1} }
    @keyframes fadeInUp { from {opacity:0; transform:translateY(10px)} to {opacity:1; transform:translateY(0)} }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================
# EXCEL HELPERS
# =============================================================
def normalize_header(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip().lower()

def looks_like_utr_header(h: str) -> bool:
    return "utr" in normalize_header(h)

def extract_utrs_from_cell(val):
    if pd.isna(val):
        return []
    return UTR_PATTERN.findall(str(val).strip())

def load_utrs_from_excel(excel_files, show_cols=False) -> set[str]:
    """
    Auto-detect the first column containing 'utr' by scanning sheets and header rows 0..4.
    Returns UPPERCASE tokens set.
    """
    utrs = set()

    for f in excel_files:
        f.seek(0)
        try:
            xls = pd.ExcelFile(f, engine="openpyxl")
        except Exception:
            st.error(f"Unable to open Excel '{f.name}'.")
            st.stop()

        found = False
        for sheet in xls.sheet_names:
            for hdr in range(0, 5):
                f.seek(0)
                try:
                    df = pd.read_excel(f, sheet_name=sheet, header=hdr, engine="openpyxl")
                except Exception:
                    continue

                if show_cols:
                    st.caption(f"🔎 {f.name} → Sheet '{sheet}', header {hdr}")
                    st.write(list(df.columns))

                utr_col = None
                for c in df.columns:
                    if looks_like_utr_header(c):
                        utr_col = c
                        break

                if utr_col is None:
                    continue

                for val in df[utr_col].dropna():
                    for token in extract_utrs_from_cell(val):
                        utrs.add(token.strip().upper())

                found = True
                break
            if found:
                break

        if not found:
            st.error(f"NO UTR-like column found in '{f.name}'.")
            st.stop()

    return utrs

# =============================================================
# PDF HELPERS
# =============================================================
def hex_to_rgb01(hex_color: str) -> tuple[float, float, float]:
    h = hex_color.lstrip("#")
    r, g, b = tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))
    return (r, g, b)

def add_visual_highlight(page: fitz.Page, rect: fitz.Rect, doc: fitz.Document,
                         color_rgb01=(1, 1, 0), alpha=0.5):
    annot = page.add_highlight_annot(rect)
    annot.set_opacity(alpha)
    annot.set_colors(stroke=color_rgb01)  # (r,g,b) in 0..1
    annot.update()
    # strip comment fields to keep it visual-only
    try:
        xref = annot.xref
        doc.xref_set_key(xref, "T", "null")
        doc.xref_set_key(xref, "Contents", "null")
        doc.xref_set_key(xref, "Popup", "null")
    except Exception:
        try:
            annot.set_info({"title": "", "content": ""})
            annot.update()
        except Exception:
            pass

# =============================================================
# FAST LEFT-COLUMN HIGHLIGHTER (single pass per page)
# =============================================================
def highlight_left_column_fast(pdf_bytes: bytes,
                               utrs_upper: set[str],
                               manual_x1: float | None = None,
                               color_hex: str = DEFAULT_HIGHLIGHT_HEX,
                               alpha: float = DEFAULT_ALPHA):
    """
    Fastest method:
      - Extract words once
      - Auto-detect left-column band from first match
      - Highlight only words in band + below header
    Returns: (BytesIO PDF, found_map, logs)
    """
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    found_map = {u: False for u in utrs_upper}
    logs = []
    color_rgb = hex_to_rgb01(color_hex)

    for pno, page in enumerate(doc, start=1):
        try:
            words = page.get_text("words") or []
        except Exception:
            words = []

        page_left = float(page.rect.x0)
        page_right = float(page.rect.x1)
        page_height = float(page.rect.height)
        y_min = page_height * HEADER_Y_TOP_RATIO  # don't touch header row

        # find first on-page match to learn left-band start (x0)
        sample_left_x = None
        for (x0, y0, x1, y1, wtxt, *_rest) in words:
            if y0 < y_min:
                continue
            wu = str(wtxt).strip().upper()
            if wu in utrs_upper:
                sample_left_x = x0
                break

        # compute band
        if manual_x1 is not None:
            x0_band = page_left
            x1_band = float(manual_x1)
            band_mode = "manual"
        else:
            if sample_left_x is not None:
                x0_band = sample_left_x
                x1_band = min(page_right, x0_band + DEFAULT_LEFT_BAND_WIDTH)
                band_mode = "auto"
            else:
                x0_band = page_left
                x1_band = min(page_left + DEFAULT_LEFT_BAND_WIDTH, page_right)
                band_mode = "fallback"

        kept = 0
        total_words = len(words)

        for (x0, y0, x1, y1, wtxt, *_rest) in words:
            if y0 < y_min:
                continue
            if not (x0 >= x0_band and x1 <= x1_band):
                continue
            wu = str(wtxt).strip().upper()
            if wu in utrs_upper:
                kept += 1
                found_map[wu] = True
                add_visual_highlight(page, fitz.Rect(x0, y0, x1, y1), doc, color_rgb, alpha)

        logs.append(
            f"Page {pno}: words={total_words}, hits={kept}, band=({x0_band:.1f}-{x1_band:.1f}), mode={band_mode}"
        )

    out = io.BytesIO()
    doc.save(out, deflate=True, garbage=4)
    doc.close()
    out.seek(0)
    return out, found_map, logs

# =============================================================
# HERO
# =============================================================
st.markdown(
    """
    <div class="hero-box">
        <h1 class="hero-title"> Reference Highlighter Web</h1>
        <p class="hero-sub">
            Upload Excel + PDF → we highlight only the <strong>left‑column</strong> IDs that exist in Excel — fast, accurate, and non‑destructive.
        </p>
    </div>
    """,
    unsafe_allow_html=True,
)
st.markdown("<br>", unsafe_allow_html=True)

# =============================================================
# LAYOUT (Upload • Settings • Result)
# =============================================================
left, mid, right = st.columns([0.45, 0.25, 0.30], gap="large")

with left:
    st.markdown("<div class='ui-card'>", unsafe_allow_html=True)
    st.subheader("📤 Upload")
    st.caption("Excel with Reference IDs + Bank PDF(s).")
    excel_files = st.file_uploader("Excel file(s)", type=["xlsx"], accept_multiple_files=True, label_visibility="collapsed")
    pdf_files = st.file_uploader("PDF file(s)", type=["pdf"], accept_multiple_files=True, label_visibility="collapsed")
    if excel_files:
        st.markdown(f"<span class='chip'>Excel: {len(excel_files)}</span>", unsafe_allow_html=True)
    if pdf_files:
        st.markdown(f"<span class='chip'>PDFs: {len(pdf_files)}</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with mid:
    st.markdown("<div class='ui-card'>", unsafe_allow_html=True)
    st.subheader("⚙️ Settings")
    manual_band = st.toggle("Manual left-column boundary (x1)", value=False, help="If OFF, we auto-detect per page.")
    manual_x1 = st.slider("Right boundary (x1, points)", 120, 480, 220, 5, disabled=not manual_band)
    st.markdown("<hr class='sep'>", unsafe_allow_html=True)
    color_hex = st.color_picker("Highlight color", DEFAULT_HIGHLIGHT_HEX)
    opacity_pct = st.slider("Opacity", 20, 90, int(DEFAULT_ALPHA * 100), 5)
    st.markdown("<span class='tiny'>Tip: defaults work for most bank PDFs.</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown("<div class='ui-card'>", unsafe_allow_html=True)
    st.subheader("⬇️ Result")
    st.caption("We’ll prepare a ZIP with highlighted PDFs + CSV report.")
    result_slot = st.empty()
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =============================================================
# CTA
# =============================================================
cta_a, cta_b, cta_c = st.columns([0.2, 0.6, 0.2])
with cta_b:
    start = st.button("🚀 Start Highlighting", use_container_width=True)

# =============================================================
# ACTION
# =============================================================
if start:
    if not excel_files or not pdf_files:
        st.error("Please upload at least one Excel and one PDF.")
        st.stop()

    # 1) Load UTRs
    with st.spinner("Loading Reference IDs from Excel…"):
        utr_set_upper = load_utrs_from_excel(excel_files, show_cols=False)

    if not utr_set_upper:
        st.warning("No Reference IDs found in the uploaded Excel file(s).")
        st.stop()

    # 2) Process PDFs
    color_hex = color_hex or DEFAULT_HIGHLIGHT_HEX
    alpha = max(0.2, min(0.9, opacity_pct / 100.0))

    progress = st.progress(0, text="Preparing…")
    total = len(pdf_files)
    zip_buffer = io.BytesIO()
    csv_rows = []
    all_logs = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        for idx, pdf_file in enumerate(pdf_files, start=1):
            progress.progress(idx / total, text=f"Processing {pdf_file.name} ({idx}/{total})…")
            pdf_bytes = pdf_file.read()

            out_pdf, found_map, logs = highlight_left_column_fast(
                pdf_bytes,
                utr_set_upper,
                manual_x1 if manual_band else None,
                color_hex=color_hex,
                alpha=alpha,
            )

            all_logs.extend([f"[{pdf_file.name}] {ln}" for ln in logs])

            out_name = pdf_file.name[:-4] + "_highlighted.pdf" if pdf_file.name.lower().endswith(".pdf") \
                       else pdf_file.name + "_highlighted.pdf"
            zipf.writestr(out_name, out_pdf.getvalue())

            for u in sorted(utr_set_upper):
                csv_rows.append({"Reference": u, "found": bool(found_map.get(u, False))})

        # CSV report
        df = pd.DataFrame(csv_rows).drop_duplicates()
        zipf.writestr("Reference_report.csv", df.to_csv(index=False))

    progress.empty()
    st.success("✅ Done! Your ZIP is ready.")

    # 3) Download in Result card
    with right:
        with result_slot.container():
            st.download_button(
                "⬇️ Download ZIP",
                data=zip_buffer.getvalue(),
                file_name="Reference_output.zip",
                mime="application/zip",
                use_container_width=True,
            )

    # 4) Optional: Debug
    with st.expander("🔧 Debug log (first 150 lines)"):
        for ln in all_logs[:150]:
            st.text(ln)

# =============================================================
# FOOTNOTE
# =============================================================
st.markdown("<hr class='sep'>", unsafe_allow_html=True)
st.caption("Tip: For best results, keep the Excel column header containing the word “UTR”. "
           "Only left‑column matches in the PDF are highlighted; layout and logos are preserved.")