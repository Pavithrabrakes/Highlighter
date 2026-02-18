import io
import re
import zipfile
import unicodedata
from collections import defaultdict, Counter
from typing import List, Tuple, Dict, Set, Optional

import streamlit as st
import pandas as pd
import fitz  # PyMuPDF

# =============================================================
# CONFIG (highlight defaults + layout heuristics)
# =============================================================
DEFAULT_HIGHLIGHT_HEX = "#FFD400"    # for IDs present in Excel
DEFAULT_NONEXCEL_HEX = "#00E0A8"     # for IDs NOT present in Excel (optional)
DEFAULT_ALPHA = 0.50

# Header ratio only used for "Left column only" mode
HEADER_Y_TOP_RATIO = 0.10            # ignore top header area (10% height)
DEFAULT_LEFT_BAND_WIDTH = 220.0
BAND_MARGIN = 8.0

# --- STRICT UTR FORMAT: Q + 7 digits with alphanumeric boundaries
ID_STRICT = re.compile(r"(?<![A-Za-z0-9])Q\d{7}(?![A-Za-z0-9])", re.IGNORECASE)

# Tolerant pattern (Excel & PDF): allow small separators between Q and digits (space, -, _, /, .)
# Example: Q 1-4/3_5.2 7 6  --> normalize to Q1435276
ID_FUZZY = re.compile(
    r"(?<![A-Za-z0-9])"          # left boundary
    r"[Qq]"                       # Q
    r"(?:[\s\-\_/\.]*\d){7}"      # 7 digits with 0+ small separators
    r"(?![A-Za-z0-9])",           # right boundary
)

# =============================================================
# STREAMLIT PAGE
# =============================================================
st.set_page_config(page_title="Reference Highlighter Web", page_icon="🖍️", layout="wide")

# --------- MODERN UI / CSS ---------
st.markdown(
    """
    <style>
    body { background: linear-gradient(135deg, #EEF3FF 0%, #FFFFFF 100%); }
    div.block-container { padding-top: 2rem; padding-bottom: 3rem; }
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
    .chip {
        display:inline-flex; align-items:center; gap:8px;
        border-radius:999px; padding:6px 12px;
        background:#EEF2FF; color:#3730A3; font-size:12.5px;
        border:1px solid rgba(55,48,163,0.18);
        margin-right:6px;
    }
    .stButton>button {
        background:#0066FF !important; color:#fff !important; border:none !important;
        border-radius:10px !important; padding:10px 22px !important; font-size:16px !important;
        box-shadow: 0 10px 24px rgba(0,102,255,0.25);
        transition: transform .16s ease, box-shadow .16s ease, background .16s ease;
    }
    .stButton>button:hover { background:#0053CC !important; transform:translateY(-1px) scale(1.02); }
    .tiny { font-size:12.5px; color:#6B7280; }
    .sep { border:0; height:1px; background:rgba(0,0,0,0.07); margin:10px 0 16px 0; }
    @keyframes fadeIn   { from {opacity:0} to {opacity:1} }
    @keyframes fadeInUp { from {opacity:0; transform:translateY(10px)} to {opacity:1; transform:translateY(0)} }
    </style>
    """,
    unsafe_allow_html=True,
)

# =============================================================
# HELPERS: normalization & color
# =============================================================
def hex_to_rgb01(hex_color: str) -> Tuple[float, float, float]:
    h = hex_color.lstrip("#")
    r, g, b = tuple(int(h[i:i+2], 16) / 255 for i in (0, 2, 4))
    return (r, g, b)

def normalize_q7(raw: str) -> Optional[str]:
    """
    Normalize a raw match (possibly with separators) to canonical 'Q########'.
    Returns None if it doesn't normalize cleanly.
    """
    if not raw:
        return None
    s = unicodedata.normalize("NFKD", raw).upper()
    s = re.sub(r"[^A-Z0-9]", "", s)   # drop separators
    if re.fullmatch(r"Q\d{7}", s):
        return s
    return None

# =============================================================
# EXCEL LOADER
# =============================================================
def normalize_header(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.replace("\xa0", " ")
    s = " ".join(s.split())
    return s.strip().lower()

def looks_like_utr_header(h: str) -> bool:
    return "utr" in normalize_header(h)

def extract_utrs_from_cell(val) -> List[str]:
    """
    Extract each valid UTR from a cell WITHOUT concatenating unrelated content.
    Accept 'Q' + 7 digits with small separators, normalize to 'Q########'.
    """
    if pd.isna(val):
        return []
    s = unicodedata.normalize("NFKD", str(val)).replace("\u00A0", " ")
    results = []
    # tolerant matches
    for m in ID_FUZZY.finditer(s):
        norm = normalize_q7(m.group(0))
        if norm:
            results.append(norm)
    # also catch already contiguous strict hits (defensive)
    for m in ID_STRICT.finditer(s):
        norm = m.group(0).upper()
        if norm not in results:
            results.append(norm)
    return results

def load_utrs_from_excel(excel_files, show_cols=False) -> Tuple[Set[str], List[str]]:
    """
    Returns:
        utrs_unique: set of unique normalized IDs
        utrs_all: list of all extracted IDs (with duplicates) for auditing
    Processes ALL sheets and ALL columns whose header contains 'utr'.
    Tries header rows 0..4 per sheet, and uses the FIRST header row that yields any UTR column for that sheet.
    """
    utrs_unique: Set[str] = set()
    utrs_all: List[str] = []
    any_found_in_file = False

    for f in excel_files:
        f.seek(0)
        try:
            xls = pd.ExcelFile(f, engine="openpyxl")
        except Exception as e:
            st.error(f"Unable to open Excel '{getattr(f, 'name', 'file')}': {e}")
            st.stop()

        file_found = False
        for sheet in xls.sheet_names:
            sheet_found = False
            for hdr in range(0, 5):
                try:
                    df = xls.parse(sheet_name=sheet, header=hdr)
                except Exception:
                    continue

                if show_cols:
                    st.caption(f"🔎 {getattr(f, 'name', 'file')} → Sheet '{sheet}', header row {hdr}")
                    st.write(list(df.columns))

                utr_cols = [c for c in df.columns if looks_like_utr_header(c)]
                if not utr_cols:
                    continue

                for col in utr_cols:
                    series = df[col].dropna()
                    for val in series:
                        ids = extract_utrs_from_cell(val)
                        for tid in ids:
                            utrs_unique.add(tid)
                            utrs_all.append(tid)

                sheet_found = True
                break  # stop trying other header rows for this sheet
            file_found = file_found or sheet_found

        if not file_found:
            st.warning(
                f"⚠️ No UTR-like column found in any sheet of '{getattr(f, 'name', 'file')}'. "
                f"Expected values like Q1234567 in a column containing 'UTR' in its header."
            )
        any_found_in_file = any_found_in_file or file_found

    if not any_found_in_file:
        st.error("No UTR values found in any uploaded Excel file.")
        st.stop()

    return utrs_unique, utrs_all

# =============================================================
# PDF TEXT PIPELINE: group words & build line strings
# =============================================================
def add_visual_highlight(page: fitz.Page, rect: fitz.Rect, doc: fitz.Document,
                         color_rgb01=(1, 1, 0), alpha=0.5):
    annot = page.add_highlight_annot(rect)
    annot.set_opacity(alpha)
    annot.set_colors(stroke=color_rgb01)
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

def group_words_by_line(words: List[Tuple]) -> Dict[Tuple[int, int], List[Tuple]]:
    """
    Group word tuples by (block_no, line_no) keeping original order (word_no).
    words element: (x0, y0, x1, y1, text, block_no, line_no, word_no)
    """
    lines: Dict[Tuple[int, int], List[Tuple]] = defaultdict(list)
    for (x0, y0, x1, y1, wtxt, block_no, line_no, word_no) in words:
        lines[(block_no, line_no)].append((x0, y0, x1, y1, wtxt, word_no))
    for key in lines:
        lines[key].sort(key=lambda t: (t[5], t[0]))
    return lines

def line_string_and_spans(line_words: List[Tuple]) -> Tuple[str, List[Tuple[int, int, fitz.Rect]]]:
    """
    Build an uppercase line string where non-alnum characters become spaces,
    and a single space is inserted between words to preserve boundaries.
    """
    line_norm_parts: List[str] = []
    spans: List[Tuple[int, int, fitz.Rect]] = []
    cursor = 0

    for (x0, y0, x1, y1, wtxt, _word_no) in line_words:
        w = unicodedata.normalize("NFKD", str(wtxt)).replace("\u00A0", " ")
        w_norm = re.sub(r"[^A-Za-z0-9]", " ", w).upper()
        w_norm = re.sub(r"\s+", " ", w_norm).strip()
        if not w_norm:
            continue

        start = cursor
        end = start + len(w_norm)
        line_norm_parts.append(w_norm)
        spans.append((start, end, fitz.Rect(x0, y0, x1, y1)))
        cursor = end + 1  # +1 for inter-word space

    return " ".join(line_norm_parts), spans

def spans_to_rect(spans: List[Tuple[int, int, fitz.Rect]], start: int, end: int) -> List[fitz.Rect]:
    rects: List[fitz.Rect] = []
    for s, e, r in spans:
        if e <= start or s >= end:
            continue
        rects.append(r)
    return rects

# =============================================================
# HIGHLIGHTERS + ID EXTRACTION FROM PDF
# =============================================================
def robust_left_band(words: List[Tuple], page_left: float, page_right: float) -> Tuple[float, float, str]:
    """Estimate a left column band using the 10th percentile of x0 of token-like words."""
    xs = []
    for (x0, y0, x1, y1, wtxt, *_rest) in words:
        w = unicodedata.normalize("NFKD", str(wtxt)).replace("\u00A0", " ")
        norm = re.sub(r"[^A-Za-z0-9]", "", w)
        if len(norm) >= 3:
            xs.append(x0)
    if xs:
        xs_sorted = sorted(xs)
        idx = max(0, int(0.10 * len(xs_sorted)) - 1)
        x0_band = xs_sorted[idx]
        x1_band = min(page_right, x0_band + DEFAULT_LEFT_BAND_WIDTH)
        return x0_band, x1_band, "auto-quantile"
    return page_left, min(page_left + DEFAULT_LEFT_BAND_WIDTH, page_right), "fallback"

def extract_and_highlight_pdf(
    pdf_bytes: bytes,
    excel_ids: Set[str],
    scope: str,
    manual_x1: Optional[float],
    color_in: str,
    color_out: str,
    alpha: float,
    tolerant_pdf: bool,
    highlight_non_excel: bool,
) -> Tuple[io.BytesIO, Dict[str, bool], List[Dict], List[str]]:
    """
    Parse the PDF, extract ALL IDs (normalized) with their pages, and highlight:
      - IDs present in Excel -> color_in
      - IDs NOT in Excel -> color_out (optional)
    Returns:
      - out_pdf: BytesIO of highlighted PDF
      - found_map: {excel_id -> True/False} if seen in PDF
      - pdf_id_rows: list of dict rows: {"PDF": name (placeholder), "Page": pno, "ID": id, "InExcel": bool}
                    NOTE: caller will fill 'PDF' with actual file name.
      - logs: processing logs for debug panel
    """
    pattern = ID_FUZZY if tolerant_pdf else ID_STRICT
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    found_map = {u: False for u in excel_ids}
    logs = []
    pdf_id_rows: List[Dict] = []

    color_in_rgb = hex_to_rgb01(color_in)
    color_out_rgb = hex_to_rgb01(color_out)

    for pno, page in enumerate(doc, start=1):
        try:
            words = page.get_text("words", sort=True) or []
        except Exception:
            words = []

        # scope filtering
        if scope == "Left column only":
            page_left = float(page.rect.x0)
            page_right = float(page.rect.x1)
            page_height = float(page.rect.height)
            y_min = page_height * HEADER_Y_TOP_RATIO

            filtered = []
            for (x0, y0, x1, y1, wtxt, block_no, line_no, word_no) in words:
                if y0 < y_min:
                    continue
                if not str(wtxt).strip():
                    continue
                filtered.append((x0, y0, x1, y1, wtxt, block_no, line_no, word_no))

            if manual_x1 is not None:
                x0_band = page_left
                x1_band = float(manual_x1)
            else:
                x0_band, x1_band, _ = robust_left_band(filtered, page_left, page_right)

            in_band = []
            for (x0, y0, x1, y1, wtxt, bno, lno, wno) in filtered:
                if (x0 >= (x0_band - BAND_MARGIN)) and (x1 <= (x1_band + BAND_MARGIN)):
                    in_band.append((x0, y0, x1, y1, wtxt, bno, lno, wno))
            lines = group_words_by_line(in_band)
        else:
            lines = group_words_by_line(words)

        page_hits = 0
        # Scan each line with pattern, map to rectangles, normalize, record & highlight
        for _key, line_words in lines.items():
            line_norm, spans = line_string_and_spans(line_words)
            if not line_norm:
                continue

            for m in pattern.finditer(line_norm):
                raw = m.group(0)
                norm_id = normalize_q7(raw)
                if not norm_id:
                    continue

                rects = spans_to_rect(spans, m.start(), m.end())
                in_excel = norm_id in excel_ids

                # collect rows for the Excel report
                pdf_id_rows.append({"PDF": "", "Page": pno, "ID": norm_id, "InExcel": in_excel})

                # highlight
                if in_excel:
                    for r in rects:
                        add_visual_highlight(page, r, doc, color_in_rgb, alpha)
                    if not found_map[norm_id]:
                        found_map[norm_id] = True
                    page_hits += 1
                else:
                    if highlight_non_excel:
                        for r in rects:
                            add_visual_highlight(page, r, doc, color_out_rgb, alpha)
                        page_hits += 1

        logs.append(f"Page {pno}: words={len(words)}, hits={page_hits}")

    out = io.BytesIO()
    doc.save(out, deflate=True, garbage=4)
    doc.close()
    out.seek(0)
    return out, found_map, pdf_id_rows, logs

# =============================================================
# HERO
# =============================================================
st.markdown(
    """
    <div class="hero-box">
        <h1 class="hero-title"> Reference Highlighter Web</h1>
        <p class="hero-sub">
            Upload Excel + PDF → we will extract all IDs from PDFs, compare with Excel,
            and produce highlighted PDFs + an Excel report (present vs not present).
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
    scope = st.radio(
        "Highlight scope",
        ["Whole page (anywhere)", "Left column only"],
        index=0,
        help="Search entire page or restrict to a left-column band."
    )
    manual_band = st.toggle("Manual left-column boundary (x1)", value=False,
                            help="Only used in 'Left column only' mode.")
    manual_x1 = st.slider("Right boundary (x1, points)", 120, 600, 260, 5,
                          disabled=(not manual_band or scope != "Left column only"))
    st.markdown("<hr class='sep'>", unsafe_allow_html=True)
    color_hex_in = st.color_picker("Highlight color (IDs present in Excel)", DEFAULT_HIGHLIGHT_HEX)
    color_hex_out = st.color_picker("Highlight color (IDs NOT in Excel)", DEFAULT_NONEXCEL_HEX)
    highlight_non_excel = st.toggle("Also highlight IDs NOT present in Excel", value=False)
    opacity_pct = st.slider("Opacity", 20, 90, int(DEFAULT_ALPHA * 100), 5)
    tolerant_pdf = st.toggle(
        "Tolerant PDF matching (allow spaces/hyphens inside IDs)", value=True,
        help="Enable to catch IDs printed with small separators like 'Q143-6784' or 'Q143 6784'."
    )
    st.markdown("<span class='tiny'>Tip: 'Whole page' guarantees every match is highlighted.</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown("<div class='ui-card'>", unsafe_allow_html=True)
    st.subheader("⬇️ Result")
    st.caption("We’ll prepare a ZIP with highlighted PDFs + Excel report.")
    result_slot = st.empty()
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

# =============================================================
# CTA
# =============================================================
cta_a, cta_b, cta_c = st.columns([0.2, 0.6, 0.2])
with cta_b:
    start = st.button("🚀 Start", use_container_width=True)

# =============================================================
# MAIN ACTION
# =============================================================
if start:
    if not excel_files or not pdf_files:
        st.error("Please upload at least one Excel and one PDF.")
        st.stop()

    # 1) Load UTRs from Excel
    with st.spinner("Loading Reference IDs from Excel…"):
        excel_ids_unique, excel_ids_all = load_utrs_from_excel(excel_files, show_cols=False)

    unique_count = len(excel_ids_unique)
    total_count = len(excel_ids_all)
    dup_counter = Counter(excel_ids_all)
    duplicates_only = {k: v for k, v in dup_counter.items() if v > 1}

    st.info(f"Extracted **{unique_count} unique** IDs from Excel "
            f"(**{total_count} total** occurrences across sheets/columns).")

    if unique_count == 0:
        st.warning("No Reference IDs found in the uploaded Excel file(s). "
                   "Expected values like 'Q1234567' in a column containing 'UTR'.")
        st.stop()

    with st.expander("🔎 Excel duplicates summary"):
        if duplicates_only:
            dup_df = pd.DataFrame(
                [{"ID": k, "count": v} for k, v in sorted(duplicates_only.items(), key=lambda x: (-x[1], x[0]))]
            )
            st.dataframe(dup_df, use_container_width=True, hide_index=True)
        else:
            st.write("No duplicates found.")

    # 2) Process PDFs: extract ALL IDs, compare with Excel, and highlight accordingly
    color_in = color_hex_in or DEFAULT_HIGHLIGHT_HEX
    color_out = color_hex_out or DEFAULT_NONEXCEL_HEX
    alpha = max(0.2, min(0.9, opacity_pct / 100.0))

    progress = st.progress(0, text="Preparing…")
    total_pdfs = len(pdf_files)
    zip_buffer = io.BytesIO()

    # Collect rows for final Excel report
    pdf_id_rows_all: List[Dict] = []   # per occurrence from PDFs
    excel_status_rows: List[Dict] = [] # one row per Excel ID + found flag

    all_logs = []

    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
        # Iterate PDFs
        for idx, pdf_file in enumerate(pdf_files, start=1):
            progress.progress(idx / total_pdfs, text=f"Processing {pdf_file.name} ({idx}/{total_pdfs})…")
            pdf_file.seek(0)
            pdf_bytes = pdf_file.read()

            out_pdf, found_map, pdf_id_rows, logs = extract_and_highlight_pdf(
                pdf_bytes=pdf_bytes,
                excel_ids=excel_ids_unique,
                scope=scope,
                manual_x1=(manual_x1 if manual_band else None),
                color_in=color_in,
                color_out=color_out,
                alpha=alpha,
                tolerant_pdf=tolerant_pdf,
                highlight_non_excel=highlight_non_excel,
            )

            # Attach file name to pdf rows
            for row in pdf_id_rows:
                row["PDF"] = pdf_file.name
            pdf_id_rows_all.extend(pdf_id_rows)

            all_logs.extend([f"[{pdf_file.name}] {ln}" for ln in logs])

            # Write highlighted PDF to ZIP
            out_name = (
                pdf_file.name[:-4] + "_highlighted.pdf"
                if pdf_file.name.lower().endswith(".pdf")
                else pdf_file.name + "_highlighted.pdf"
            )
            zipf.writestr(out_name, out_pdf.getvalue())

            # Update found flags (we want one row per Excel ID, aggregate later)
            for u in sorted(excel_ids_unique):
                excel_status_rows.append({"ID": u, "FoundInPDF": bool(found_map.get(u, False))})

        # Build the Excel report workbook in memory
        # Sheet 1: PDF_IDs (every occurrence in PDFs)
        pdf_df = pd.DataFrame(pdf_id_rows_all, columns=["PDF", "Page", "ID", "InExcel"]).sort_values(["PDF", "Page", "ID"])

        # Sheet 2: Excel_IDs (unique IDs from Excel with found flag)
        excel_df = pd.DataFrame(excel_status_rows).drop_duplicates().sort_values(["ID"])

        # Sheet 3: Summary
        pdf_ids_unique = set(pdf_df["ID"].unique()) if not pdf_df.empty else set()
        in_both = sorted(list(pdf_ids_unique.intersection(excel_ids_unique)))
        pdf_only = sorted(list(pdf_ids_unique.difference(excel_ids_unique)))
        excel_only = sorted(list(excel_ids_unique.difference(pdf_ids_unique)))

        summary_rows = [
            {"Metric": "Excel Unique IDs", "Value": len(excel_ids_unique)},
            {"Metric": "PDF Unique IDs", "Value": len(pdf_ids_unique)},
            {"Metric": "Present in BOTH", "Value": len(in_both)},
            {"Metric": "Present only in PDF (not in Excel)", "Value": len(pdf_only)},
            {"Metric": "Present only in Excel (not in PDF)", "Value": len(excel_only)},
        ]
        summary_df = pd.DataFrame(summary_rows)

        # Write an .xlsx with 3 sheets
        xlsx_buffer = io.BytesIO()
        with pd.ExcelWriter(xlsx_buffer, engine="openpyxl") as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Summary")
            # Also write the actual ID lists in summary for reference
            pd.DataFrame({"Both": in_both}).to_excel(writer, index=False, sheet_name="IDs_Both")
            pd.DataFrame({"PDF_only": pdf_only}).to_excel(writer, index=False, sheet_name="IDs_PDF_only")
            pd.DataFrame({"Excel_only": excel_only}).to_excel(writer, index=False, sheet_name="IDs_Excel_only")
            pdf_df.to_excel(writer, index=False, sheet_name="PDF_IDs")
            excel_df.to_excel(writer, index=False, sheet_name="Excel_IDs")
        zipf.writestr("ID_Report.xlsx", xlsx_buffer.getvalue())

    progress.empty()
    st.success("✅ Done! Your ZIP is ready.")

    with right:
        with result_slot.container():
            st.download_button(
                "⬇️ Download ZIP",
                data=zip_buffer.getvalue(),
                file_name="Reference_output.zip",
                mime="application/zip",
                use_container_width=True,
            )

    with st.expander("🔧 Debug log (first 200 lines)"):
        for ln in all_logs[:200]:
            st.text(ln)

st.caption(
    "Notes: We extract all PDF IDs (Q+7 digits) and compare with Excel. "
    "By default we highlight only IDs that exist in Excel; enable the toggle to also highlight IDs not in Excel."
)
