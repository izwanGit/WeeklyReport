import streamlit as st
import io
import traceback
import sys
import os
import base64
import datetime
import re

try:
    import fitz  # PyMuPDF
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ── Page Config ──
st.set_page_config(
    page_title="Monthly Report | PETRONAS",
    page_icon=os.path.join(BASE_DIR, "PETRONAS_LOGO_SQUARE.png"),
    layout="wide",
)

# ── Branding Helpers ──
def _image_to_data_uri(path, mime_type):
    try:
        with open(os.path.join(BASE_DIR, path), 'rb') as f:
            data = f.read()
        return f"data:{mime_type};base64,{base64.b64encode(data).decode()}"
    except:
        return ""

_logo_square_uri = _image_to_data_uri("PETRONAS_LOGO_SQUARE.png", "image/png")
_logo_sidebar_uri = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg", "image/svg+xml")


# ── Premium Corporate CSS ──
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif !important;
    }

    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    [data-testid="stSidebarNav"],
    .stDeployButton {
        display: none !important;
        visibility: hidden !important;
    }
    header[data-testid="stHeader"] {
        background: transparent !important;
        border-bottom: none !important;
    }

    [data-testid="stAppViewContainer"] > .main { transition: none !important; }

    [data-testid="stSidebar"] { animation: none !important; }
    [data-testid="stSidebarNav"],
    [data-testid="stSidebarNavItems"],
    [data-testid="stSidebarNavSeparator"],
    [data-testid="stStatusWidget"] {
        display: none !important;
        visibility: hidden !important;
    }
    header[data-testid="stHeader"] { background: #F8FAFC !important; }
    [data-testid="stSidebar"] { border-right: none !important; }

    .main .block-container {
        padding-top: 1rem !important;
        max-width: 1200px !important;
    }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #00B1A9 !important; font-weight: 700 !important;
    }
    .stButton > button, .stDownloadButton > button {
        background: #00B1A9 !important; color: white !important;
        border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; transition: all 0.3s ease !important;
        padding: 0.6rem 1.4rem !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #008C86 !important; transform: translateY(-1px) !important; color: white !important;
    }
    [data-testid="stFileUploader"] {
        border: 2px dashed rgba(0, 177, 169, 0.35) !important;
        border-radius: 12px !important; padding: 16px 20px !important;
    }
    [data-testid="stMetric"] {
        background: #FFFFFF !important; border: 1px solid #E2E8F0 !important;
        border-left: 4px solid #00B1A9 !important; border-radius: 12px !important;
        padding: 1rem 1.2rem !important; box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
    }
    [data-testid="stMetricValue"] { color: #00B1A9 !important; font-weight: 800 !important; }
    [data-testid="stMetricLabel"] { color: #4A5568 !important; font-weight: 500 !important; }

    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
        margin-bottom: 0px !important;
    }
    [data-testid="stSidebar"] .stTextInput { margin-bottom: -10px !important; }

    .stDeployButton, [data-testid="stDeployButton"], [data-testid="stAppDeployButton"] { display: none !important; }
    #MainMenu { visibility: hidden; } footer { visibility: hidden; }

    .section-label {
        font-size: 0.7rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #94A3B8;
        margin-bottom: 12px;
        margin-top: 8px;
    }

    [data-testid="stSidebarNav"] { display: none !important; }

    .sidebar-nav {
        display: block;
        padding: 10px 16px;
        margin: 3px 12px;
        border-radius: 8px;
        text-decoration: none !important;
        color: #334155 !important;
        font-family: 'Inter', sans-serif;
        font-weight: 600;
        font-size: 0.88rem;
        transition: all 0.2s ease;
    }
    .sidebar-nav:hover {
        background: #F1F5F9;
        color: #1E293B !important;
        text-decoration: none !important;
    }
    .sidebar-nav.active {
        background: linear-gradient(135deg, #00B1A9, #008C86);
        color: white !important;
        font-weight: 700;
        box-shadow: 0 4px 12px rgba(0,177,169,0.25);
    }
    .sidebar-sep { border: none; border-top: 1px solid #E2E8F0; margin: 16px 12px; }
    .genie-link {
        font-size: 0.85rem; font-weight: 500;
        color: #31333F !important; text-decoration: none !important;
        transition: all 0.2s ease !important; cursor: pointer !important;
    }
    .genie-link:hover { color: #00B1A9 !important; text-decoration: none !important; }
</style>
""", unsafe_allow_html=True)


# ── Sidebar ──
with st.sidebar:
    st.markdown(f"""
<div style="text-align:center; padding:8px 0 20px 0;">
    <a href="/" target="_self" style="display:inline-block;">
        <img src="{_logo_sidebar_uri}" style="height:56px; transition: transform 0.2s; cursor: pointer;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'"/>
    </a>
</div>
""", unsafe_allow_html=True)

    st.markdown("### Report Settings")
    _months = ["January", "February", "March", "April", "May", "June",
               "July", "August", "September", "October", "November", "December"]
    _now = datetime.date.today()
    sel_month = st.selectbox("Report Month", options=_months, index=_now.month - 1)
    _current_year = _now.year
    _years = [str(y) for y in range(_current_year - 1, _current_year + 3)]
    sel_year = st.selectbox("Report Year", options=_years, index=1)

    st.markdown("<div style='margin-top: -10px;'></div>", unsafe_allow_html=True)
    st.markdown("### Data Upload")
    st.markdown("<a href='#' target='_blank' class='genie-link'>Power BI PDF Export ↗</a>", unsafe_allow_html=True)
    pdf_file = st.file_uploader("Power BI PDF Export", type=['pdf'], label_visibility="collapsed")

    TEMPLATE_PATH = os.path.join(BASE_DIR, "template.pptx")
    pptx_available_locally = os.path.exists(TEMPLATE_PATH)
    uploaded_template = None
    if not pptx_available_locally:
        st.warning("template.pptx not found in app folder.")
        uploaded_template = st.file_uploader("Upload PPTX Template", type=['pptx'])
    else:
        st.success("template.pptx detected.")

st.markdown("""
<a href="/" target="_self" style="text-decoration: none; display: inline-flex; align-items: center; gap: 8px; font-weight: 600; color: #64748B; margin-bottom: 16px; transition: color 0.2s ease;">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><line x1="19" y1="12" x2="5" y2="12"></line><polyline points="12 19 5 12 12 5"></polyline></svg>
    Back to Hub
</a>
""", unsafe_allow_html=True)

# ── Header Banner ──
st.markdown(f"""
<style>
.banner-title {{ color: #FFFFFF !important; text-transform: uppercase !important; font-weight: 800 !important; text-shadow: 0px 2px 4px rgba(0,0,0,0.3) !important; margin: 0 !important; line-height: 1.1 !important; white-space: nowrap; font-size: clamp(1.2rem, 3.5vw, 1.8rem) !important; letter-spacing: 0.1px; }}
.banner-subtitle {{ color: #FFFFFF !important; font-weight: 400 !important; text-shadow: 0px 1px 3px rgba(0,0,0,0.2) !important; margin: 4px 0 0 0 !important; white-space: nowrap; font-size: clamp(0.85rem, 2vw, 1.0rem) !important; opacity: 0.95 !important; }}
</style>
<div style="display: flex; align-items: center; gap: 24px; padding: 22px 32px; background-color: #00B1A9; border-radius: 20px; margin-bottom: 2rem; box-shadow: 0 12px 35px rgba(0, 177, 169, 0.25); overflow: hidden; border: 1px solid rgba(255, 255, 255, 0.15);">
<img src="{_logo_square_uri}" style="height: 80px; flex-shrink: 0; filter: drop-shadow(1px 1px 0 white) drop-shadow(-1px -1px 0 white) drop-shadow(1px -1px 0 white) drop-shadow(-1px 1px 0 white);" />
<div style="min-width: 0;">
<h1 class="banner-title">Monthly PPTX Automation</h1>
<p class="banner-subtitle">Power BI dashboard export to corporate PowerPoint deck — zero-touch pipeline.</p>
</div>
</div>
""", unsafe_allow_html=True)


# ── Dependency Check ──
if not PPTX_AVAILABLE:
    st.error("Required libraries (python-pptx, PyMuPDF) are not installed. Please run: pip install -r requirements.txt")
    st.stop()


# ══════════════════════════════════════════════════════════════
# HELPER FUNCTIONS
# ══════════════════════════════════════════════════════════════

def _emu_to_inches(emu):
    """PowerPoint uses English Metric Units: 1 inch = 914400 EMU."""
    return emu / 914400.0


def _auto_crop_bottom(png_bytes, margin=20, threshold=245):
    """
    Trim blank whitespace from the bottom of a PNG image.
    Used for PDF pages 7 & 14 (module lists) where the list length
    varies month-to-month, leaving ugly whitespace below.
    """
    from PIL import Image
    img = Image.open(io.BytesIO(png_bytes)).convert("RGB")
    pixels = img.load()
    w, h = img.size
    last_content_row = h - 1
    for y in range(h - 1, -1, -1):
        row_is_blank = True
        for x in range(0, w, 5):
            r, g, b = pixels[x, y]
            if r < threshold or g < threshold or b < threshold:
                row_is_blank = False
                break
        if not row_is_blank:
            last_content_row = y
            break
    crop_bottom = min(last_content_row + margin, h)
    if crop_bottom < h - 30:
        img = img.crop((0, 0, w, crop_bottom))
    out = io.BytesIO()
    img.save(out, format="PNG")
    return out.getvalue()


def _extract_month_values(page_text):
    """
    Extract (month_abbr, numeric_value) pairs from fitz text extraction.
    Returns a list sorted chronologically.
    Returns [] if extraction fails.
    """
    MONTHS_ORDER = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    lines = [l.strip() for l in page_text.split('\n') if l.strip()]
    found = {}

    # Pass 1: "Mon YYYY" line, value within next 5 lines
    for i, line in enumerate(lines):
        m = re.match(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$', line)
        if m:
            mon = m.group(1)
            for j in range(i + 1, min(i + 6, len(lines))):
                nm = re.match(r'^(\d+(?:\.\d+)?)\s*$', lines[j])
                if nm:
                    found[mon] = float(nm.group(1))
                    break

    # Pass 2: value within 5 lines before "Mon YYYY"
    if not found:
        for i, line in enumerate(lines):
            m = re.match(r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}$', line)
            if m:
                mon = m.group(1)
                for j in range(i - 1, max(i - 6, -1), -1):
                    nm = re.match(r'^(\d+(?:\.\d+)?)\s*$', lines[j])
                    if nm:
                        if mon not in found:
                            found[mon] = float(nm.group(1))
                        break

    return [(mon, found[mon]) for mon in MONTHS_ORDER if mon in found]


def _replace_in_para(para, pattern, replacement, flags=re.IGNORECASE):
    """
    Replace text matching pattern in a paragraph while preserving formatting.
    Puts all text in the first run, clears remaining runs.
    Returns True if replacement was made.
    """
    if not para.runs:
        return False
    full_text = "".join(run.text for run in para.runs)
    if not re.search(pattern, full_text, flags):
        return False
    new_text = re.sub(pattern, replacement, full_text, flags=flags)
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""
    return True


def update_dates_in_pptx(prs, new_month, new_year):
    """
    Replace date references across ALL slides in the PPTX:
    - "March 2026" → "{new_month} {new_year}"
    - Standalone year (e.g. "2025") → "{new_year}"
    - Standalone month name → "{new_month}"
    Returns a list of log messages.
    """
    MONTH_NAMES = [
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    ]
    MONTH_ABBREVS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                     'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    all_months = sorted(MONTH_NAMES + MONTH_ABBREVS, key=len, reverse=True)
    month_pat = '|'.join(re.escape(m) for m in all_months)

    changes = []
    for slide_num, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue
            for para in shape.text_frame.paragraphs:
                original = "".join(run.text for run in para.runs)

                # Pattern: "Month Year" (e.g., "March 2026")
                if _replace_in_para(
                    para,
                    rf'({month_pat})\s+(20\d{{2}})',
                    f'{new_month} {new_year}'
                ):
                    changes.append(f"  DATE | Slide {slide_num + 1}: month+year updated (was: '{original.strip()}')")
                    continue

                # Pattern: standalone month name (e.g., "March")
                if _replace_in_para(
                    para,
                    rf'\b({month_pat})\b',
                    new_month
                ):
                    changes.append(f"  DATE | Slide {slide_num + 1}: month name updated")
                    continue

                # Pattern: standalone 4-digit year (e.g., "2025 Dashboard")
                if _replace_in_para(
                    para,
                    r'\b20\d{2}\b',
                    str(new_year)
                ):
                    changes.append(f"  DATE | Slide {slide_num + 1}: year updated")

    return changes


def _compute_change(values):
    """
    Given a list of (month_abbr, value) tuples (chronological),
    return (prev_val, curr_val, abs_diff, direction_str) comparing last two entries.
    Returns None if fewer than 2 data points.
    """
    if len(values) < 2:
        return None
    curr = int(round(values[-1][1]))
    prev = int(round(values[-2][1]))
    diff = abs(curr - prev)
    direction = "increased" if curr > prev else "decreased" if curr < prev else "unchanged"
    return prev, curr, diff, direction


def update_summary_bullets(prs, sel_month, sel_year,
                            sr_trend_vals, sr_ageing_vals,
                            inc_trend_vals, inc_ageing_vals):
    """
    Update auto-generated summary bullets on:
      - Slide 4 (idx 3): SR ticket comparison + SR ageing comparison
      - Slide 8 (idx 7): INC ticket comparison + INC ageing comparison

    Searches text frames for paragraphs containing 'ticket'/'ageing' keywords
    and replaces them with freshly computed comparisons.
    """
    changes = []

    def make_ticket_bullet(vals, label):
        result = _compute_change(vals)
        if not result:
            return None
        prev, curr, diff, direction = result
        if direction == "unchanged":
            return f"{label} tickets remained unchanged at {curr} in {sel_month} {sel_year}"
        return (f"{label} tickets logged {direction} by {diff} in {sel_month} {sel_year} "
                f"({prev} → {curr})")

    def make_ageing_bullet(vals, label):
        result = _compute_change(vals)
        if not result:
            return None
        prev, curr, diff, direction = result
        if direction == "unchanged":
            return f"{label} ageing remained unchanged at {curr} in {sel_month} {sel_year}"
        return (f"{label} ageing {direction} from {prev} to {curr} in {sel_month} {sel_year} "
                f"({'+' if direction == 'increased' else '-'}{diff})")

    SUMMARY_CONFIG = [
        # (slide_idx_0based, ticket_vals, ageing_vals, label)
        (3, sr_trend_vals,  sr_ageing_vals,  "SR"),
        (7, inc_trend_vals, inc_ageing_vals, "INC"),
    ]

    for slide_idx, ticket_vals, ageing_vals, label in SUMMARY_CONFIG:
        if slide_idx >= len(prs.slides):
            continue

        slide = prs.slides[slide_idx]
        ticket_bullet = make_ticket_bullet(ticket_vals, label)
        ageing_bullet = make_ageing_bullet(ageing_vals, label)

        ticket_updated = False
        ageing_updated = False

        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue
            for para in shape.text_frame.paragraphs:
                if not para.runs:
                    continue
                para_text_lower = "".join(run.text for run in para.runs).lower()

                # Match ticket bullet: contains "ticket" AND a number
                if (not ticket_updated and ticket_bullet and
                        'ticket' in para_text_lower and
                        re.search(r'\d', para_text_lower)):
                    para.runs[0].text = ticket_bullet
                    for run in para.runs[1:]:
                        run.text = ""
                    changes.append(f"  SUMM | Slide {slide_idx + 1}: Ticket bullet → '{ticket_bullet}'")
                    ticket_updated = True
                    continue

                # Match ageing bullet: contains "ageing" or "aging" AND a number
                elif (not ageing_updated and ageing_bullet and
                        ('ageing' in para_text_lower or 'aging' in para_text_lower) and
                        re.search(r'\d', para_text_lower)):
                    para.runs[0].text = ageing_bullet
                    for run in para.runs[1:]:
                        run.text = ""
                    changes.append(f"  SUMM | Slide {slide_idx + 1}: Ageing bullet → '{ageing_bullet}'")
                    ageing_updated = True

        if not ticket_updated:
            if ticket_vals:
                changes.append(f"  SUMM | Slide {slide_idx + 1} WARN: Could not find {label} ticket bullet to update.")
            else:
                changes.append(f"  SUMM | Slide {slide_idx + 1} WARN: No {label} trend data from PDF — ticket bullet skipped.")
        if not ageing_updated:
            if ageing_vals:
                changes.append(f"  SUMM | Slide {slide_idx + 1} WARN: Could not find {label} ageing bullet to update.")
            else:
                changes.append(f"  SUMM | Slide {slide_idx + 1} WARN: No {label} ageing data from PDF — ageing bullet skipped.")

    return changes


# ══════════════════════════════════════════════════════════════
# PROCESSING ENGINE — FULL 13-IMAGE PIPELINE + SMART TEXT
# Maps ALL PDF pages 2–14 to PPT slides 3–10.
# Updates dates + summary bullets automatically.
# ══════════════════════════════════════════════════════════════
def process_monthly_report(pdf_bytes, pptx_bytes, sel_month, sel_year):
    """
    Full automation engine for HCM Dashboard PPTX replacement.

    Steps:
    1. Extract all 14 PDF pages as high-res PNGs
    2. Extract metric data from PDF pages 3, 4, 10, 11 (for smart summary)
    3. Auto-crop bottom whitespace on PDF pages 7 & 14 (module lists)
    4. Replace 13 picture shapes across PPT slides 3–10
    5. Update all dates across the entire presentation
    6. Update summary bullet text on slides 4 & 8 with computed comparisons

    Returns: (pptx_bytes, log_list, replaced_count, pdf_images_dict)
    """

    # ── Phase 1: Extract all PDF pages as high-res PNGs ──
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pdf_images = {}   # 1-indexed: page_number → png_bytes
    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))  # ~300 DPI
        pdf_images[page_num + 1] = pix.tobytes("png")

    log = []
    log.append(f"INFO | PDF loaded: {len(pdf)} pages extracted at 300 DPI")

    # ── Phase 2: Extract metric data for smart summary ──
    sr_trend_vals, sr_ageing_vals, inc_trend_vals, inc_ageing_vals = [], [], [], []

    DATA_PAGE_MAP = {
        3:  ("sr_trend",   "SR Ticket Trend"),
        4:  ("sr_ageing",  "SR Ageing"),
        10: ("inc_trend",  "INC Ticket Trend"),
        11: ("inc_ageing", "INC Ageing"),
    }
    raw_metrics = {}
    for pg_num, (key, label) in DATA_PAGE_MAP.items():
        if pg_num <= len(pdf):
            pg = pdf.load_page(pg_num - 1)
            raw_metrics[key] = pg.get_text()

    sr_trend_vals  = _extract_month_values(raw_metrics.get("sr_trend",  ""))
    sr_ageing_vals = _extract_month_values(raw_metrics.get("sr_ageing", ""))
    inc_trend_vals = _extract_month_values(raw_metrics.get("inc_trend", ""))
    inc_ageing_vals= _extract_month_values(raw_metrics.get("inc_ageing",""))

    log.append(f"INFO | SR Trend data:   {sr_trend_vals  or 'none extracted (chart may be image-only)'}")
    log.append(f"INFO | SR Ageing data:  {sr_ageing_vals or 'none extracted (chart may be image-only)'}")
    log.append(f"INFO | INC Trend data:  {inc_trend_vals or  'none extracted (chart may be image-only)'}")
    log.append(f"INFO | INC Ageing data: {inc_ageing_vals or 'none extracted (chart may be image-only)'}")

    # ── Phase 3: Auto-crop bottom whitespace on module list pages ──
    for pg in (7, 14):
        if pg in pdf_images:
            old_size = len(pdf_images[pg])
            pdf_images[pg] = _auto_crop_bottom(pdf_images[pg])
            new_size = len(pdf_images[pg])
            if new_size < old_size:
                log.append(f"INFO | PDF page {pg}: Auto-cropped ({old_size // 1024}KB → {new_size // 1024}KB)")

    # ── Phase 4: Open PPTX template ──
    prs = Presentation(io.BytesIO(pptx_bytes))
    log.append(f"INFO | PPTX loaded: {len(prs.slides)} slides in template")

    # ── Phase 5: Image replacement ──
    # Slide (0-based idx) → [pdf pages (1-based)] → [labels]
    SLIDE_CONFIG = [
        (2, [2,  3],      ["SR SLA Performance",      "SR Ticket Trend"]),
        (3, [4],          ["SR Ageing Trend"]),
        (4, [5],          ["SR Root Cause Table"]),
        (5, [6,  7],      ["SR Category Distribution", "SR Module Ticket List"]),
        (6, [8,  9,  10], ["INC Response SLA",         "INC Resolution SLA", "INC Ticket Trend"]),
        (7, [11],         ["INC Ageing Trend"]),
        (8, [12],         ["INC Root Cause Table"]),
        (9, [13, 14],     ["INC Category Distribution","INC Module Ticket List"]),
    ]

    LOGO_TOP_MAX    = 0.5   # inches (logo sits at top=0.081)
    LOGO_HEIGHT_MAX = 0.6   # inches (logo height is 0.402)
    replaced = 0

    for slide_idx, pdf_pages, labels in SLIDE_CONFIG:
        if slide_idx >= len(prs.slides):
            log.append(f"WARN | Slide {slide_idx + 1} does not exist. Skipping.")
            continue

        slide = prs.slides[slide_idx]

        # Collect non-logo PICTURE shapes, sorted top→bottom, left→right
        candidates = []
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            s_top    = _emu_to_inches(shape.top)
            s_height = _emu_to_inches(shape.height)
            if s_top < LOGO_TOP_MAX and s_height < LOGO_HEIGHT_MAX:
                continue  # skip template logo
            candidates.append(shape)

        candidates.sort(key=lambda s: (s.top, s.left))

        if len(candidates) < len(pdf_pages):
            log.append(
                f"WARN | Slide {slide_idx + 1}: Expected {len(pdf_pages)} picture(s) "
                f"but found {len(candidates)}. Partial replacement may occur."
            )

        for k, (pdf_page, label) in enumerate(zip(pdf_pages, labels)):
            if k >= len(candidates):
                log.append(f"WARN | Slide {slide_idx + 1}: No picture slot #{k+1} for '{label}'.")
                break

            if pdf_page not in pdf_images:
                log.append(f"WARN | PDF page {pdf_page} not available. Skipping '{label}'.")
                continue

            old_shape = candidates[k]
            img_io = io.BytesIO(pdf_images[pdf_page])

            slide.shapes.add_picture(
                img_io,
                old_shape.left, old_shape.top,
                old_shape.width, old_shape.height
            )
            old_shape._element.getparent().remove(old_shape._element)

            bbox = (
                _emu_to_inches(old_shape.left),  _emu_to_inches(old_shape.top),
                _emu_to_inches(old_shape.width), _emu_to_inches(old_shape.height),
            )
            log.append(
                f"  OK | Slide {slide_idx + 1}: '{label}' <- PDF p{pdf_page} "
                f"(L={bbox[0]:.2f} T={bbox[1]:.2f} W={bbox[2]:.2f} H={bbox[3]:.2f})"
            )
            replaced += 1

    # ── Phase 6: Smart date replacement ──
    log.append("INFO | Running date replacement across all slides...")
    date_changes = update_dates_in_pptx(prs, sel_month, str(sel_year))
    if date_changes:
        log.extend(date_changes)
        log.append(f"INFO | Date replacement: {len(date_changes)} text updates made.")
    else:
        log.append("INFO | Date replacement: no date text found to update.")

    # ── Phase 7: Smart summary bullet update ──
    log.append("INFO | Running summary bullet update on slides 4 & 8...")
    summary_changes = update_summary_bullets(
        prs, sel_month, sel_year,
        sr_trend_vals, sr_ageing_vals,
        inc_trend_vals, inc_ageing_vals
    )
    log.extend(summary_changes)

    # ── Phase 8: Save ──
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    log.append(f"INFO | Complete: {replaced}/13 images replaced successfully.")
    # Return pdf_images too so the preview section can show them
    return output.read(), log, replaced, pdf_images


# ── Resolve which PPTX template to use ──
resolved_pptx = None
if uploaded_template is not None:
    resolved_pptx = "uploaded"
elif pptx_available_locally:
    resolved_pptx = "local"


# ══════════════════════════════════════════════════════════════
# GENERATE SECTION
# ══════════════════════════════════════════════════════════════
if pdf_file and resolved_pptx:
    st.markdown('<p class="section-label">Step 2 &mdash; Validation &amp; Generation</p>', unsafe_allow_html=True)

    pdf_file.seek(0)
    try:
        preview_pdf = fitz.open(stream=pdf_file.read(), filetype="pdf")
        pdf_page_count = len(preview_pdf)
        pdf_file.seek(0)
    except:
        pdf_page_count = "?"

    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("PDF Pages Detected", pdf_page_count)
    with m2:
        st.metric("Target Slides", "8 (Slides 3–10)")
    with m3:
        st.metric("Image Swaps", "13")

    with st.expander("PDF → PPT Mapping (what goes where)", expanded=False):
        st.markdown("""
| PDF Page | Content | → PPT Slide | Target | Auto-feature |
|:---:|---|:---:|---|---|
| ~~1~~ | *Landing/Filters* | *Skip* | *Not used* | — |
| **2** | SR SLA Performance | **Slide 3** | Left image | — |
| **3** | SR Ticket Trend | **Slide 3** | Right image | Data extracted for summary |
| **4** | SR Ageing Trend | **Slide 4** | Histogram image | Data extracted for summary |
| **5** | SR Root Cause Table | **Slide 5** | Table image | — |
| **6** | SR Category % | **Slide 6** | Left image | — |
| **7** | SR Module List | **Slide 6** | Right image | ✂️ Auto-cropped |
| **8** | INC Response SLA | **Slide 7** | Top-left image | — |
| **9** | INC Resolution SLA | **Slide 7** | Bottom-left image | — |
| **10** | INC Ticket Trend | **Slide 7** | Right image | Data extracted for summary |
| **11** | INC Ageing Trend | **Slide 8** | Histogram image | Data extracted for summary |
| **12** | INC Root Cause Table | **Slide 9** | Table image | — |
| **13** | INC Category % | **Slide 10** | Left image | — |
| **14** | INC Module List | **Slide 10** | Right image | ✂️ Auto-cropped |
        """)
        st.info(
            "**Auto-features applied automatically:**\n"
            "- ✂️ Module list images (pages 7 & 14) are cropped to remove blank whitespace\n"
            "- 📅 All dates updated throughout the entire PPTX\n"
            "- 📊 Summary bullets on Slides 4 & 8 auto-computed from trend data\n"
            "- 🔒 Template logo is always protected and never replaced",
            icon="✨"
        )

    if pdf_page_count != "?" and pdf_page_count < 14:
        st.warning(
            f"⚠️ Expected 14 pages in the PDF, but found {pdf_page_count}. "
            "Some dashboard images may not be replaced."
        )

    st.markdown("")

    if st.button("Generate Monthly Report", use_container_width=True, type="primary"):
        with st.spinner("Extracting visuals, updating text, and building the presentation..."):
            try:
                if uploaded_template is not None:
                    template_bytes = uploaded_template.read()
                else:
                    with open(TEMPLATE_PATH, "rb") as f:
                        template_bytes = f.read()

                out_bytes, build_logs, img_count, pdf_images_result = process_monthly_report(
                    pdf_file.read(),
                    template_bytes,
                    sel_month,
                    int(sel_year)
                )

                if img_count == 13:
                    st.success(f"✅ Perfect — all {img_count}/13 images replaced + dates & summary updated.")
                elif img_count > 0:
                    st.warning(f"⚠️ Partial — {img_count}/13 images replaced. Check the build log.")
                else:
                    st.error("❌ No images were replaced. Check your PDF and PPTX template.")

                with st.expander("Build Log", expanded=(img_count < 13)):
                    for msg in build_logs:
                        if "WARN" in msg:
                            st.warning(msg)
                        elif msg.startswith("  OK") or msg.startswith("  DATE") or msg.startswith("  SUMM"):
                            st.text(msg)
                        else:
                            st.info(msg)

                # ── PPTX Preview ──
                if img_count > 0 and pdf_images_result:
                    st.markdown("")
                    st.markdown('<p class="section-label">Preview — What Was Placed Per Slide</p>', unsafe_allow_html=True)

                    PREVIEW_CONFIG = [
                        ("Slide 3 — SR SLA + Ticket Trend",        [2, 3]),
                        ("Slide 4 — SR Ageing Trend",               [4]),
                        ("Slide 5 — SR Root Cause Table",           [5]),
                        ("Slide 6 — SR Category % + Module List",   [6, 7]),
                        ("Slide 7 — INC Response + Resolution + Trend", [8, 9, 10]),
                        ("Slide 8 — INC Ageing Trend",              [11]),
                        ("Slide 9 — INC Root Cause Table",          [12]),
                        ("Slide 10 — INC Category % + Module List", [13, 14]),
                    ]

                    for slide_label, pdf_page_nums in PREVIEW_CONFIG:
                        available = [p for p in pdf_page_nums if p in pdf_images_result]
                        if not available:
                            continue
                        with st.expander(f"🖼️ {slide_label}", expanded=False):
                            cols = st.columns(len(available))
                            for col, pg_num in zip(cols, available):
                                with col:
                                    st.caption(f"PDF Page {pg_num}")
                                    st.image(pdf_images_result[pg_num], use_container_width=True)

                # ── Download ──
                if img_count > 0:
                    st.markdown("")
                    st.markdown('<p class="section-label">Step 3 &mdash; Download</p>', unsafe_allow_html=True)

                    st.info(
                        "📝 **Text note:** All text in the PPTX is untouched EXCEPT:\n"
                        "- Dates (month/year) across all slides — **auto-updated**\n"
                        "- Summary bullet points on Slides 4 & 8 — **auto-updated** (if data was extractable from the PDF)\n\n"
                        "Please review Slides 4 & 8 summary text if the build log shows 'none extracted'.",
                        icon="ℹ️"
                    )

                    st.download_button(
                        label="Download Final Report (.pptx)",
                        data=out_bytes,
                        file_name=f"Monthly_Report_{sel_month}_{sel_year}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"An error occurred during generation: {e}")
                st.code(traceback.format_exc(), language="text")

elif not resolved_pptx and pdf_file:
    st.warning("⚠️ No PowerPoint template available. Please upload a PPTX template in the sidebar.")

else:
    st.markdown("""
    <div style="
        text-align: center; padding: 60px 40px;
        background: linear-gradient(180deg, #FFFFFF 0%, #F0FAFA 100%);
        border: 1px solid #E2E8F0; border-top: 4px solid #00B1A9;
        border-radius: 16px; margin-top: 8px;
    ">
        <h2 style="color: #1A202C; font-weight: 800; margin: 0 0 10px 0;">Upload Required</h2>
        <p style="color: #718096; max-width: 520px; margin: 0 auto; line-height: 1.7; font-size: 0.95rem;">
            Please select the target Report Month and Year, and upload your Power BI PDF export using the sidebar.
            The corporate PowerPoint template will be automatically loaded from the system.
        </p>
    </div>
    """, unsafe_allow_html=True)
