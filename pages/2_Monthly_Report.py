# coding: utf-8
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

# -- Page Config --
st.set_page_config(
    page_title="Monthly Report | PETRONAS",
    page_icon=os.path.join(BASE_DIR, "PETRONAS_LOGO_SQUARE.png"),
    layout="wide",
)

# -- Branding Helpers --
def _image_to_data_uri(path, mime_type):
    try:
        with open(os.path.join(BASE_DIR, path), 'rb') as f:
            data = f.read()
        return f"data:{mime_type};base64,{base64.b64encode(data).decode()}"
    except:
        return ""

_logo_square_uri = _image_to_data_uri("PETRONAS_LOGO_SQUARE.png", "image/png")
_logo_sidebar_uri = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg", "image/svg+xml")


# -- Premium Corporate CSS --
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


# -- Sidebar --
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
    st.markdown("<a href='#' target='_blank' class='genie-link'>Power BI PDF Export</a>", unsafe_allow_html=True)
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

# -- Header Banner --
st.markdown(f"""
<style>
.banner-title {{ color: #FFFFFF !important; text-transform: uppercase !important; font-weight: 800 !important; text-shadow: 0px 2px 4px rgba(0,0,0,0.3) !important; margin: 0 !important; line-height: 1.1 !important; white-space: nowrap; font-size: clamp(1.2rem, 3.5vw, 1.8rem) !important; letter-spacing: 0.1px; }}
.banner-subtitle {{ color: #FFFFFF !important; font-weight: 400 !important; text-shadow: 0px 1px 3px rgba(0,0,0,0.2) !important; margin: 4px 0 0 0 !important; white-space: nowrap; font-size: clamp(0.85rem, 2vw, 1.0rem) !important; opacity: 0.95 !important; }}
</style>
<div style="display: flex; align-items: center; gap: 24px; padding: 22px 32px; background-color: #00B1A9; border-radius: 20px; margin-bottom: 2rem; box-shadow: 0 12px 35px rgba(0, 177, 169, 0.25); overflow: hidden; border: 1px solid rgba(255, 255, 255, 0.15);">
<img src="{_logo_square_uri}" style="height: 80px; flex-shrink: 0; filter: drop-shadow(1px 1px 0 white) drop-shadow(-1px -1px 0 white) drop-shadow(1px -1px 0 white) drop-shadow(-1px 1px 0 white);" />
<div style="min-width: 0;">
<h1 class="banner-title">Monthly PPTX Automation</h1>
<p class="banner-subtitle">Power BI dashboard export to corporate PowerPoint deck - zero-touch pipeline.</p>
</div>
</div>
""", unsafe_allow_html=True)

# -- Dependency Check --
if not PPTX_AVAILABLE:
    st.error("Required libraries (python-pptx, PyMuPDF) are not installed. Please run: pip install -r requirements.txt")
    st.stop()


# ==============================================================
# HELPER FUNCTIONS
# ==============================================================

def _emu_to_inches(emu):
    """PowerPoint uses English Metric Units: 1 inch = 914400 EMU."""
    return emu / 914400.0


def _auto_crop_bottom(png_bytes, margin=20, threshold=245):
    """Trim blank whitespace from the bottom of a PNG image."""
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


def _get_image_size(png_bytes):
    """Return (width_px, height_px) of a PNG image."""
    from PIL import Image
    with Image.open(io.BytesIO(png_bytes)) as img:
        return img.size


def _calc_fit_rect(img_w, img_h, box_left, box_top, box_w, box_h):
    """
    Fit image inside bounding box preserving aspect ratio, centred.
    Used for CHART and TABLE replacements (slides 4, 5, 8, 9).
    All box_* values are in EMU; img_w/img_h are pixels (ratio only matters).
    Returns (left, top, width, height) all in EMU.
    """
    scale = min(box_w / img_w, box_h / img_h)
    new_w = int(img_w * scale)
    new_h = int(img_h * scale)
    offset_x = (box_w - new_w) // 2
    offset_y = (box_h - new_h) // 2
    return box_left + offset_x, box_top + offset_y, new_w, new_h


def _calc_list_rect(img_w, img_h, box_left, box_top, box_w):
    """
    Width-constrained placement for variable-length module lists (slides 6, 10).
    Locks to the bounding box width and lets height follow the actual
    (already auto-cropped) image content.
    Used so that a month with 5 modules is short and one with 20 is tall,
    with no stretching either way.
    Returns (left, top, width, height) all in EMU.
    """
    scale = box_w / img_w
    new_h = int(img_h * scale)
    return box_left, box_top, box_w, new_h


def _extract_month_values_spatial(page):
    """
    Extract (month_abbr, numeric_value) pairs using spatial text analysis.
    Uses fitz page object (not raw text) to get text blocks with X/Y positions.
    Groups data labels that are vertically aligned with month axis labels.
    Returns list sorted chronologically, or [] on failure.
    """
    MONTHS_ORDER = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    MON_PAT = re.compile(
        r'^(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\b',
        re.IGNORECASE
    )

    try:
        blocks = page.get_text("dict")["blocks"]
    except:
        return []

    # Collect all text spans with their positions
    spans = []
    for block in blocks:
        if "lines" not in block:
            continue
        for line in block["lines"]:
            for span in line["spans"]:
                text = span["text"].strip()
                if not text:
                    continue
                bbox = span["bbox"]  # (x0, y0, x1, y1)
                cx = (bbox[0] + bbox[2]) / 2  # center x
                cy = (bbox[1] + bbox[3]) / 2  # center y
                spans.append((text, cx, cy, bbox))

    # Find month labels (typically on the X-axis, near the bottom)
    month_spans = []
    for text, cx, cy, bbox in spans:
        m = MON_PAT.match(text)
        if m:
            month_spans.append((m.group(1)[:3].capitalize(), cx, cy))

    if not month_spans:
        return []

    # Month labels are usually at X-axis (highest Y values among month spans)
    # Sort by Y descending to find the axis row
    month_spans_sorted = sorted(month_spans, key=lambda x: x[2], reverse=True)
    # Use the Y level of the first month as reference (they should be on same row)
    axis_y = month_spans_sorted[0][2]
    # Keep only months near the axis Y (within 20pt tolerance)
    axis_months = [(m, cx, cy) for m, cx, cy in month_spans if abs(cy - axis_y) < 20]

    if not axis_months:
        return []

    # Find numeric values - look for numbers ABOVE each month's X position
    # Data labels in bar/line charts sit above the data point, which is
    # vertically above the axis label
    found = {}
    for mon, mcx, mcy in axis_months:
        if mon in found:
            continue
        # Look for numbers vertically above this month label (same X zone, lower Y)
        best_val = None
        best_dist = float('inf')
        for text, cx, cy, bbox in spans:
            # Must be above the month label (smaller Y) and in same X zone
            if cy >= mcy:
                continue
            x_dist = abs(cx - mcx)
            if x_dist > 40:  # too far horizontally
                continue
            # Try to parse as number
            clean = text.replace(',', '').replace('%', '').strip()
            nm = re.match(r'^-?(\d+(?:\.\d+)?)$', clean)
            if not nm:
                continue
            val = float(nm.group(1))
            # Prefer the value closest to the axis (highest Y that's still above)
            dist = mcy - cy
            if dist < best_dist:
                best_dist = dist
                best_val = val

        if best_val is not None:
            found[mon] = best_val

    return [(mon, found[mon]) for mon in MONTHS_ORDER if mon in found]


def _extract_month_values_text(page_text):
    """
    Fallback: Extract (month_abbr, numeric_value) pairs from raw text lines.
    Handles Power BI date formats: "Jan 2026", "Jan 26'", "Jan '26", "Jan 26".
    """
    MONTHS_ORDER = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                    'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']
    MON_PAT = r'(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)'
    DATE_PAT = re.compile(
        rf"^{MON_PAT}\s+(?:'?\d{{2}}'?|\d{{4}})$", re.IGNORECASE
    )

    lines = [l.strip() for l in page_text.split('\n') if l.strip()]
    found = {}

    # Pass 1: value AFTER month label
    for i, line in enumerate(lines):
        m = DATE_PAT.match(line)
        if m:
            mon = m.group(1)[:3].capitalize()
            for j in range(i + 1, min(i + 6, len(lines))):
                nm = re.match(r'^(\d+(?:,\d+)*(?:\.\d+)?)\s*$', lines[j])
                if nm:
                    val = float(nm.group(1).replace(',', ''))
                    if mon not in found:
                        found[mon] = val
                    break

    # Pass 2: value BEFORE month label
    if not found:
        for i, line in enumerate(lines):
            m = DATE_PAT.match(line)
            if m:
                mon = m.group(1)[:3].capitalize()
                for j in range(i - 1, max(i - 6, -1), -1):
                    nm = re.match(r'^(\d+(?:,\d+)*(?:\.\d+)?)\s*$', lines[j])
                    if nm:
                        val = float(nm.group(1).replace(',', ''))
                        if mon not in found:
                            found[mon] = val
                        break

    return [(mon, found[mon]) for mon in MONTHS_ORDER if mon in found]


def _extract_with_fallback(pdf, page_num, key, log):
    """
    Try spatial extraction first, fall back to text-based.
    page_num is 1-indexed.
    Returns list of (month, value) tuples.
    """
    if page_num > len(pdf):
        return []

    page = pdf.load_page(page_num - 1)

    # Strategy 1: Spatial analysis (most accurate)
    vals = _extract_month_values_spatial(page)
    if vals:
        log.append(f"INFO | {key}: Spatial extraction OK ({len(vals)} points)")
        return vals

    # Strategy 2: Text-based fallback
    raw_text = page.get_text()
    vals = _extract_month_values_text(raw_text)
    if vals:
        log.append(f"INFO | {key}: Text fallback OK ({len(vals)} points)")
        return vals

    # Nothing worked - dump raw text for debugging
    snippet = raw_text[:300].replace('\n', ' | ')
    log.append(f"INFO | {key}: NO DATA. Raw: '{snippet}...'")
    return []


def _replace_in_para(para, pattern, replacement, flags=re.IGNORECASE):
    """Replace text matching pattern in a paragraph, keeping first run's formatting."""
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
    """Replace month/year date references across ALL slides."""
    MONTH_NAMES = ['January', 'February', 'March', 'April', 'May', 'June',
                   'July', 'August', 'September', 'October', 'November', 'December']
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
                if _replace_in_para(para, rf'({month_pat})\s+(20\d{{2}})', f'{new_month} {new_year}'):
                    changes.append(f"  DATE | Slide {slide_num + 1}: '{original.strip()[:60]}'")
                    continue
                if _replace_in_para(para, r'\b(20\d{2})\b', str(new_year)):
                    changes.append(f"  DATE | Slide {slide_num + 1}: year updated")
    return changes


def _analyze_trend(values):
    """
    Deep trend analysis over available data points.
    Returns dict with: direction, diff, pct_change, streak, trend_desc, is_peak, is_trough
    """
    if len(values) < 2:
        return None

    curr = int(round(values[-1][1]))
    prev = int(round(values[-2][1]))
    diff = abs(curr - prev)
    direction = "increased" if curr > prev else "decreased" if curr < prev else "unchanged"

    # Percentage change
    if prev > 0:
        pct = round((diff / prev) * 100)
    else:
        pct = 0

    # Consecutive streak analysis (how many months in same direction)
    streak = 1
    if len(values) >= 3:
        for k in range(len(values) - 2, 0, -1):
            v_curr = values[k][1]
            v_prev = values[k - 1][1]
            if direction == "increased" and v_curr > v_prev:
                streak += 1
            elif direction == "decreased" and v_curr < v_prev:
                streak += 1
            else:
                break

    # Peak / trough detection
    all_vals = [v[1] for v in values]
    is_peak = curr == max(all_vals) and curr > prev
    is_trough = curr == min(all_vals) and curr < prev

    return {
        "curr": curr,
        "prev": prev,
        "diff": diff,
        "pct": pct,
        "direction": direction,
        "streak": streak,
        "is_peak": is_peak,
        "is_trough": is_trough,
    }


def _make_smart_ticket_bullet(vals, sel_month, sel_year):
    """Generate an intelligent ticket trend summary bullet."""
    if not vals or len(vals) < 2:
        return None

    a = _analyze_trend(vals)
    if not a:
        return None

    curr, prev, diff, pct = a["curr"], a["prev"], a["diff"], a["pct"]
    direction = a["direction"]

    if direction == "unchanged":
        return f"Ticket volume remained steady at {curr} in {sel_month} {sel_year}"

    verb = "increased" if direction == "increased" else "decreased"

    # Build the core sentence
    parts = [f"Ticket logged {verb} by {diff}"]

    # Add percentage if meaningful
    if pct > 0 and prev > 0:
        parts[0] += f" ({pct}%)"

    parts[0] += f" in {sel_month} {sel_year}, from {prev} to {curr}"

    # Add streak context if notable
    if a["streak"] >= 3:
        parts.append(f"This marks {a['streak']} consecutive months of {verb[:-1] if verb.endswith('d') else verb}ing trend")
    elif a["streak"] == 2:
        parts.append(f"continuing the {verb[:-1] if verb.endswith('d') else verb}ing trend from last month")

    # Add peak/trough context
    if a["is_peak"]:
        parts.append("reaching the highest level in the observed period")
    elif a["is_trough"]:
        parts.append("reaching the lowest level in the observed period")

    return ". ".join(parts)


def _make_smart_ageing_bullet(vals, sel_month, sel_year):
    """Generate an intelligent ageing trend summary bullet."""
    if not vals or len(vals) < 2:
        return None

    a = _analyze_trend(vals)
    if not a:
        return None

    curr, prev, diff, pct = a["curr"], a["prev"], a["diff"], a["pct"]
    direction = a["direction"]

    if direction == "unchanged":
        if curr == 0:
            return f"No ageing tickets recorded in {sel_month} {sel_year}, maintaining zero backlog"
        return f"Ageing ticket count remained at {curr} in {sel_month} {sel_year}"

    if curr == 0 and prev > 0:
        return f"All ageing tickets resolved - count dropped from {prev} to 0 in {sel_month} {sel_year}"

    verb = "increased" if direction == "increased" else "decreased"
    sign = "+" if direction == "increased" else "-"

    parts = [f"Ageing ticket {verb} from {prev} to {curr} in {sel_month} {sel_year} ({sign}{diff})"]

    # Add percentage context for significant changes
    if pct >= 50 and diff >= 3:
        parts.append(f"a significant {pct}% {'rise' if direction == 'increased' else 'reduction'}")
    elif pct > 0 and prev > 0:
        parts[0] = parts[0].replace(f"({sign}{diff})", f"({sign}{diff}, {pct}%)")

    # Streak context
    if a["streak"] >= 3:
        parts.append(f"continuing a {a['streak']}-month {'upward' if direction == 'increased' else 'downward'} trend")

    return ". ".join(parts)


def update_summary_bullets(prs, sel_month, sel_year,
                            sr_trend_vals, sr_ageing_vals,
                            inc_trend_vals, inc_ageing_vals):
    """
    Update summary bullets on Slides 4 and 8 with smart analysis.
    SAFETY: Only modifies paragraphs inside text frames that
    contain a 'Summary' heading. Title text frames are NEVER touched.
    """
    changes = []

    CONFIG = [
        (3, sr_trend_vals,  sr_ageing_vals,  "SR"),
        (7, inc_trend_vals, inc_ageing_vals, "INC"),
    ]

    for slide_idx, ticket_vals, ageing_vals, label in CONFIG:
        if slide_idx >= len(prs.slides):
            continue
        slide = prs.slides[slide_idx]
        ticket_bullet = _make_smart_ticket_bullet(ticket_vals, sel_month, sel_year)
        ageing_bullet = _make_smart_ageing_bullet(ageing_vals, sel_month, sel_year)
        ticket_done = False
        ageing_done = False

        for shape in slide.shapes:
            if not hasattr(shape, "text_frame"):
                continue
            tf = shape.text_frame

            # -- SAFETY: Only touch text frames containing "Summary" --
            has_summary = any(
                'summary' in "".join(r.text for r in p.runs).strip().lower()
                for p in tf.paragraphs
            )
            if not has_summary:
                continue

            for para in tf.paragraphs:
                if not para.runs:
                    continue
                pt = "".join(r.text for r in para.runs)
                ptl = pt.lower()
                if 'summary' in ptl and len(ptl) < 20:
                    continue
                if not ticket_done and ticket_bullet and 'ticket' in ptl:
                    para.runs[0].text = ticket_bullet
                    for r in para.runs[1:]:
                        r.text = ""
                    changes.append(f"  SUMM | Slide {slide_idx+1}: '{ticket_bullet}'")
                    ticket_done = True
                elif not ageing_done and ageing_bullet and ('ageing' in ptl or 'aging' in ptl):
                    para.runs[0].text = ageing_bullet
                    for r in para.runs[1:]:
                        r.text = ""
                    changes.append(f"  SUMM | Slide {slide_idx+1}: '{ageing_bullet}'")
                    ageing_done = True

        if not ticket_done:
            changes.append(f"  SUMM | Slide {slide_idx+1}: {label} ticket - {'no data' if not ticket_vals else 'no match'}")
        if not ageing_done:
            changes.append(f"  SUMM | Slide {slide_idx+1}: {label} ageing - {'no data' if not ageing_vals else 'no match'}")

    return changes


# ==============================================================
# PROCESSING ENGINE v3 - HANDLES PICTURE, CHART, AND TABLE
# ==============================================================
def process_monthly_report(pdf_bytes, pptx_bytes, sel_month, sel_year):
    """
    Full automation engine.
    """
    from collections import defaultdict

    # -- Phase 1: Extract PDF page images --
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pdf_images = {}
    for pn in range(len(pdf)):
        page = pdf.load_page(pn)
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
        pdf_images[pn + 1] = pix.tobytes("png")
    log = [f"INFO | PDF loaded: {len(pdf)} pages at 300 DPI"]

    # -- Phase 2: Extract trend data (Smart Spatial Extraction) --
    sr_trend  = _extract_with_fallback(pdf, 3, "sr_trend", log)
    sr_ageing = _extract_with_fallback(pdf, 4, "sr_ageing", log)
    inc_trend = _extract_with_fallback(pdf, 10, "inc_trend", log)
    inc_ageing = _extract_with_fallback(pdf, 11, "inc_ageing", log)

    # Detailed logging for verification
    for lbl, vals in [
        ("SR Trend (p3)",    sr_trend),
        ("SR Ageing (p4)",   sr_ageing),
        ("INC Trend (p10)",  inc_trend),
        ("INC Ageing (p11)", inc_ageing),
    ]:
        if vals:
            pairs = ", ".join(f"{m}={int(v)}" for m, v in vals)
            log.append(f"INFO | Extract {lbl}: {pairs}")

    # -- Phase 3: Auto-crop module list pages --
    for pg in (7, 14):
        if pg in pdf_images:
            old = len(pdf_images[pg])
            pdf_images[pg] = _auto_crop_bottom(pdf_images[pg])
            nw = len(pdf_images[pg])
            if nw < old:
                log.append(f"INFO | PDF p{pg}: cropped ({old//1024}KB -> {nw//1024}KB)")

    # -- Phase 4: Open PPTX --
    prs = Presentation(io.BytesIO(pptx_bytes))
    log.append(f"INFO | PPTX: {len(prs.slides)} slides")

    # -- Phase 5: Replacement map --
    # (slide_idx_0, pdf_page, strategy, hint, label)
    MAP = [
        (2,  2, "BBOX", (0.314, 2.276, 3.782, 1.200), "SR SLA Performance"),
        (2,  3, "BBOX", (4.570, 2.276, 8.307, 3.286), "SR Ticket Trend"),
        (3,  4, "CHART", None, "SR Ageing Trend"),
        (4,  5, "TABLE", None, "SR Root Cause Table"),
        (5,  6, "BBOX", (0.488, 2.641, 6.533, 2.449), "SR Category Distribution"),
        (5,  7, "BBOX", (7.346, 1.448, 3.477, 5.459), "SR Module Ticket List"),
        (6,  8, "BBOX", (0.404, 2.406, 3.240, 1.313), "INC Response SLA"),
        (6,  9, "BBOX", (0.315, 4.294, 3.417, 1.365), "INC Resolution SLA"),
        (6, 10, "BBOX", (3.932, 2.406, 9.124, 3.567), "INC Ticket Trend"),
        (7, 11, "CHART", None, "INC Ageing Trend"),
        (8, 12, "TABLE", None, "INC Root Cause Table"),
        (9, 13, "POS", 0, "INC Category Distribution"),
        (9, 14, "POS", 1, "INC Module Ticket List"),
    ]

    LOGO_TOP = 0.5
    LOGO_H   = 0.6
    replaced  = 0

    groups = defaultdict(list)
    for e in MAP:
        groups[e[0]].append(e)

    for si in sorted(groups.keys()):
        entries = groups[si]
        if si >= len(prs.slides):
            log.append(f"WARN | Slide {si+1} missing")
            continue

        slide = prs.slides[si]

        # Collect non-logo PICTURE shapes
        pics = []
        for sh in slide.shapes:
            if sh.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            if _emu_to_inches(sh.top) < LOGO_TOP and _emu_to_inches(sh.height) < LOGO_H:
                continue
            pics.append(sh)

        used = set()
        jobs = []  # (entry, left, top, width, height, shape_to_delete)

        for entry in entries:
            _, pp, strat, hint, label = entry
            if pp not in pdf_images:
                log.append(f"WARN | PDF p{pp} missing. Skip '{label}'.")
                continue

            target = None

            if strat == "BBOX":
                tl, tt, tw, th = hint
                best, bdist = None, float('inf')
                for sh in pics:
                    if id(sh) in used:
                        continue
                    d = (((_emu_to_inches(sh.left) - tl)**2) +
                         ((_emu_to_inches(sh.top)  - tt)**2) +
                         ((_emu_to_inches(sh.width) - tw)**2) +
                         ((_emu_to_inches(sh.height)- th)**2)) ** 0.5
                    if d < bdist:
                        bdist, best = d, sh
                if best and bdist < 3.0:
                    target = best
                    used.add(id(best))
                else:
                    log.append(f"WARN | Slide {si+1}: No PICTURE for '{label}' (dist={bdist:.2f})")

            elif strat == "CHART":
                for sh in slide.shapes:
                    try:
                        if sh.shape_type == MSO_SHAPE_TYPE.CHART:
                            target = sh
                            break
                    except:
                        pass
                if not target:
                    ba = 0
                    for sh in slide.shapes:
                        try:
                            st = sh.shape_type
                        except:
                            continue
                        if st == MSO_SHAPE_TYPE.PICTURE or st == 17:
                            continue
                        ti = _emu_to_inches(sh.top)
                        hi = _emu_to_inches(sh.height)
                        if ti < LOGO_TOP and hi < LOGO_H:
                            continue
                        if _emu_to_inches(sh.width) < 1.0 or hi < 0.5:
                            continue
                        a = sh.width * sh.height
                        if a > ba:
                            ba, target = a, sh
                if not target:
                    log.append(f"WARN | Slide {si+1}: No CHART for '{label}'")

            elif strat == "TABLE":
                for sh in slide.shapes:
                    try:
                        if sh.has_table:
                            target = sh
                            break
                    except:
                        pass
                if not target:
                    for sh in slide.shapes:
                        try:
                            if sh.shape_type == MSO_SHAPE_TYPE.TABLE:
                                target = sh
                                break
                        except:
                            pass
                if not target:
                    ba = 0
                    for sh in slide.shapes:
                        try:
                            st = sh.shape_type
                        except:
                            continue
                        if st == MSO_SHAPE_TYPE.PICTURE or st == 17:
                            continue
                        ti = _emu_to_inches(sh.top)
                        hi = _emu_to_inches(sh.height)
                        if ti < LOGO_TOP and hi < LOGO_H:
                            continue
                        if _emu_to_inches(sh.width) < 1.0 or hi < 0.5:
                            continue
                        a = sh.width * sh.height
                        if a > ba:
                            ba, target = a, sh
                if not target:
                    log.append(f"WARN | Slide {si+1}: No TABLE for '{label}'")

            elif strat == "POS":
                avail = sorted([s for s in pics if id(s) not in used], key=lambda s: s.left)
                idx = hint
                if idx < len(avail):
                    target = avail[idx]
                    used.add(id(target))
                else:
                    log.append(f"WARN | Slide {si+1}: Only {len(avail)} pics, need #{idx} for '{label}'")

            if target:
                jobs.append((entry, target.left, target.top, target.width, target.height, target))

        # Delete old shapes first, then insert new images
        for _, _, _, _, _, sh in jobs:
            try:
                sh._element.getparent().remove(sh._element)
            except Exception as ex:
                log.append(f"WARN | Delete failed: {ex}")

        for entry, left, top, width, height, _ in jobs:
            pp, label, strat = entry[1], entry[4], entry[2]
            img_bytes = pdf_images[pp]

            # -- Smart placement based on strategy --
            if strat in ("CHART", "TABLE"):
                # Aspect-ratio fit, centred inside the bounding box.
                # Prevents stretching when chart/table box dimensions differ
                # from the PDF page content's natural proportions.
                iw, ih = _get_image_size(img_bytes)
                ins_l, ins_t, ins_w, ins_h = _calc_fit_rect(iw, ih, left, top, width, height)
            elif strat == "POS":
                # Width-constrained: use full box width, height follows the
                # already-cropped image. Bulletproof for variable-length lists.
                iw, ih = _get_image_size(img_bytes)
                ins_l, ins_t, ins_w, ins_h = _calc_list_rect(iw, ih, left, top, width)
            else:
                # BBOX: template has been manually adjusted; use exact box.
                ins_l, ins_t, ins_w, ins_h = left, top, width, height

            slide.shapes.add_picture(io.BytesIO(img_bytes), ins_l, ins_t, ins_w, ins_h)
            b = (_emu_to_inches(ins_l), _emu_to_inches(ins_t),
                 _emu_to_inches(ins_w), _emu_to_inches(ins_h))
            log.append(f"  OK | Slide {entry[0]+1}: '{label}' [{strat}] <- PDF p{pp} "
                       f"(L={b[0]:.2f} T={b[1]:.2f} W={b[2]:.2f} H={b[3]:.2f})")
            replaced += 1

    # -- Phase 6: Date replacement --
    log.append("INFO | Updating dates...")
    dc = update_dates_in_pptx(prs, sel_month, str(sel_year))
    log.extend(dc)
    log.append(f"INFO | {len(dc)} date updates.")

    # -- Phase 7: Summary bullets (SAFE) --
    log.append("INFO | Updating summary on slides 4 and 8...")
    sc = update_summary_bullets(prs, sel_month, sel_year, sr_trend, sr_ageing, inc_trend, inc_ageing)
    log.extend(sc)

    # -- Phase 8: Save --
    out = io.BytesIO()
    prs.save(out)
    out.seek(0)
    log.append(f"INFO | Done: {replaced}/13 images replaced.")
    return out.read(), log, replaced, pdf_images


# -- Resolve PPTX template --
resolved_pptx = None
if uploaded_template is not None:
    resolved_pptx = "uploaded"
elif pptx_available_locally:
    resolved_pptx = "local"


# ==============================================================
# GENERATE SECTION
# ==============================================================
if pdf_file and resolved_pptx:
    st.markdown('<p class="section-label">Step 2 - Validation and Generation</p>', unsafe_allow_html=True)

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
        st.metric("Target Slides", "8 (Slides 3-10)")
    with m3:
        st.metric("Image Swaps", "13")

    with st.expander("PDF to PPT Mapping", expanded=False):
        st.markdown("""
| PDF | Content | Slide | Method | Notes |
|:---:|---|:---:|---|---|
| 1 | Landing | *Skip* | - | - |
| **2** | SR SLA | **3** | BBOX match | Left picture |
| **3** | SR Trend | **3** | BBOX match | Right picture |
| **4** | SR Ageing | **4** | Delete chart | Replaces native histogram |
| **5** | SR Root Cause | **5** | Delete table | Replaces native table |
| **6** | SR Category % | **6** | BBOX match | Left picture |
| **7** | SR Module List | **6** | BBOX match | Right picture (auto-cropped) |
| **8** | INC Response SLA | **7** | BBOX match | Top-left picture |
| **9** | INC Resolution SLA | **7** | BBOX match | Bottom-left picture |
| **10** | INC Trend | **7** | BBOX match | Right picture |
| **11** | INC Ageing | **8** | Delete chart | Replaces native histogram |
| **12** | INC Root Cause | **9** | Delete table | Replaces native table |
| **13** | INC Category % | **10** | Position sort | Left picture |
| **14** | INC Module List | **10** | Position sort | Right picture (auto-cropped) |
        """)
        st.info(
            "**Automated features:**\n"
            "- Module lists (pages 7 and 14) are auto-cropped to remove whitespace\n"
            "- Dates are updated across all slides\n"
            "- Summary bullets on Slides 4 and 8 are auto-computed from extracted chart data\n"
            "- Corporate logo is protected and title text is never modified"
        )

    if pdf_page_count != "?" and pdf_page_count < 14:
        st.warning(f"Expected 14 pages but found {pdf_page_count}. Some images may be missing.")

    st.markdown("")

    if st.button("Generate Monthly Report", use_container_width=True, type="primary"):
        with st.spinner("Extracting visuals, replacing charts and tables, updating text..."):
            try:
                if uploaded_template is not None:
                    template_bytes = uploaded_template.read()
                else:
                    with open(TEMPLATE_PATH, "rb") as f:
                        template_bytes = f.read()

                out_bytes, build_logs, img_count, pdf_imgs = process_monthly_report(
                    pdf_file.read(), template_bytes, sel_month, int(sel_year)
                )

                if img_count == 13:
                    st.success(f"Complete - all {img_count}/13 images replaced and text updated.")
                elif img_count > 0:
                    st.warning(f"Partial - {img_count}/13 images replaced. Review the build log.")
                else:
                    st.error("No images were replaced. Please check your input files.")

                with st.expander("Build Log", expanded=(img_count < 13)):
                    for msg in build_logs:
                        if "WARN" in msg:
                            st.warning(msg)
                        elif msg.startswith("  OK") or msg.startswith("  DATE") or msg.startswith("  SUMM"):
                            st.text(msg)
                        else:
                            st.info(msg)

                # -- Preview --
                if img_count > 0 and pdf_imgs:
                    st.markdown("")
                    st.markdown('<p class="section-label">Preview - Images Placed Per Slide</p>', unsafe_allow_html=True)
                    PREVIEW = [
                        ("Slide 3 - SR SLA and Trend",               [2, 3]),
                        ("Slide 4 - SR Ageing (chart replaced)",      [4]),
                        ("Slide 5 - SR Root Cause (table replaced)",  [5]),
                        ("Slide 6 - SR Category and Module List",     [6, 7]),
                        ("Slide 7 - INC SLA and Trend",              [8, 9, 10]),
                        ("Slide 8 - INC Ageing (chart replaced)",     [11]),
                        ("Slide 9 - INC Root Cause (table replaced)", [12]),
                        ("Slide 10 - INC Category and Module List",   [13, 14]),
                    ]
                    for title, pages in PREVIEW:
                        avail = [p for p in pages if p in pdf_imgs]
                        if not avail:
                            continue
                        with st.expander(title, expanded=False):
                            cols = st.columns(len(avail))
                            for col, pg in zip(cols, avail):
                                with col:
                                    st.caption(f"PDF Page {pg}")
                                    st.image(pdf_imgs[pg], use_container_width=True)

                # -- Download --
                if img_count > 0:
                    st.markdown("")
                    st.markdown('<p class="section-label">Step 3 - Download</p>', unsafe_allow_html=True)
                    st.info(
                        "**What was automatically updated:**\n"
                        "- All 13 dashboard images replaced\n"
                        "- Charts on Slides 4 and 8 replaced with PDF images\n"
                        "- Tables on Slides 5 and 9 replaced with PDF images\n"
                        "- Dates updated across all slides\n"
                        "- Summary bullets on Slides 4 and 8 computed from chart data (if extractable)\n\n"
                        "Review summary text on Slides 4 and 8 if the build log shows warnings."
                    )
                    st.download_button(
                        label="Download Final Report (.pptx)",
                        data=out_bytes,
                        file_name=f"Monthly_Report_{sel_month}_{sel_year}.pptx",
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                        use_container_width=True
                    )

            except Exception as e:
                st.error(f"An error occurred: {e}")
                st.code(traceback.format_exc(), language="text")

elif not resolved_pptx and pdf_file:
    st.warning("No PPTX template found. Please upload one in the sidebar.")

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
            Select the Report Month and Year, then upload your Power BI PDF export in the sidebar.
        </p>
    </div>
    """, unsafe_allow_html=True)
