import streamlit as st
import io
import traceback
import sys
import os
import base64
import datetime

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

    /* KILL ALL BLINKING */
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
    
    /* Ensure no transitions that look like blinks — but leave sidebar alone */
    [data-testid="stAppViewContainer"] > .main { transition: none !important; }
    
    [data-testid="stSidebar"] {
        animation: none !important;
    }
    [data-testid="stSidebarNav"],
    [data-testid="stSidebarNavItems"],
    [data-testid="stSidebarNavSeparator"],
    [data-testid="stStatusWidget"] {
        display: none !important;
        visibility: hidden !important;
    }
    header[data-testid="stHeader"] {
        background: #F8FAFC !important;
    }
    [data-testid="stSidebar"] { border-right: none !important; }

    .main .block-container {
        padding-top: 1rem !important;
        max-width: 1200px !important;
    }
    [data-testid="stSidebar"] { border-right: none !important; }
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

    /* Tighten Sidebar Spacing */
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 {
        margin-bottom: 0px !important;
    }
    [data-testid="stSidebar"] .stTextInput { margin-bottom: -10px !important; }

    /* Hide Streamlit chrome */
    .stDeployButton, [data-testid="stDeployButton"], [data-testid="stAppDeployButton"] { display: none !important; }
    #MainMenu { visibility: hidden; } footer { visibility: hidden; }

    /* Section dividers */
    .section-label {
        font-size: 0.7rem;
        font-weight: 700;
        text-transform: uppercase;
        letter-spacing: 1px;
        color: #94A3B8;
        margin-bottom: 12px;
        margin-top: 8px;
    }

    /* ── Hide default Streamlit sidebar nav ── */
    [data-testid="stSidebarNav"] { display: none !important; }

    /* ── Custom Sidebar Nav ── */
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
    .sidebar-sep {
        border: none;
        border-top: 1px solid #E2E8F0;
        margin: 16px 12px;
    }
    .genie-link {
        font-size: 0.85rem;
        font-weight: 500;
        color: #31333F !important;
        text-decoration: none !important;
        transition: all 0.2s ease !important;
        cursor: pointer !important;
    }
    .genie-link:hover {
        color: #00B1A9 !important;
        text-decoration: none !important;
    }
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
    _months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    _now = datetime.date.today()
    sel_month = st.selectbox(
        "Report Month",
        options=_months,
        index=_now.month - 1
    )
    _current_year = _now.year
    _years = [str(y) for y in range(_current_year - 1, _current_year + 3)]
    sel_year = st.selectbox(
        "Report Year",
        options=_years,
        index=1
    )

    st.markdown("<div style='margin-top: -10px;'></div>", unsafe_allow_html=True)
    st.markdown("### Data Upload")
    st.markdown("<a href='#' target='_blank' class='genie-link'>Power BI PDF Export ↗</a>", unsafe_allow_html=True)
    pdf_file = st.file_uploader("Power BI PDF Export", type=['pdf'], label_visibility="collapsed")

    # PPTX Fallback: if template.pptx not found, let user upload it
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


# ── EMU conversion helper ──
def _emu_to_inches(emu):
    """PowerPoint uses English Metric Units: 1 inch = 914400 EMU."""
    return emu / 914400.0


# ══════════════════════════════════════════════════════════════
# PROCESSING ENGINE (Spec-Compliant)
# Based on engineering-grade analysis of the 11-slide PPTX
# template and 14-page Power BI PDF export.
# ══════════════════════════════════════════════════════════════
def process_monthly_report(pdf_bytes, pptx_bytes):
    """
    Precision automation engine for HCM Dashboard PPTX replacement.

    Only Slides 3, 6, 7, and 10 contain replaceable dashboard pictures.
    Slides 1, 2, 4, 5, 8, 9, 11 are NEVER touched.
    Template logo (Picture 3, ~2.3x0.4 at top-left) is ALWAYS protected.

    PDF pages used: 2, 3, 6, 7, 8, 9, 10, 13 (8 total out of 14).
    PDF pages skipped: 1 (landing), 4, 5, 11, 12 (native charts/tables), 14 (no target).
    """

    # ── Phase 1: Extract all PDF pages as high-res PNGs ──
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pdf_images = {}  # 1-indexed
    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))  # 4x zoom ~ 300 DPI
        pdf_images[page_num + 1] = pix.tobytes("png")

    log = []
    log.append(f"INFO | PDF loaded: {len(pdf)} pages extracted at 300 DPI")

    # ── Phase 2: Open the PPTX template ──
    prs = Presentation(io.BytesIO(pptx_bytes))
    log.append(f"INFO | PPTX loaded: {len(prs.slides)} slides in template")

    # ── Phase 3: Precise replacement mapping ──
    # (slide_index_0based, pdf_page_1indexed, target_bbox_inches, label)
    # bbox = (left, top, width, height)
    REPLACEMENT_MAP = [
        # Slide 3 (idx 2): Service Request — SLA + Ticket Trend
        (2,  2, (0.314, 2.276, 3.782, 1.200), "SR SLA Performance"),
        (2,  3, (4.570, 2.276, 8.307, 3.286), "SR Ticket Trend"),
        # Slide 6 (idx 5): Service Request — Category % + Module List
        (5,  6, (0.488, 2.641, 6.533, 2.449), "SR Category Distribution"),
        (5,  7, (7.346, 1.448, 3.477, 5.459), "SR Module Ticket List"),
        # Slide 7 (idx 6): Incident — Response + Resolution + Trend
        (6,  8, (0.404, 2.406, 3.240, 1.313), "INC Response SLA"),
        (6,  9, (0.315, 4.294, 3.417, 1.365), "INC Resolution SLA"),
        (6, 10, (3.932, 2.406, 9.124, 3.567), "INC Ticket Trend"),
        # Slide 10 (idx 9): Incident — Category Distribution
        (9, 13, (0.862, 2.647, 11.700, 2.524), "INC Category Distribution"),
    ]

    # ── Phase 4: Logo protection ──
    # Template logo (Picture 3): L=0.000, T=0.081, W=2.299, H=0.402
    LOGO_TOP_MAX = 0.5
    LOGO_HEIGHT_MAX = 0.6

    # ── Phase 5: Execute replacements ──
    replaced = 0
    TOLERANCE = 0.15  # inches for bbox matching

    for slide_idx, pdf_page, target_bbox, label in REPLACEMENT_MAP:
        if slide_idx >= len(prs.slides):
            log.append(f"WARN | Slide {slide_idx + 1} does not exist. Skipping '{label}'.")
            continue

        if pdf_page not in pdf_images:
            log.append(f"WARN | PDF page {pdf_page} not found. Skipping '{label}'.")
            continue

        slide = prs.slides[slide_idx]
        t_left, t_top, t_width, t_height = target_bbox

        # Collect PICTURE shapes, excluding the template logo
        candidates = []
        for shape in slide.shapes:
            if shape.shape_type != MSO_SHAPE_TYPE.PICTURE:
                continue
            s_top = _emu_to_inches(shape.top)
            s_height = _emu_to_inches(shape.height)
            if s_top < LOGO_TOP_MAX and s_height < LOGO_HEIGHT_MAX:
                continue
            candidates.append(shape)

        if not candidates:
            log.append(f"WARN | Slide {slide_idx + 1}: No replaceable pictures found for '{label}'.")
            continue

        # Find the picture closest to the target bbox
        best_shape = None
        best_distance = float('inf')

        for shape in candidates:
            s_left = _emu_to_inches(shape.left)
            s_top = _emu_to_inches(shape.top)
            s_width = _emu_to_inches(shape.width)
            s_height = _emu_to_inches(shape.height)

            dist = (
                (s_left - t_left) ** 2 +
                (s_top - t_top) ** 2 +
                (s_width - t_width) ** 2 +
                (s_height - t_height) ** 2
            ) ** 0.5

            if dist < best_distance:
                best_distance = dist
                best_shape = shape

        if best_shape is None or best_distance > TOLERANCE * 4:
            log.append(
                f"WARN | Slide {slide_idx + 1}: No picture matched bbox for '{label}' "
                f"(closest dist: {best_distance:.3f}). Skipping."
            )
            continue

        # Insert new image at the old shape's exact position/size
        img_io = io.BytesIO(pdf_images[pdf_page])
        slide.shapes.add_picture(
            img_io,
            best_shape.left,
            best_shape.top,
            best_shape.width,
            best_shape.height
        )

        # Remove the old shape
        sp = best_shape._element
        sp.getparent().remove(sp)

        actual_bbox = (
            _emu_to_inches(best_shape.left),
            _emu_to_inches(best_shape.top),
            _emu_to_inches(best_shape.width),
            _emu_to_inches(best_shape.height),
        )
        log.append(
            f"  OK | Slide {slide_idx + 1}: '{label}' <- PDF page {pdf_page} "
            f"(L={actual_bbox[0]:.2f} T={actual_bbox[1]:.2f} "
            f"W={actual_bbox[2]:.2f} H={actual_bbox[3]:.2f}, dist={best_distance:.3f})"
        )
        replaced += 1

    # ── Phase 6: Save ──
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    log.append(f"INFO | Complete: {replaced}/8 dashboard images replaced successfully.")
    return output.read(), log, replaced


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
        st.metric("Target Slides", "4 (3, 6, 7, 10)")
    with m3:
        st.metric("Image Swaps", "8")

    with st.expander("PDF to PPT Mapping (what goes where)", expanded=False):
        st.markdown("""
| PDF Page | Content | PPT Slide | Target |
|:---:|---|:---:|---|
| ~~1~~ | *Landing/Filters* | *Skip* | *Not used* |
| **2** | SR SLA Performance | **Slide 3** | Left image |
| **3** | SR Ticket Trend | **Slide 3** | Right image |
| ~~4~~ | *SR Ageing Trend* | *Skip* | *Native PPT chart* |
| ~~5~~ | *SR Root Cause Table* | *Skip* | *Native PPT table* |
| **6** | SR Category % | **Slide 6** | Left image |
| **7** | SR Module List | **Slide 6** | Right image |
| **8** | INC Response SLA | **Slide 7** | Top-left image |
| **9** | INC Resolution SLA | **Slide 7** | Bottom-left image |
| **10** | INC Ticket Trend | **Slide 7** | Right image |
| ~~11~~ | *INC Ageing Trend* | *Skip* | *Native PPT chart* |
| ~~12~~ | *INC Root Cause Table* | *Skip* | *Native PPT table* |
| **13** | INC Category % | **Slide 10** | Full-width image |
| ~~14~~ | *INC Module List* | *Skip* | *No target* |
        """)

    if pdf_page_count != "?" and pdf_page_count < 13:
        st.warning(
            f"Expected at least 13 pages in the PDF, but found {pdf_page_count}. "
            "Some dashboard images may not be replaced."
        )

    st.markdown("")

    if st.button("Generate Monthly Report", use_container_width=True, type="primary"):
        with st.spinner("Extracting visuals and building the presentation..."):
            try:
                if uploaded_template is not None:
                    template_bytes = uploaded_template.read()
                else:
                    with open(TEMPLATE_PATH, "rb") as f:
                        template_bytes = f.read()

                out_bytes, build_logs, img_count = process_monthly_report(
                    pdf_file.read(),
                    template_bytes
                )

                if img_count == 8:
                    st.success(f"Perfect — all {img_count}/8 dashboard images replaced successfully.")
                elif img_count > 0:
                    st.warning(f"Partial success — {img_count}/8 images replaced. Check the build log.")
                else:
                    st.error("No images were replaced. Check your PDF and PPTX template.")

                with st.expander("Build Log", expanded=(img_count < 8)):
                    for msg in build_logs:
                        if msg.startswith("WARN"):
                            st.warning(msg)
                        elif msg.startswith("  OK"):
                            st.text(msg)
                        else:
                            st.info(msg)

                if img_count > 0:
                    st.markdown("")
                    st.markdown('<p class="section-label">Step 3 &mdash; Download</p>', unsafe_allow_html=True)

                    st.info(
                        "**Manual edits still required on Slides 4 and 8:**\n"
                        "- Update the ageing chart data (native PPT chart)\n"
                        "- Update the summary bullet text (ticket counts and ageing numbers)\n\n"
                        "**Slides 5 and 9** have native tables — update row values manually if needed.",
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
    st.warning("No PowerPoint template available. Please upload a PPTX template in the sidebar.")

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
