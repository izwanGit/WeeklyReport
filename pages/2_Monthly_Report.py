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
    page_icon="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png",
    layout="wide",
)

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
</style>
""", unsafe_allow_html=True)


# ── Helpers ──
def _image_to_data_uri(path, mime_type):
    try:
        with open(os.path.join(BASE_DIR, path), 'rb') as f:
            data = f.read()
        return f"data:{mime_type};base64,{base64.b64encode(data).decode()}"
    except:
        return ""

_logo_banner_uri = _image_to_data_uri("PETRONAS_LOGO_SQUARE.png", "image/png")
_logo_sidebar_uri = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg", "image/svg+xml")


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
    pdf_file = st.file_uploader("Power BI PDF Export", type=['pdf'])

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
<img src="{_logo_banner_uri}" style="height: 80px; flex-shrink: 0; filter: drop-shadow(1px 1px 0 white) drop-shadow(-1px -1px 0 white) drop-shadow(1px -1px 0 white) drop-shadow(-1px 1px 0 white);" />
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


TEMPLATE_PATH = os.path.join(BASE_DIR, "template.pptx")
pptx_available = os.path.exists(TEMPLATE_PATH)


# ── Processing Engine ──
def process_monthly_report(pdf_bytes, pptx_bytes):
    """
    Core automation engine.
    1. Extracts each PDF page as a high-resolution PNG (4x zoom = ~300 DPI).
    2. Opens the PPTX template.
    3. Replaces image placeholders on slides 3-10 with sequentially mapped PDF pages.
    4. Returns the final PPTX bytes and a structured build log.
    """

    # ── Phase 1: Extract high-res images from the Power BI PDF ──
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pdf_images = []
    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)
        # 4x zoom matrix produces extremely crisp output for presentation use
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
        pdf_images.append(pix.tobytes("png"))

    # ── Phase 2: Open the PPTX template ──
    prs = Presentation(io.BytesIO(pptx_bytes))

    # ── Phase 3: Image replacement engine ──
    # Slide index (0-based) → number of images to replace on that slide
    # Slide 3 = index 2, Slide 4 = index 3, etc.
    mapping = {
        2: 2,  # Slide 3:  SLA + Ticket Trend
        3: 1,  # Slide 4:  1 histogram visual
        4: 1,  # Slide 5:  1 table (as image)
        5: 2,  # Slide 6:  2 visuals
        6: 3,  # Slide 7:  3 visuals
        7: 1,  # Slide 8:  1 histogram visual
        8: 1,  # Slide 9:  1 table (as image)
        9: 2,  # Slide 10: 2 visuals
    }

    pdf_idx = 0
    log = []

    for slide_idx, num_images in mapping.items():
        if slide_idx >= len(prs.slides):
            log.append(f"WARN | Slide {slide_idx + 1} does not exist in the template. Skipped.")
            continue

        slide = prs.slides[slide_idx]

        # Collect all picture shapes on this slide
        target_shapes = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                target_shapes.append(shape)
            elif getattr(shape, 'is_placeholder', False):
                try:
                    if shape.placeholder_format.type == 18:  # Picture placeholder
                        target_shapes.append(shape)
                except:
                    pass

        # Sort top-to-bottom, left-to-right for deterministic sequencing
        target_shapes.sort(key=lambda s: (s.top, s.left))

        for k in range(num_images):
            if pdf_idx >= len(pdf_images):
                log.append(f"WARN | Ran out of PDF pages at page {pdf_idx + 1}. Remaining slides were not updated.")
                break

            if k >= len(target_shapes):
                log.append(f"WARN | Slide {slide_idx + 1}: Expected {num_images} image(s) but found only {len(target_shapes)}. Partial replacement.")
                break

            old_shape = target_shapes[k]
            img_io = io.BytesIO(pdf_images[pdf_idx])

            # Place the new image at the exact position and size of the old one
            slide.shapes.add_picture(
                img_io,
                old_shape.left,
                old_shape.top,
                old_shape.width,
                old_shape.height
            )

            # Remove old shape from the XML tree
            sp = old_shape._element
            sp.getparent().remove(sp)

            log.append(f"  OK | Slide {slide_idx + 1}: Image {k + 1}/{num_images} replaced with PDF page {pdf_idx + 1}")
            pdf_idx += 1

    # ── Phase 5: Save the polished presentation ──
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)

    return output.read(), log, pdf_idx


# ── Generate Section ──
if pdf_file and pptx_available:
    st.markdown('<p class="section-label">Step 2 &mdash; Validation & Generation</p>', unsafe_allow_html=True)

    # Preview metrics
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
        st.metric("Total Image Swaps", "13")

    st.markdown("")

    if st.button("Generate Monthly Report", use_container_width=True, type="primary"):
        with st.spinner("Extracting visuals and building the presentation..."):
            try:
                with open(TEMPLATE_PATH, "rb") as f:
                    template_bytes = f.read()

                out_bytes, build_logs, img_count = process_monthly_report(
                    pdf_file.read(),
                    template_bytes
                )

                st.success(f"Presentation built successfully — {img_count} slide images replaced automagically.")

                with st.expander("Build Log", expanded=False):
                    for msg in build_logs:
                        if msg.startswith("WARN"):
                            st.warning(msg)
                        else:
                            st.text(msg)

                st.markdown("")
                st.markdown('<p class="section-label">Step 3 &mdash; Download</p>', unsafe_allow_html=True)

                st.info(
                    "Please verify the summary text on Slides 4 and 8 manually "
                    "(e.g., ticket counts and ageing numbers) as these metrics require manual update.",
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

