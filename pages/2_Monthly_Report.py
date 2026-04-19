import streamlit as st
import io
import traceback
import sys
import os
import base64

# Conditionally import PyMuPDF and python-pptx safely
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

st.set_page_config(
    page_title="Monthly Report Generator | PETRONAS",
    page_icon="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png",
    layout="wide",
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif !important;
    }
    [data-testid="stSidebar"] { border-right: 2px solid #00B1A9 !important; }
    .stButton > button, .stDownloadButton > button {
        background: #00B1A9 !important; color: white !important;
        border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; transition: all 0.3s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #008C86 !important; transform: translateY(-1px) !important; color: white !important;
    }
    [data-testid="stFileUploader"] { 
        border: 2px dashed rgba(0, 177, 169, 0.4) !important; 
        border-radius: 12px !important; padding: 16px 20px !important; 
    }
    /* Hide Deploy button */
    .stDeployButton { display: none !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

def _image_to_data_uri(path, mime_type):
    try:
        with open(os.path.join(BASE_DIR, path), 'rb') as f:
            data = f.read()
        return f"data:{mime_type};base64,{base64.b64encode(data).decode()}"
    except:
        return ""

_logo_sidebar_uri = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg", "image/svg+xml")

with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 0; margin-top: -30px; margin-bottom: 25px;">
        <img src="{_logo_sidebar_uri}" style="height: 60px;" />
    </div>
    """, unsafe_allow_html=True)
    st.markdown("### PPTX Settings")
    old_month = st.text_input("Replace this month:", value="February 2026", help="The month name currently in your template PowerPoint.")
    new_month = st.text_input("With this month:", value="March 2026", help="The new month name for the generated report.")

st.markdown("""
<div style="display: flex; align-items: center; gap: 20px; padding: 20px 30px; background-color: #00B1A9; border-radius: 20px; margin-bottom: 2rem; box-shadow: 0 12px 35px rgba(0, 177, 169, 0.25);">
    <div style="min-width: 0;">
        <h1 style="color: white; margin: 0; font-weight: 800; font-size: 1.8rem; text-transform: uppercase;">Reporting Engine: PPTX Automation</h1>
        <p style="color: #E2E8F0; margin: 4px 0 0 0; font-size: 1.1rem;">Zero-touch bridge from Power BI Dashboard Export to PowerPoint Deck.</p>
    </div>
</div>
""", unsafe_allow_html=True)

if not PPTX_AVAILABLE:
    st.error("Missing python-pptx or PyMuPDF. Please re-run dependencies installation.")
    st.stop()

st.markdown("### 1. Upload Assets")
c1, c2 = st.columns(2)
with c1:
    pdf_file = st.file_uploader("Upload Power BI PDF Export (13 Pages)", type=['pdf'])
with c2:
    pptx_file = st.file_uploader("Upload PPTX Template", type=['pptx'])

def process_monthly_report(pdf_bytes, pptx_bytes, old_text, new_text):
    # 1. Extract Images
    pdf = fitz.open(stream=pdf_bytes, filetype="pdf")
    pdf_images = []
    
    # We enforce high quality rendering for the PBI charts
    for page_num in range(len(pdf)):
        page = pdf.load_page(page_num)
        # 300 DPI equivalent for extremely crisp text/tables in PPT
        pix = page.get_pixmap(matrix=fitz.Matrix(4, 4))
        pdf_images.append(pix.tobytes("png"))
        
    # 2. Open PPTX
    prs = Presentation(io.BytesIO(pptx_bytes))
    
    # 3. Global Text Replacement
    # Replaces 'February 2026' with 'March 2026' everywhere
    for slide in prs.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                for paragraph in shape.text_frame.paragraphs:
                    for run in paragraph.runs:
                        if old_text in run.text:
                            run.text = run.text.replace(old_text, new_text)
            if shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        for paragraph in cell.text_frame.paragraphs:
                            for run in paragraph.runs:
                                if old_text in run.text:
                                    run.text = run.text.replace(old_text, new_text)

    # 4. Image/Placeholder Replacement (The Mapping Engine)
    # Target slides: 3, 4, 5, 6, 7, 8, 9, 10
    # Which corresponds to array indices: 2, 3, 4, 5, 6, 7, 8, 9
    mapping = {
        2: 2, # Slide 3 -> 2 images
        3: 1, # Slide 4 -> 1 image
        4: 1, # Slide 5 -> 1 image (Table screenshot)
        5: 2, # Slide 6 -> 2 images
        6: 3, # Slide 7 -> 3 images
        7: 1, # Slide 8 -> 1 image
        8: 1, # Slide 9 -> 1 image (Table screenshot)
        9: 2  # Slide 10 -> 2 images
    }
    
    pdf_idx = 0
    log = []
    
    for slide_idx, num_images in mapping.items():
        if slide_idx >= len(prs.slides):
            log.append(f"⚠️ Slide {slide_idx+1} does not exist in template. Skipping.")
            continue
            
        slide = prs.slides[slide_idx]
        
        # Collect all picture shapes and picture placeholders
        target_shapes = []
        for shape in slide.shapes:
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                target_shapes.append(shape)
            elif getattr(shape, 'is_placeholder', False) and shape.placeholder_format.type == 18:
                target_shapes.append(shape)
                
        # Sort them Top-to-Bottom, Left-to-Right for deterministic replacement
        target_shapes.sort(key=lambda s: (s.top, s.left))
        
        for k in range(num_images):
            if k >= len(target_shapes):
                log.append(f"⚠️ Slide {slide_idx+1}: Expected {num_images} images but only found {len(target_shapes)}.")
                break
            if pdf_idx >= len(pdf_images):
                log.append(f"⚠️ Ran out of PDF pages! Stopped at page {pdf_idx}.")
                break
                
            old_shape = target_shapes[k]
            img_io = io.BytesIO(pdf_images[pdf_idx])
            
            # Insert the new HD Image identically where the old one was
            slide.shapes.add_picture(
                img_io,
                old_shape.left,
                old_shape.top,
                old_shape.width,
                old_shape.height
            )
            
            # Delete old shape using XML removal to clean up
            sp = old_shape._element
            sp.getparent().remove(sp)
            
            log.append(f"✅ Slide {slide_idx+1}: Replaced image/placeholder with PowerBI PDF Page {pdf_idx+1}")
            pdf_idx += 1
            
    # Save the polished presentation
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output.read(), log

if pdf_file and pptx_file:
    st.markdown("### 2. Generate Deck")
    if st.button("🚀 Process Power BI and Build PPTX", use_container_width=True, type="primary"):
        with st.spinner("Extracting High-Res Graphics and Automating Deck..."):
            try:
                out_bytes, build_logs = process_monthly_report(pdf_file.read(), pptx_file.read(), old_month, new_month)
                
                st.success("✅ Presentation successfully built!")
                
                # Show logs in expander
                with st.expander("Show Build Logs"):
                    for msg in build_logs:
                        st.text(msg)
                
                st.warning(f"**Reminder**: All dates matching `{old_month}` were successfully replaced with `{new_month}`. However, you may need to manually verify or update numbered metrics (e.g., 'Ticket logged increased by 41') in Slides 4 and 8.", icon="ℹ️")
                
                st.download_button(
                    label="⬇️ Download Final Monthly Report (PPTX)",
                    data=out_bytes,
                    file_name=f"Monthly_Report_{new_month.replace(' ', '_')}.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
            except Exception as e:
                st.error("There was an error generating the document.")
                st.code(traceback.format_exc())
else:
    st.info("👆 Upload both the Power BI PDF file and the PPTX template to proceed.")
