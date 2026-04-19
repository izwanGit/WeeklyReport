import streamlit as st
import base64
import os
import sys

# Conditional import for PyInstaller path resolution
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(
    page_title="PETRONAS Report Hub",
    page_icon="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Premium CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif !important;
        background-color: #F8FAFC !important;
    }
    .main .block-container { padding-top: 2rem !important; max-width: 1200px !important; }
    [data-testid="stSidebar"] { border-right: 2px solid #00B1A9 !important; background-color: #FFFFFF !important; }
    
    .hub-card {
        background: white;
        border: 1px solid #E2E8F0;
        border-radius: 16px;
        padding: 30px;
        margin-bottom: 24px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.05), 0 2px 4px -1px rgba(0, 0, 0, 0.03);
        transition: all 0.3s ease;
        border-left: 6px solid #00B1A9;
    }
    .hub-card:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1), 0 4px 6px -2px rgba(0, 0, 0, 0.05);
        border-color: #00B1A9;
    }
    .hub-title {
        color: #1E293B;
        font-size: 1.5rem;
        font-weight: 700;
        margin-bottom: 12px;
        display: flex;
        align-items: center;
        gap: 12px;
    }
    .hub-desc {
        color: #475569;
        font-size: 1rem;
        line-height: 1.6;
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

_logo_banner_uri = _image_to_data_uri("PETRONAS_LOGO_SQUARE.png", "image/png")
_logo_sidebar_uri = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg", "image/svg+xml")

with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 0; margin-top: -30px; margin-bottom: 15px;">
        <img src="{_logo_sidebar_uri}" style="height: 60px;" />
    </div>
    """, unsafe_allow_html=True)
    st.info("👈 Select a tool from the menu above to get started.")

st.markdown(f"""
<style>
    .banner-title {{ color: #FFFFFF !important; text-transform: uppercase !important; font-weight: 800 !important; text-shadow: 0px 2px 4px rgba(0,0,0,0.3) !important; margin: 0 !important; line-height: 1.1 !important; white-space: nowrap; font-size: clamp(1.5rem, 4vw, 2.2rem) !important; letter-spacing: 0.1px; }}
    .banner-subtitle {{ color: #E2E8F0 !important; font-weight: 500 !important; text-shadow: 0px 1px 3px rgba(0,0,0,0.2) !important; margin: 8px 0 0 0 !important; white-space: nowrap; font-size: clamp(1rem, 2vw, 1.2rem) !important; }}
</style>
<div style="display: flex; align-items: center; gap: 30px; padding: 40px 40px; background: linear-gradient(135deg, #00B1A9 0%, #008C86 100%); border-radius: 24px; margin-bottom: 3rem; box-shadow: 0 20px 40px rgba(0, 177, 169, 0.25); overflow: hidden; border: 1px solid rgba(255, 255, 255, 0.2);">
    <img src="{_logo_banner_uri}" style="height: 100px; flex-shrink: 0; filter: drop-shadow(0px 4px 8px rgba(0,0,0,0.2));" />
    <div style="min-width: 0;">
        <h1 class="banner-title">MyCareerX Report Hub</h1>
        <p class="banner-subtitle">Automated reporting toolkit for Weekly BAU & Monthly Management updates</p>
    </div>
</div>
""", unsafe_allow_html=True)

col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="hub-card">
        <div class="hub-title">📊 Weekly Report Generator</div>
        <div class="hub-desc">
            Convert standard MyGenie Excel exports into a highly polished, production-ready HTML email draft for your weekly management updates.
            <br><br>
            <b>Features:</b>
            <ul style="margin-top: 8px; margin-bottom: 0;">
                <li>Automated ticketing status calculation</li>
                <li>One-click Outlook integration</li>
                <li>Historic snapshot saving</li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)

with col2:
    st.markdown("""
    <div class="hub-card" style="border-left-color: #008C86;">
        <div class="hub-title">📋 Monthly PPTX Automation</div>
        <div class="hub-desc">
            Directly bridge the gap between your latest Power BI Analytics Dashboard and the corporate PowerPoint presentation. 
            <br><br>
            <b>Features:</b>
            <ul style="margin-top: 8px; margin-bottom: 0;">
                <li>Lossless high-resolution image extraction from Power BI PDFs</li>
                <li>Automatic corporate template integration</li>
                <li>Intelligent date & text replacement</li>
            </ul>
        </div>
    </div>
    """, unsafe_allow_html=True)
