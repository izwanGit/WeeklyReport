import streamlit as st
import base64
import os
import sys

if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(
    page_title="PETRONAS MyCareerX Report Hub",
    page_icon="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Premium Corporate CSS ──
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

html, body, [data-testid="stAppViewContainer"] {
    font-family: 'Inter', sans-serif !important;
    background-color: #F8FAFC !important;
}
.main .block-container {
    padding-top: 1.5rem !important;
    max-width: 1100px !important;
}
[data-testid="stSidebar"] {
    border-right: 2px solid #00B1A9 !important;
    background-color: #FFFFFF !important;
}

/* ── Kill default Streamlit nav + chrome ── */
[data-testid="stSidebarNav"] { display: none !important; }
.stDeployButton, [data-testid="stDeployButton"], [data-testid="stAppDeployButton"] { display: none !important; }
#MainMenu { visibility: hidden; }
footer { visibility: hidden; }

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

/* ── Hub Card System ── */
.hub-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 28px;
    margin-top: 32px;
}
.hub-card-link-wrapper {
    text-decoration: none !important;
    display: block;
    color: inherit !important;
    height: 100%;
    outline: none;
}
.hub-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 16px;
    box-shadow: 0 4px 15px rgba(0,0,0,0.04);
    transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1);
    position: relative;
    overflow: hidden;
    display: flex;
    flex-direction: column;
    height: 100%;
    cursor: pointer;
}
.hub-card:hover {
    transform: translateY(-8px) scale(1.02);
    box-shadow: 0 20px 40px rgba(0, 177, 169, 0.15);
    border-color: #00B1A9;
}
.hub-card-banner {
    height: 110px;
    background: linear-gradient(135deg, #F0FDFA 0%, #E0F7F5 100%);
    border-bottom: 1px solid #C6F7F3;
    display: flex;
    align-items: center;
    justify-content: center;
    position: relative;
    overflow: hidden;
}
.hub-card-banner::after {
    content: '';
    position: absolute;
    width: 200%; height: 200%;
    background: radial-gradient(circle at 20% 50%, rgba(0,177,169,0.06) 0%, transparent 50%),
                radial-gradient(circle at 80% 50%, rgba(0,177,169,0.04) 0%, transparent 40%);
    top: -50%; left: -50%;
}
.hub-card-icon {
    width: 64px; height: 64px;
    background: #FFFFFF;
    border-radius: 14px;
    display: flex; align-items: center; justify-content: center;
    box-shadow: 0 8px 16px rgba(0,177,169,0.15);
    border: 2px solid #00B1A9;
    z-index: 1;
}
.hub-card-icon svg {
    width: 32px; height: 32px;
    stroke: #00B1A9;
    fill: none; stroke-width: 2; stroke-linecap: round; stroke-linejoin: round;
}
.hub-card-content {
    padding: 28px;
    flex-grow: 1;
    display: flex;
    flex-direction: column;
}
.hub-card-title {
    font-size: 1.25rem;
    font-weight: 800;
    color: #1E293B;
    margin-bottom: 12px;
    letter-spacing: -0.01em;
}
.hub-card-desc {
    font-size: 0.92rem;
    color: #64748B;
    line-height: 1.65;
    margin-bottom: 22px;
}
.hub-card-features {
    list-style: none;
    padding: 0; margin: 0 0 24px 0;
}
.hub-card-features li {
    font-size: 0.85rem;
    color: #475569;
    padding: 6px 0;
    display: flex;
    align-items: center;
    gap: 10px;
}
.hub-card-features li::before {
    content: '';
    display: inline-block;
    width: 7px; height: 7px;
    background: #00B1A9;
    border-radius: 50%;
    flex-shrink: 0;
}
.hub-card-footer {
    margin-top: auto;
    padding-top: 18px;
    border-top: 1px solid #E2E8F0;
    display: flex;
    justify-content: space-between;
    align-items: center;
    color: #00B1A9;
    font-weight: 800;
    font-size: 0.95rem;
    text-transform: uppercase;
    letter-spacing: 0.5px;
    transition: color 0.2s ease;
}
.hub-card:hover .hub-card-footer {
    color: #008C86;
    border-top-color: #CBD5E1;
}
.hub-card-footer svg {
    transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
}
.hub-card:hover .hub-card-footer svg {
    transform: translateX(6px);
}
.hub-footer {
    margin-top: 48px;
    padding: 16px 0;
    border-top: 1px solid #E2E8F0;
    text-align: center;
    font-size: 0.75rem;
    color: #94A3B8;
    letter-spacing: 0.2px;
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
<a href="/" target="_self"><img src="{_logo_sidebar_uri}" style="height:56px; transition: transform 0.2s;" onmouseover="this.style.transform='scale(1.05)'" onmouseout="this.style.transform='scale(1)'"/></a>
</div>
""", unsafe_allow_html=True)


    st.markdown("""
<div style="margin-top: 10px; padding: 18px; background-color: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 12px; box-shadow: 0 4px 6px rgba(0,0,0,0.02);">
    <div style="font-size: 0.72rem; font-weight: 800; color: #64748B; text-transform: uppercase; letter-spacing: 0.6px; margin-bottom: 16px; border-bottom: 1px solid #F1F5F9; padding-bottom: 8px;">System Status</div>
    
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
        <div style="display: flex; align-items: center; gap: 8px; font-size: 0.85rem; color: #334155; font-weight: 500;">
            <div style="width: 8px; height: 8px; border-radius: 50%; background-color: #10B981; box-shadow: 0 0 0 2px rgba(16,185,129,0.2);"></div>
            Network
        </div>
        <span style="font-size: 0.8rem; color: #64748B;">Connected</span>
    </div>

    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 12px;">
        <div style="display: flex; align-items: center; gap: 8px; font-size: 0.85rem; color: #334155; font-weight: 500;">
            <div style="width: 8px; height: 8px; border-radius: 50%; background-color: #10B981; box-shadow: 0 0 0 2px rgba(16,185,129,0.2);"></div>
            Environment
        </div>
        <span style="font-size: 0.8rem; color: #64748B;">Production</span>
    </div>

    <div style="display: flex; justify-content: space-between; align-items: center;">
        <div style="display: flex; align-items: center; gap: 8px; font-size: 0.85rem; color: #334155; font-weight: 500;">
            <div style="width: 8px; height: 8px; border-radius: 50%; background-color: #00B1A9; box-shadow: 0 0 0 2px rgba(0,177,169,0.2);"></div>
            Version
        </div>
        <span style="font-size: 0.8rem; color: #64748B; font-family: monospace;">v3.0.1</span>
    </div>
</div>

<div style="margin-top: 32px; padding:0 16px; font-size:0.75rem; color:#94A3B8; line-height:1.6; text-align: center;">
    <strong>PETRONAS ERP HCM Support</strong><br>
    Internal Use Only &copy; 2026
</div>
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
<h1 class="banner-title">MyCareerX Report Hub</h1>
<p class="banner-subtitle">Centralized reporting toolkit for BAU Support weekly updates and monthly management presentations.</p>
</div>
</div>
""", unsafe_allow_html=True)


# ── Card Grid ──
st.markdown("""
<div class="hub-grid">

<!-- Weekly Report Card -->
<a href="/Weekly_Report" target="_self" class="hub-card-link-wrapper">
<div class="hub-card">
    <div class="hub-card-banner">
        <div class="hub-card-icon">
            <svg viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
        </div>
    </div>
    <div class="hub-card-content">
        <div class="hub-card-title">Weekly Report</div>
        <div class="hub-card-desc">
            Transform MyGenie Excel exports into polished, production-ready HTML email reports with automated ticket analysis.
        </div>
        <ul class="hub-card-features">
            <li>Automated ageing ticket calculations</li>
            <li>Historical trend snapshots</li>
            <li>Direct Outlook draft integration</li>
        </ul>
        <div class="hub-card-footer">
            Launch Module
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>
        </div>
    </div>
</div>
</a>

<!-- Monthly Report Card -->
<a href="/Monthly_Report" target="_self" class="hub-card-link-wrapper">
<div class="hub-card">
    <div class="hub-card-banner">
        <div class="hub-card-icon">
            <svg viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
        </div>
    </div>
    <div class="hub-card-content">
        <div class="hub-card-title">Monthly Report</div>
        <div class="hub-card-desc">
            Bridge your Power BI analytics dashboard directly into the corporate PowerPoint template with high-fidelity extraction.
        </div>
        <ul class="hub-card-features">
            <li>300 DPI lossless chart extraction</li>
            <li>Automatic slide image replacement</li>
            <li>Zero-touch template preservation</li>
        </ul>
        <div class="hub-card-footer">
            Launch Module
            <svg width="20" height="20" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><line x1="5" y1="12" x2="19" y2="12"></line><polyline points="12 5 19 12 12 19"></polyline></svg>
        </div>
    </div>
</div>
</a>

</div>
""", unsafe_allow_html=True)


# ── Footer ──
st.markdown("""
<div class="hub-footer">
PETRONAS ERP HCM Support &mdash; Internal Use Only
</div>
""", unsafe_allow_html=True)
