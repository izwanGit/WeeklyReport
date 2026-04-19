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
.hub-card {
    background: #FFFFFF;
    border: 1px solid #E2E8F0;
    border-radius: 14px;
    padding: 32px 28px 28px 28px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.04), 0 4px 12px rgba(0,0,0,0.03);
    transition: all 0.25s ease;
    position: relative;
    overflow: hidden;
}
.hub-card::before {
    content: '';
    position: absolute;
    top: 0; left: 0; right: 0;
    height: 4px;
    background: linear-gradient(90deg, #00B1A9, #008C86);
}
.hub-card:hover {
    transform: translateY(-3px);
    box-shadow: 0 8px 24px rgba(0,0,0,0.08);
    border-color: #CBD5E1;
}
.hub-card-icon {
    width: 44px; height: 44px;
    border-radius: 10px;
    display: flex; align-items: center; justify-content: center;
    margin-bottom: 18px;
    background: linear-gradient(135deg, #F0FDFA, #E0F7F5);
    border: 1px solid #C6F7F3;
}
.hub-card-icon svg {
    width: 22px; height: 22px;
    stroke: #00897B;
    fill: none;
    stroke-width: 2;
    stroke-linecap: round;
    stroke-linejoin: round;
}
.hub-card-title {
    font-size: 1.15rem;
    font-weight: 700;
    color: #1E293B;
    margin-bottom: 10px;
    letter-spacing: -0.01em;
}
.hub-card-desc {
    font-size: 0.9rem;
    color: #64748B;
    line-height: 1.65;
    margin-bottom: 18px;
}
.hub-card-features {
    list-style: none;
    padding: 0; margin: 0;
}
.hub-card-features li {
    font-size: 0.82rem;
    color: #475569;
    padding: 5px 0;
    display: flex;
    align-items: center;
    gap: 8px;
}
.hub-card-features li::before {
    content: '';
    display: inline-block;
    width: 6px; height: 6px;
    background: #00B1A9;
    border-radius: 50%;
    flex-shrink: 0;
}
.hub-badge {
    display: inline-block;
    font-size: 0.65rem;
    font-weight: 700;
    text-transform: uppercase;
    letter-spacing: 0.8px;
    padding: 3px 8px;
    border-radius: 4px;
    margin-bottom: 14px;
}
.hub-badge-weekly {
    background: #E0F7F5;
    color: #00796B;
}
.hub-badge-monthly {
    background: #E8EAF6;
    color: #3949AB;
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
<img src="{_logo_sidebar_uri}" style="height:56px;" />
</div>
<p style="font-size:0.68rem; font-weight:800; color:#00B1A9; letter-spacing:1.5px; padding:0 16px; margin:0 0 8px 0;">MODULES</p>
<a href="/" target="_self" class="sidebar-nav active">Report Hub</a>
<a href="/Weekly_Report" target="_self" class="sidebar-nav">Weekly Report</a>
<a href="/Monthly_Report" target="_self" class="sidebar-nav">Monthly Report</a>
<hr class="sidebar-sep">
""", unsafe_allow_html=True)

    st.markdown("""
<div style="padding:0 16px; font-size:0.78rem; color:#94A3B8; line-height:1.5;">
PETRONAS ERP HCM Support<br>Internal Use Only
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
<div class="hub-card">
<div class="hub-card-icon">
<svg viewBox="0 0 24 24"><path d="M4 4h16c1.1 0 2 .9 2 2v12c0 1.1-.9 2-2 2H4c-1.1 0-2-.9-2-2V6c0-1.1.9-2 2-2z"/><polyline points="22,6 12,13 2,6"/></svg>
</div>
<span class="hub-badge hub-badge-weekly">Weekly</span>
<div class="hub-card-title">Email Report Generator</div>
<div class="hub-card-desc">
Transform MyGenie Excel exports into polished, production-ready HTML email reports
with automated ticket analysis and one-click Outlook delivery.
</div>
<ul class="hub-card-features">
<li>Automated ageing ticket calculations</li>
<li>Historical trend snapshots</li>
<li>Direct Outlook draft integration</li>
<li>Formatted HTML copy-to-clipboard</li>
</ul>
</div>
<div class="hub-card">
<div class="hub-card-icon">
<svg viewBox="0 0 24 24"><rect x="2" y="3" width="20" height="14" rx="2" ry="2"/><line x1="8" y1="21" x2="16" y2="21"/><line x1="12" y1="17" x2="12" y2="21"/></svg>
</div>
<span class="hub-badge hub-badge-monthly">Monthly</span>
<div class="hub-card-title">PPTX Deck Automation</div>
<div class="hub-card-desc">
Bridge your Power BI analytics dashboard directly into the corporate PowerPoint template
with automated high-fidelity image extraction and intelligent date replacement.
</div>
<ul class="hub-card-features">
<li>300 DPI lossless chart extraction</li>
<li>Automatic slide image replacement</li>
<li>Global date and text substitution</li>
<li>Zero-touch template preservation</li>
</ul>
</div>
</div>
""", unsafe_allow_html=True)


# ── Footer ──
st.markdown("""
<div class="hub-footer">
PETRONAS ERP HCM Support &mdash; Internal Use Only
</div>
""", unsafe_allow_html=True)
