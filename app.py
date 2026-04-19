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

    /* ── Hide Streamlit chrome ── */
    .stDeployButton, [data-testid="stDeployButton"], [data-testid="stAppDeployButton"] { display: none !important; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }

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

    /* ── Footer Bar ── */
    .hub-footer {
        margin-top: 48px;
        padding: 16px 0;
        border-top: 1px solid #E2E8F0;
        text-align: center;
        font-size: 0.75rem;
        color: #94A3B8;
        letter-spacing: 0.2px;
    }

    /* ── Sidebar Nav Hint ── */
    .nav-hint {
        background: #F1F5F9;
        border: 1px solid #E2E8F0;
        border-radius: 10px;
        padding: 14px 16px;
        font-size: 0.82rem;
        color: #475569;
        line-height: 1.6;
    }
    .nav-hint strong {
        color: #1E293B;
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
    <div style="text-align:center; padding: 0; margin-top: -30px; margin-bottom: 20px;">
        <img src="{_logo_sidebar_uri}" style="height: 60px;" />
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="nav-hint">
        <strong>Navigation</strong><br>
        Use the page selector above to switch between the Weekly Report and Monthly Report tools.
    </div>
    """, unsafe_allow_html=True)


# ── Header Banner ──
st.markdown(f"""
<div style="
    display: flex; align-items: center; gap: 28px;
    padding: 32px 36px;
    background: linear-gradient(135deg, #00B1A9 0%, #00897B 100%);
    border-radius: 16px;
    margin-bottom: 12px;
    box-shadow: 0 8px 32px rgba(0, 141, 134, 0.2);
    border: 1px solid rgba(255,255,255,0.12);
">
    <img src="{_logo_banner_uri}" style="height: 72px; flex-shrink: 0; filter: brightness(1.05);" />
    <div style="min-width: 0;">
        <h1 style="color:#FFFFFF; margin:0; font-weight:800; font-size:1.65rem; text-transform:uppercase; letter-spacing:0.5px; line-height:1.15;">
            MyCareerX Report Hub
        </h1>
        <p style="color:rgba(255,255,255,0.85); margin:6px 0 0 0; font-size:0.95rem; font-weight:400; line-height:1.4;">
            Centralized reporting toolkit for BAU Support weekly updates and monthly management presentations.
        </p>
    </div>
</div>
""", unsafe_allow_html=True)


# ── Card Grid ──
st.markdown("""
<div class="hub-grid">

    <!-- Weekly Report Card -->
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

    <!-- Monthly Report Card -->
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
