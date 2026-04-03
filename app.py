import streamlit as st
import pandas as pd
import json
import os
import datetime
from jinja2 import Environment, FileSystemLoader
import sys
import tempfile
import base64

# Conditional import for win32com
try:
    if sys.platform == 'win32':
        import win32com.client as win32
    else:
        win32 = None
except ImportError:
    win32 = None

# ----------------------------------------------------
# Configuration & Constants
# ----------------------------------------------------
if getattr(sys, 'frozen', False):
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

HISTORY_FILE = "history.json"
TEMPLATE_FILE = os.path.join(BASE_DIR, "template.html")

# Expected Columns
COL_AGEING = "SR Ageing"
COL_WO_NO = "Work Order No."
COL_SUMMARY = "Summary"
COL_USER_TSG = "User/TSG"
COL_STATUS = "WO Status"
COL_STATUS_REASON = "WO Status Reason"
COL_ASSIGNEE = "Assignee"

# ----------------------------------------------------
# Helper Functions
# ----------------------------------------------------
def get_column_by_prefix(df, prefix):
    """Safely find a column that starts with or contains the prefix, ignoring case"""
    for col in df.columns:
        if prefix.lower() in str(col).lower():
            return col
    return None

def process_tickets(df):
    """Normalize and clean the dataframe columns based on expectations"""
    ageing_col = get_column_by_prefix(df, "ageing")
    if ageing_col is None:
        st.warning("Could not find 'Ageing' column in the uploaded sheet.")
        return None, None
    
    df[ageing_col] = pd.to_numeric(df[ageing_col], errors='coerce').fillna(0)
    norm_df = df.copy()
    
    def map_col(prefix, default_name):
        col = get_column_by_prefix(df, prefix)
        if col:
            norm_df[default_name] = df[col]
        else:
            norm_df[default_name] = ""
            
    map_col("work order", COL_WO_NO)
    map_col("summary", COL_SUMMARY)
    map_col("user/tsg", COL_USER_TSG)
    map_col("status", COL_STATUS)
    
    for col in df.columns:
        if "reason" in str(col).lower() and "status" in str(col).lower():
            norm_df[COL_STATUS_REASON] = df[col]
        elif "status" in str(col).lower() and "reason" not in str(col).lower():
            norm_df[COL_STATUS] = df[col]
            
    map_col("assignee", COL_ASSIGNEE)
    norm_df[COL_AGEING] = norm_df[ageing_col]
    
    return norm_df, norm_df[COL_AGEING]

def load_history():
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, 'r') as f:
                return json.load(f)
        except Exception as e:
            st.error(f"Error reading history file: {e}")
            return []
    return []

def save_history(history):
    try:
        with open(HISTORY_FILE, 'w') as f:
            json.dump(history, f, indent=4)
        return True
    except Exception as e:
        st.error(f"Error saving history file: {e}")
        return False

def push_to_outlook(html_body, subject="Weekly SR & Incident Update"):
    if sys.platform != 'win32' or win32 is None:
        st.error("Outlook integration is only supported on Windows machines with pywin32 installed.")
        return False
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Display(True)
        return True
    except Exception as e:
        st.error(f"Failed to open Outlook draft: {str(e)}")
        return False

# ----------------------------------------------------
# Page Config & Premium Styling
# ----------------------------------------------------
st.set_page_config(
    page_title="PETRONAS Weekly Report Generator",
    page_icon="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ============================================================
# MEGA PREMIUM PETRONAS CSS
# ============================================================
st.markdown("""
<style>
    /* ========== GOOGLE FONTS ========== */
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');

    /* ========== ROOT VARIABLES ========== */
    :root {
        --petronas-green: #00A19C;
        --petronas-dark: #007A77;
        --petronas-light: #E6F7F6;
        --petronas-glow: rgba(0, 161, 156, 0.15);
        --bg-dark: #0A0F1C;
        --bg-card: #111827;
        --bg-card-hover: #1A2332;
        --text-primary: #F1F5F9;
        --text-secondary: #94A3B8;
        --text-muted: #64748B;
        --border-color: rgba(0, 161, 156, 0.2);
        --shadow-glow: 0 0 20px rgba(0, 161, 156, 0.15);
    }

    /* ========== GLOBAL RESET ========== */
    html, body, [data-testid="stAppViewContainer"] {
        background-color: var(--bg-dark) !important;
        color: var(--text-primary) !important;
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif !important;
    }

    .main .block-container {
        padding: 2rem 3rem !important;
        max-width: 1400px !important;
    }

    /* ========== HEADER / TOP BAR ========== */
    header[data-testid="stHeader"] {
        background: linear-gradient(135deg, #0A0F1C 0%, #111827 100%) !important;
        border-bottom: 1px solid var(--border-color) !important;
    }

    /* ========== SIDEBAR - PETRONAS DARK GLASS ========== */
    [data-testid="stSidebar"] {
        background: linear-gradient(180deg, #0D1321 0%, #111827 40%, #0A0F1C 100%) !important;
        border-right: 1px solid var(--border-color) !important;
    }

    [data-testid="stSidebar"] * {
        color: var(--text-primary) !important;
        font-family: 'Inter', sans-serif !important;
    }

    [data-testid="stSidebar"] .stMarkdown p {
        color: var(--text-secondary) !important;
    }

    [data-testid="stSidebar"] hr {
        border-color: var(--border-color) !important;
    }

    /* Sidebar header styling */
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 {
        color: var(--petronas-green) !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
    }

    /* ========== HEADINGS ========== */
    h1 {
        color: var(--text-primary) !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 800 !important;
        letter-spacing: -0.03em !important;
        font-size: 2rem !important;
    }

    h2, h3 {
        color: var(--text-primary) !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 700 !important;
        letter-spacing: -0.02em !important;
    }

    /* ========== BUTTONS ========== */
    .stButton > button {
        background: linear-gradient(135deg, var(--petronas-green) 0%, var(--petronas-dark) 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-family: 'Inter', sans-serif !important;
        padding: 0.65rem 1.5rem !important;
        font-size: 0.9rem !important;
        letter-spacing: 0.02em !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
        box-shadow: 0 4px 14px rgba(0, 161, 156, 0.3) !important;
    }

    .stButton > button:hover {
        background: linear-gradient(135deg, #00BFB8 0%, var(--petronas-green) 100%) !important;
        box-shadow: 0 6px 24px rgba(0, 161, 156, 0.45) !important;
        transform: translateY(-2px) !important;
        color: white !important;
    }

    .stButton > button:active {
        transform: translateY(0px) !important;
    }

    /* ========== DOWNLOAD BUTTON ========== */
    .stDownloadButton > button {
        background: linear-gradient(135deg, var(--petronas-green) 0%, var(--petronas-dark) 100%) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        font-weight: 600 !important;
        font-family: 'Inter', sans-serif !important;
        padding: 0.65rem 1.5rem !important;
        box-shadow: 0 4px 14px rgba(0, 161, 156, 0.3) !important;
        transition: all 0.3s cubic-bezier(0.4, 0, 0.2, 1) !important;
    }

    .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #00BFB8 0%, var(--petronas-green) 100%) !important;
        box-shadow: 0 6px 24px rgba(0, 161, 156, 0.45) !important;
        transform: translateY(-2px) !important;
        color: white !important;
    }

    /* ========== FILE UPLOADER ========== */
    [data-testid="stFileUploader"] {
        background: var(--bg-card) !important;
        border: 2px dashed var(--border-color) !important;
        border-radius: 12px !important;
        padding: 1rem !important;
        transition: all 0.3s ease !important;
    }

    [data-testid="stFileUploader"]:hover {
        border-color: var(--petronas-green) !important;
        box-shadow: var(--shadow-glow) !important;
    }

    [data-testid="stFileUploader"] * {
        color: var(--text-secondary) !important;
    }

    [data-testid="stFileUploader"] button {
        background: var(--petronas-green) !important;
        color: white !important;
        border-radius: 8px !important;
    }

    /* ========== SELECT BOX / INPUT FIELDS ========== */
    [data-testid="stSelectbox"] > div > div,
    .stDateInput > div > div > input,
    .stTextInput > div > div > input,
    .stTextArea > div > div > textarea {
        background-color: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 10px !important;
        font-family: 'Inter', sans-serif !important;
    }

    [data-testid="stSelectbox"] label,
    .stDateInput label,
    .stTextInput label,
    .stTextArea label {
        color: var(--text-secondary) !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 500 !important;
    }

    /* ========== METRICS ========== */
    [data-testid="stMetric"] {
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 14px !important;
        padding: 1.2rem !important;
        box-shadow: var(--shadow-glow) !important;
        transition: all 0.3s ease !important;
    }

    [data-testid="stMetric"]:hover {
        border-color: var(--petronas-green) !important;
        box-shadow: 0 0 30px rgba(0, 161, 156, 0.25) !important;
        transform: translateY(-2px) !important;
    }

    [data-testid="stMetricLabel"] {
        color: var(--text-secondary) !important;
        font-weight: 500 !important;
        font-size: 0.85rem !important;
    }

    [data-testid="stMetricValue"] {
        color: var(--petronas-green) !important;
        font-weight: 800 !important;
        font-size: 2rem !important;
    }

    [data-testid="stMetricDelta"] {
        font-weight: 600 !important;
    }

    /* ========== TABS ========== */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: transparent !important;
    }

    .stTabs [data-baseweb="tab"] {
        background: var(--bg-card) !important;
        color: var(--text-secondary) !important;
        border-radius: 10px !important;
        border: 1px solid var(--border-color) !important;
        padding: 0.5rem 1.25rem !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
        transition: all 0.3s ease !important;
    }

    .stTabs [aria-selected="true"] {
        background: linear-gradient(135deg, var(--petronas-green) 0%, var(--petronas-dark) 100%) !important;
        color: white !important;
        border-color: var(--petronas-green) !important;
        box-shadow: 0 4px 14px rgba(0, 161, 156, 0.3) !important;
    }

    .stTabs [data-baseweb="tab-highlight"] {
        display: none !important;
    }

    .stTabs [data-baseweb="tab-border"] {
        display: none !important;
    }

    /* ========== EXPANDER ========== */
    .streamlit-expanderHeader {
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 10px !important;
        color: var(--text-primary) !important;
        font-family: 'Inter', sans-serif !important;
        font-weight: 600 !important;
    }

    .streamlit-expanderContent {
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color) !important;
        border-top: none !important;
        border-radius: 0 0 10px 10px !important;
    }

    /* ========== ALERTS / INFO / SUCCESS / WARNING / ERROR ========== */
    [data-testid="stAlert"] {
        border-radius: 12px !important;
        font-family: 'Inter', sans-serif !important;
    }

    .stAlert [data-testid="stMarkdownContainer"] {
        font-family: 'Inter', sans-serif !important;
    }

    /* ========== HORIZONTAL RULE ========== */
    hr {
        border-color: var(--border-color) !important;
    }

    /* ========== MARKDOWN ========== */
    .stMarkdown p, .stMarkdown li {
        color: var(--text-secondary) !important;
        font-family: 'Inter', sans-serif !important;
    }

    /* ========== IFRAME (Email preview) ========== */
    iframe {
        border-radius: 12px !important;
        border: 1px solid var(--border-color) !important;
        box-shadow: var(--shadow-glow) !important;
    }

    /* ========== TEXT AREA (Raw HTML) ========== */
    .stTextArea textarea {
        background-color: #0D1117 !important;
        color: #C9D1D9 !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 10px !important;
        font-family: 'JetBrains Mono', 'Fira Code', 'Cascadia Code', monospace !important;
        font-size: 0.8rem !important;
    }

    /* ========== SCROLLBAR ========== */
    ::-webkit-scrollbar {
        width: 6px;
        height: 6px;
    }

    ::-webkit-scrollbar-track {
        background: var(--bg-dark);
    }

    ::-webkit-scrollbar-thumb {
        background: var(--petronas-green);
        border-radius: 10px;
    }

    ::-webkit-scrollbar-thumb:hover {
        background: #00BFB8;
    }

    /* ========== TOOLTIP / POPOVER ========== */
    [data-baseweb="popover"] {
        background: var(--bg-card) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 10px !important;
    }

    /* ========== CUSTOM PETRONAS GLOW ANIMATION ========== */
    @keyframes petronasGlow {
        0% { box-shadow: 0 0 5px rgba(0, 161, 156, 0.2); }
        50% { box-shadow: 0 0 20px rgba(0, 161, 156, 0.4); }
        100% { box-shadow: 0 0 5px rgba(0, 161, 156, 0.2); }
    }

    /* ========== HIDE STREAMLIT BRANDING ========== */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: visible;}
</style>
""", unsafe_allow_html=True)

# ============================================================
# HEADER WITH PETRONAS LOGO
# ============================================================
st.markdown("""
<div style="
    display: flex; 
    align-items: center; 
    gap: 24px; 
    padding: 20px 28px; 
    background: linear-gradient(135deg, #111827 0%, #1A2332 100%); 
    border: 1px solid rgba(0, 161, 156, 0.2); 
    border-radius: 16px; 
    margin-bottom: 1.5rem;
    box-shadow: 0 0 30px rgba(0, 161, 156, 0.08);
">
    <img src="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png" 
         alt="PETRONAS" 
         style="height: 55px; filter: brightness(1.1);" />
    <div>
        <h1 style="margin: 0 !important; padding: 0 !important; font-size: 1.75rem !important; 
                    background: linear-gradient(135deg, #FFFFFF 0%, #00A19C 100%); 
                    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
                    font-weight: 800 !important; letter-spacing: -0.03em !important;">
            Weekly SR & Incident Report Generator
        </h1>
        <p style="margin: 4px 0 0 0 !important; color: #94A3B8 !important; font-size: 0.95rem; font-weight: 400;">
            Automate your MyGenie Excel exports into production-ready HTML email reports.
        </p>
    </div>
</div>
""", unsafe_allow_html=True)


# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    # Sidebar logo + branding
    st.markdown("""
    <div style="text-align:center; margin-bottom: 20px; padding: 15px 0;">
        <img src="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png" 
             style="height: 40px; margin-bottom: 10px; filter: brightness(1.1);" />
        <div style="height: 3px; background: linear-gradient(90deg, transparent, #00A19C, transparent); margin: 10px auto; width: 60%;"></div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📋 Report Configuration")
    st.markdown("")

    report_date = st.date_input("📅 Report Date", datetime.date.today())
    report_date_str = report_date.strftime("%d %B %Y")

    st.markdown("")
    st.markdown("---")
    st.markdown("")

    st.markdown("### 📂 Upload Data")
    uploaded_file = st.file_uploader("Upload MyGenie Excel (.xlsx)", type=['xlsx', 'xls'])

    sr_sheet = None
    inc_sheet = None

    if uploaded_file:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = xl.sheet_names
            st.markdown("")
            st.markdown("##### 📑 Sheet Mapping")
            sr_sheet = st.selectbox("SR Tickets Sheet", sheet_names, index=0)
            inc_sheet = st.selectbox("INC Tickets Sheet", sheet_names, index=min(1, len(sheet_names)-1))
        except Exception as e:
            st.error(f"❌ Error reading Excel: {e}")
            uploaded_file = None

    # Sidebar footer
    st.markdown("---")
    st.markdown("""
    <div style="text-align:center; padding: 10px 0;">
        <p style="font-size: 0.7rem; color: #64748B !important; margin: 0;">
            PETRONAS Weekly Report Tool v1.0<br/>
            © 2026 PETRONAS. Internal Use Only.
        </p>
    </div>
    """, unsafe_allow_html=True)


# ============================================================
# MAIN CONTENT
# ============================================================
if uploaded_file and sr_sheet and inc_sheet:

    try:
        # Load sheets
        df_sr = pd.read_excel(uploaded_file, sheet_name=sr_sheet)
        df_inc = pd.read_excel(uploaded_file, sheet_name=inc_sheet)

        # --- SR Processing ---
        norm_sr, sr_ageing = process_tickets(df_sr)
        if norm_sr is not None:
            sr_total = len(norm_sr)
            sr_gt_30 = norm_sr[norm_sr[COL_AGEING] > 30]
            sr_gt_30_count = len(sr_gt_30)
            sr_15_30 = norm_sr[(norm_sr[COL_AGEING] >= 15) & (norm_sr[COL_AGEING] <= 30)]
            sr_15_30_count = len(sr_15_30)
            sr_1_14 = norm_sr[(norm_sr[COL_AGEING] >= 1) & (norm_sr[COL_AGEING] <= 14)]
            sr_1_14_count = len(sr_1_14)
            sr_gt_1_count = len(norm_sr[norm_sr[COL_AGEING] > 1])
            sr_gt_1_pct = round((sr_gt_1_count / sr_total * 100) if sr_total > 0 else 0, 2)
            sr_gt_30_pct = round((sr_gt_30_count / sr_total * 100) if sr_total > 0 else 0, 2)
            sr_gt_30_dict = sr_gt_30[[COL_AGEING, COL_WO_NO, COL_SUMMARY, COL_USER_TSG, COL_STATUS, COL_STATUS_REASON, COL_ASSIGNEE]].to_dict('records')
            sr_15_30_dict = sr_15_30[[COL_AGEING, COL_WO_NO, COL_SUMMARY, COL_USER_TSG, COL_STATUS, COL_STATUS_REASON, COL_ASSIGNEE]].to_dict('records')

        # --- INC Processing ---
        norm_inc, inc_ageing = process_tickets(df_inc)
        if norm_inc is not None:
            inc_total = len(norm_inc)
            inc_gt_90_count = len(norm_inc[norm_inc[COL_AGEING] > 90])
            inc_61_90_count = len(norm_inc[(norm_inc[COL_AGEING] >= 61) & (norm_inc[COL_AGEING] <= 90)])
            inc_31_60_count = len(norm_inc[(norm_inc[COL_AGEING] >= 31) & (norm_inc[COL_AGEING] <= 60)])
            inc_15_30_count = len(norm_inc[(norm_inc[COL_AGEING] >= 15) & (norm_inc[COL_AGEING] <= 30)])
            inc_8_14_count  = len(norm_inc[(norm_inc[COL_AGEING] >= 8) & (norm_inc[COL_AGEING] <= 14)])
            inc_3_7_count   = len(norm_inc[(norm_inc[COL_AGEING] >= 3) & (norm_inc[COL_AGEING] <= 7)])
            inc_gt_1_count = len(norm_inc[norm_inc[COL_AGEING] > 1])
            inc_gt_1_pct = round((inc_gt_1_count / inc_total * 100) if inc_total > 0 else 0, 2)

        # --- History Tracking (4 weeks) ---
        history = load_history()
        short_date = report_date.strftime("%d-%b-%Y")

        if len(history) == 0 or history[-1].get("date") != short_date:
            new_record = {
                "date": short_date,
                "sr_count_gt_30": sr_gt_30_count,
                "sr_count_15_30": sr_15_30_count,
                "sr_count_1_14": sr_1_14_count,
                "inc_count_gt_90": inc_gt_90_count,
                "inc_count_61_90": inc_61_90_count,
                "inc_count_31_60": inc_31_60_count,
                "inc_count_15_30": inc_15_30_count,
                "inc_count_8_14": inc_8_14_count,
                "inc_count_3_7": inc_3_7_count
            }
            history.append(new_record)
            if len(history) > 4:
                history = history[-4:]
            save_history(history)

        # Prepare arrays for template
        trend_dates = [h["date"] for h in history]
        sr_trend_gt_30 = [h["sr_count_gt_30"] for h in history]
        sr_trend_15_30 = [h["sr_count_15_30"] for h in history]
        sr_trend_1_14 = [h["sr_count_1_14"] for h in history]
        inc_trend_gt_90 = [h["inc_count_gt_90"] for h in history]
        inc_trend_61_90 = [h["inc_count_61_90"] for h in history]
        inc_trend_31_60 = [h["inc_count_31_60"] for h in history]
        inc_trend_15_30 = [h["inc_count_15_30"] for h in history]
        inc_trend_8_14  = [h["inc_count_8_14"] for h in history]
        inc_trend_3_7   = [h["inc_count_3_7"] for h in history]

        # ========== KPI METRICS CARDS ==========
        st.markdown("### 📊 Key Metrics Overview")
        st.markdown(f"<p style='color:#64748B; margin-top:-10px;'>Report date: <b style='color:#00A19C;'>{report_date_str}</b></p>", unsafe_allow_html=True)

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            st.metric("Total SR Tickets", sr_total)
        with m2:
            st.metric("SR Ageing > 30d", sr_gt_30_count)
        with m3:
            st.metric("SR > 1 day %", f"{sr_gt_1_pct}%")
        with m4:
            st.metric("Total INC Tickets", inc_total)
        with m5:
            st.metric("INC > 1 day %", f"{inc_gt_1_pct}%")

        st.markdown("")

        # ========== GENERATE HTML ==========
        env = Environment(loader=FileSystemLoader(BASE_DIR))
        try:
            template = env.get_template("template.html")
            html_output = template.render(
                report_date=report_date_str,
                sr_total=sr_total,
                sr_ageing_more_than_1_day_pct=sr_gt_1_pct,
                sr_ageing_more_than_30_days_pct=sr_gt_30_pct,
                sr_ageing_gt_30_tickets=sr_gt_30_dict,
                sr_ageing_15_30_tickets=sr_15_30_dict,
                inc_total=inc_total,
                inc_ageing_more_than_1_day_pct=inc_gt_1_pct,
                trend_dates=trend_dates,
                sr_trend_gt_30=sr_trend_gt_30,
                sr_trend_15_30=sr_trend_15_30,
                sr_trend_1_14=sr_trend_1_14,
                inc_trend_gt_90=inc_trend_gt_90,
                inc_trend_61_90=inc_trend_61_90,
                inc_trend_31_60=inc_trend_31_60,
                inc_trend_15_30=inc_trend_15_30,
                inc_trend_8_14=inc_trend_8_14,
                inc_trend_3_7=inc_trend_3_7
            )

            # ========== TABBED INTERFACE ==========
            tab_preview, tab_source, tab_export = st.tabs(["📧 Email Preview", "💻 HTML Source", "🚀 Export & Actions"])

            with tab_preview:
                st.markdown("""
                <div style="
                    background: #111827; 
                    border: 1px solid rgba(0,161,156,0.2); 
                    border-radius: 16px; 
                    padding: 8px; 
                    margin-top: 10px;
                    box-shadow: 0 0 30px rgba(0,161,156,0.08);
                ">
                """, unsafe_allow_html=True)
                st.components.v1.html(html_output, height=900, scrolling=True)
                st.markdown("</div>", unsafe_allow_html=True)

            with tab_source:
                st.markdown("##### Raw HTML — Copy for manual paste into Outlook")
                st.code(html_output, language="html")

            with tab_export:
                st.markdown("### 🚀 Export Options")
                st.markdown("")

                exp1, exp2, exp3 = st.columns(3)

                with exp1:
                    st.markdown("""
                    <div style="
                        background: #111827; 
                        border: 1px solid rgba(0,161,156,0.2); 
                        border-radius: 14px; 
                        padding: 24px; 
                        text-align: center;
                        box-shadow: 0 0 20px rgba(0,161,156,0.08);
                    ">
                        <p style="font-size: 2rem; margin: 0;">💾</p>
                        <p style="color: #F1F5F9 !important; font-weight: 700; margin: 8px 0 4px 0;">Download HTML</p>
                        <p style="color: #64748B !important; font-size: 0.8rem; margin-bottom: 16px;">Save as .html file</p>
                    </div>
                    """, unsafe_allow_html=True)
                    st.download_button(
                        label="Download .html",
                        data=html_output,
                        file_name=f"Weekly_Report_{report_date.strftime('%Y%m%d')}.html",
                        mime="text/html",
                        use_container_width=True
                    )

                with exp2:
                    st.markdown("""
                    <div style="
                        background: #111827; 
                        border: 1px solid rgba(0,161,156,0.2); 
                        border-radius: 14px; 
                        padding: 24px; 
                        text-align: center;
                        box-shadow: 0 0 20px rgba(0,161,156,0.08);
                    ">
                        <p style="font-size: 2rem; margin: 0;">📋</p>
                        <p style="color: #F1F5F9 !important; font-weight: 700; margin: 8px 0 4px 0;">Copy to Clipboard</p>
                        <p style="color: #64748B !important; font-size: 0.8rem; margin-bottom: 16px;">Copy HTML source code</p>
                    </div>
                    """, unsafe_allow_html=True)
                    if st.button("Copy HTML Source", use_container_width=True):
                        st.code(html_output[:200] + "...", language="html")
                        st.info("💡 Use the HTML Source tab to copy the full source.")

                with exp3:
                    st.markdown("""
                    <div style="
                        background: #111827; 
                        border: 1px solid rgba(0,161,156,0.2); 
                        border-radius: 14px; 
                        padding: 24px; 
                        text-align: center;
                        box-shadow: 0 0 20px rgba(0,161,156,0.08);
                    ">
                        <p style="font-size: 2rem; margin: 0;">✉️</p>
                        <p style="color: #F1F5F9 !important; font-weight: 700; margin: 8px 0 4px 0;">Outlook Draft</p>
                        <p style="color: #64748B !important; font-size: 0.8rem; margin-bottom: 16px;">Push directly to Outlook</p>
                    </div>
                    """, unsafe_allow_html=True)
                    if sys.platform == 'win32':
                        if st.button("Push to Outlook Draft", use_container_width=True):
                            success = push_to_outlook(html_output, f"Weekly SR & Incident Report - {report_date_str}")
                            if success:
                                st.success("✅ Draft created in Outlook!")
                    else:
                        st.button("Outlook (Windows Only)", use_container_width=True, disabled=True)
                        st.caption("⚠️ Only available on Windows with Outlook installed.")

        except Exception as e:
            st.error(f"❌ Error rendering template: {e}")

    except Exception as e:
        st.error(f"❌ An unexpected error occurred: {e}")

else:
    # ========== EMPTY STATE ==========
    st.markdown("")
    st.markdown("")

    st.markdown("""
    <div style="
        text-align: center; 
        padding: 80px 40px; 
        background: linear-gradient(135deg, #111827 0%, #1A2332 100%); 
        border: 1px dashed rgba(0, 161, 156, 0.3); 
        border-radius: 20px;
        box-shadow: 0 0 40px rgba(0, 161, 156, 0.05);
    ">
        <img src="https://upload.wikimedia.org/wikipedia/commons/2/22/PETRONAS_Logo_%28for_solid_white_background%29.png" 
             style="height: 50px; margin-bottom: 25px; opacity: 0.7; filter: brightness(1.1);" />
        <h2 style="color: #F1F5F9 !important; font-size: 1.5rem !important; font-weight: 700 !important; margin: 0 0 10px 0 !important;">
            Ready to Generate Your Weekly Report
        </h2>
        <p style="color: #64748B !important; font-size: 1rem; max-width: 500px; margin: 0 auto 25px auto; line-height: 1.6;">
            Upload your MyGenie Excel export using the sidebar to automatically generate the formatted HTML email report.
        </p>
        <div style="display: inline-flex; gap: 8px; align-items: center; 
                    background: rgba(0,161,156,0.1); padding: 10px 20px; border-radius: 10px; 
                    border: 1px solid rgba(0,161,156,0.2);">
            <span style="font-size: 1.1rem;">👈</span>
            <span style="color: #00A19C; font-weight: 600; font-size: 0.9rem;">Use the sidebar to get started</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("")
    st.markdown("")

    # Feature highlight cards
    f1, f2, f3 = st.columns(3)
    with f1:
        st.markdown("""
        <div style="
            background: #111827; 
            border: 1px solid rgba(0,161,156,0.15); 
            border-radius: 14px; 
            padding: 28px; 
            text-align: center;
            transition: all 0.3s ease;
            height: 200px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        ">
            <p style="font-size: 2.2rem; margin: 0 0 12px 0;">⚡</p>
            <p style="color: #F1F5F9 !important; font-weight: 700; font-size: 1rem; margin: 0 0 8px 0;">Instant Processing</p>
            <p style="color: #64748B !important; font-size: 0.82rem; line-height: 1.5; margin: 0;">
                Upload your Excel and get a fully formatted email in seconds. No manual copy-paste required.
            </p>
        </div>
        """, unsafe_allow_html=True)

    with f2:
        st.markdown("""
        <div style="
            background: #111827; 
            border: 1px solid rgba(0,161,156,0.15); 
            border-radius: 14px; 
            padding: 28px; 
            text-align: center;
            transition: all 0.3s ease;
            height: 200px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        ">
            <p style="font-size: 2.2rem; margin: 0 0 12px 0;">📈</p>
            <p style="color: #F1F5F9 !important; font-weight: 700; font-size: 1rem; margin: 0 0 8px 0;">4-Week Trends</p>
            <p style="color: #64748B !important; font-size: 0.82rem; line-height: 1.5; margin: 0;">
                Automatically tracks and displays a 4-week historical snapshot of your SR & Incident ageing data.
            </p>
        </div>
        """, unsafe_allow_html=True)

    with f3:
        st.markdown("""
        <div style="
            background: #111827; 
            border: 1px solid rgba(0,161,156,0.15); 
            border-radius: 14px; 
            padding: 28px; 
            text-align: center;
            transition: all 0.3s ease;
            height: 200px;
            display: flex;
            flex-direction: column;
            justify-content: center;
        ">
            <p style="font-size: 2.2rem; margin: 0 0 12px 0;">🔒</p>
            <p style="color: #F1F5F9 !important; font-weight: 700; font-size: 1rem; margin: 0 0 8px 0;">Fully Offline</p>
            <p style="color: #64748B !important; font-size: 0.82rem; line-height: 1.5; margin: 0;">
                Runs entirely on your local machine. No data ever leaves your computer. Secure by design.
            </p>
        </div>
        """, unsafe_allow_html=True)
