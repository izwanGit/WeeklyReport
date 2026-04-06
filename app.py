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

# ----------------------------------------------------
# Helper Functions
# ----------------------------------------------------
def find_col(df, target):
    """Find actual column name in df matching target (case-insensitive, stripped)."""
    t = target.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == t:
            return c
    return None

def clean_status(val):
    if pd.isna(val):
        return ""
    return str(val).strip().lower()

def is_active_status(status_str):
    s = clean_status(status_str)
    if not s:
        return False
    inactive = {"closed", "resolved", "cancelled", "canceled"}
    return s not in inactive

def find_sheet_by_columns(xl, required_columns):
    """
    Scans a pd.ExcelFile to find a sheet containing ALL required columns.
    Uses case-insensitive matching. Skips 'Untitled' sheets.
    Returns (sheet_name, DataFrame) or (None, None).
    """
    required_lower = {c.strip().lower() for c in required_columns}
    for sheet_name in xl.sheet_names:
        if "untitled" in sheet_name.lower():
            continue
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            if df.empty or len(df.columns) < len(required_columns):
                continue
            sheet_cols_lower = {str(c).strip().lower() for c in df.columns}
            if required_lower.issubset(sheet_cols_lower):
                return sheet_name, df
        except Exception:
            continue
    return None, None

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

# Premium CSS
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif !important;
    }
    .main .block-container { padding-top: 2rem !important; max-width: 1400px !important; }
    [data-testid="stSidebar"] { border-right: 2px solid #00A19C !important; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #00A19C !important; font-weight: 700 !important;
    }
    .stButton > button, .stDownloadButton > button {
        background: linear-gradient(135deg, #00A19C 0%, #008C87 100%) !important;
        color: white !important; border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; transition: all 0.3s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: linear-gradient(135deg, #00BFB8 0%, #00A19C 100%) !important;
        transform: translateY(-1px) !important; color: white !important;
    }
    [data-testid="stMetric"] {
        background: #FFFFFF !important; border: 1px solid #E2E8F0 !important;
        border-left: 4px solid #00A19C !important; border-radius: 12px !important;
        padding: 1.1rem 1.2rem !important; box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
    }
    [data-testid="stMetricValue"] { color: #00A19C !important; font-weight: 800 !important; font-size: 1.8rem !important; }
    [data-testid="stMetricLabel"] { color: #4A5568 !important; font-weight: 500 !important; }
    .stTabs [data-baseweb="tab"] { border-radius: 10px !important; background: #FFFFFF !important; }
    .stTabs [aria-selected="true"] { background: linear-gradient(135deg, #00A19C 0%, #008C87 100%) !important; color: white !important; }
    [data-testid="stFileUploader"] { border: 2px dashed rgba(0, 161, 156, 0.35) !important; border-radius: 12px !important; }
    #MainMenu {visibility: hidden;} footer {visibility: hidden;}
    .stDeployButton {display: none !important;}
</style>
""", unsafe_allow_html=True)

# Header
def _image_to_data_uri(path, mime_type):
    try:
        with open(os.path.join(BASE_DIR, path), 'rb') as f:
            data = f.read()
        return f"data:{mime_type};base64,{base64.b64encode(data).decode()}"
    except:
        return ""

_logo_banner_uri = _image_to_data_uri("PETRONAS_LOGO_SQUARE.png", "image/png")
_logo_sidebar_uri = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg", "image/svg+xml")

st.markdown(f"""
<div style="display: flex; align-items: center; gap: 24px; padding: 24px 30px; background: linear-gradient(135deg, #00A19C 0%, #008C87 100%); border-radius: 16px; margin-bottom: 2rem;">
    <img src="{_logo_banner_uri}" style="height: 85px;" />
    <div>
        <h1 style="margin: 0 !important; color: #FFFFFF !important; font-weight: 800 !important;">Weekly SR & Incident Report Generator</h1>
        <p style="margin: 6px 0 0 0 !important; color: #E6F7F6 !important;">Automate your MyGenie Excel exports into production-ready HTML email reports.</p>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar
with st.sidebar:
    st.markdown(f"""
    <div style="text-align:center; padding: 12px 0 16px 0;">
        <img src="{_logo_sidebar_uri}" style="height: 65px;" />
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("### Report Configuration")
    report_date = st.date_input("Report Date", datetime.date.today())
    report_date_str = report_date.strftime("%d %B %Y")
    
    st.markdown("---")
    st.markdown("### Data Upload")
    
    sr_wo_file = st.file_uploader("Upload SR & Work Order Excel", type=['xlsx', 'xls'], key="sr_wo")
    inc_file = st.file_uploader("Upload Incident Excel", type=['xlsx', 'xls'], key="inc")

if sr_wo_file and inc_file:
    try:
        # Read each file into ExcelFile ONCE to avoid file-pointer exhaustion
        sr_wo_file.seek(0)
        xl_sr_wo = pd.ExcelFile(sr_wo_file)
        inc_file.seek(0)
        xl_inc = pd.ExcelFile(inc_file)

        sr_required = {"Service Request Ageing Days", "Service Request ID", "Service Request Status"}
        wo_required = {"Service Request Ageing Days", "Work Order ID", "Work Order Summary", "Work Order Status"}
        inc_required = {"Incident Ageing Days", "Incident ID", "Status"}

        sr_sheet_name, df_sr_raw = find_sheet_by_columns(xl_sr_wo, sr_required)
        wo_sheet_name, df_wo_raw = find_sheet_by_columns(xl_sr_wo, wo_required)
        inc_sheet_name, df_inc_raw = find_sheet_by_columns(xl_inc, inc_required)

        st.info(f"Detected — SR: `{sr_sheet_name}`, WO: `{wo_sheet_name}`, INC: `{inc_sheet_name}`")

        if sr_sheet_name is None:
            st.error(f"Could not locate the Service Request sheet. Required columns: {sr_required}")
            st.stop()
        if wo_sheet_name is None:
            st.error(f"Could not locate the Work Order sheet. Required columns: {wo_required}")
            st.stop()
        if inc_sheet_name is None:
            st.error(f"Could not locate the Incident sheet. Required columns: {inc_required}")
            st.stop()

        # --- SR Metric Calculations (from SR sheet) ---
        df_sr = df_sr_raw.copy()
        sr_ageing_col = find_col(df_sr, "Service Request Ageing Days")
        sr_status_col = find_col(df_sr, "Service Request Status")
        df_sr[sr_ageing_col] = pd.to_numeric(df_sr[sr_ageing_col], errors='coerce')
        df_sr = df_sr.dropna(subset=[sr_ageing_col])
        df_sr = df_sr[df_sr[sr_status_col].apply(is_active_status)]

        sr_total = len(df_sr)
        sr_gt_30_count = int((df_sr[sr_ageing_col] > 30).sum())
        sr_15_30_count = int(((df_sr[sr_ageing_col] >= 15) & (df_sr[sr_ageing_col] <= 30)).sum())
        sr_1_14_count  = int(((df_sr[sr_ageing_col] >= 1) & (df_sr[sr_ageing_col] <= 14)).sum())
        sr_gt_1_count  = int((df_sr[sr_ageing_col] > 1).sum())
        sr_gt_1_pct  = round((sr_gt_1_count / sr_total * 100) if sr_total > 0 else 0, 2)
        sr_gt_30_pct = round((sr_gt_30_count / sr_total * 100) if sr_total > 0 else 0, 2)

        # --- SR Details (from WO sheet) ---
        df_wo = df_wo_raw.copy()
        wo_ageing_col     = find_col(df_wo, "Service Request Ageing Days")
        wo_status_col     = find_col(df_wo, "Work Order Status")
        wo_id_col         = find_col(df_wo, "Work Order ID")
        wo_summary_col    = find_col(df_wo, "Work Order Summary")
        wo_customer_col   = find_col(df_wo, "Customer Full Name (Service Request)") or find_col(df_wo, "Customer Full Name")
        wo_reason_col     = find_col(df_wo, "Work Order Status Reason")
        wo_assignee_col   = find_col(df_wo, "Work Order Assignee")

        df_wo[wo_ageing_col] = pd.to_numeric(df_wo[wo_ageing_col], errors='coerce')
        df_wo = df_wo.dropna(subset=[wo_ageing_col])
        if wo_status_col:
            df_wo = df_wo[df_wo[wo_status_col].apply(is_active_status)]

        def _safe(row, col):
            if col is None: return ""
            v = row.get(col, "")
            return "" if pd.isna(v) else str(v)

        def extract_wo_records(subset):
            records = []
            for _, row in subset.iterrows():
                records.append({
                    'SR Ageing': int(row[wo_ageing_col]),
                    'Work Order No.': _safe(row, wo_id_col),
                    'Summary': _safe(row, wo_summary_col),
                    'User/TSG': _safe(row, wo_customer_col),
                    'WO Status': _safe(row, wo_status_col),
                    'WO Status Reason': _safe(row, wo_reason_col),
                    'Assignee': _safe(row, wo_assignee_col),
                })
            return records

        df_wo_gt_30  = df_wo[df_wo[wo_ageing_col] > 30]
        df_wo_15_30  = df_wo[(df_wo[wo_ageing_col] >= 15) & (df_wo[wo_ageing_col] <= 30)]
        sr_ageing_gt_30_tickets  = extract_wo_records(df_wo_gt_30)
        sr_ageing_15_30_tickets  = extract_wo_records(df_wo_15_30)

        # --- INC Metric Calculations (from INC sheet) ---
        df_inc = df_inc_raw.copy()
        inc_ageing_col = find_col(df_inc, "Incident Ageing Days")
        inc_status_col = find_col(df_inc, "Status")
        inc_active_col = find_col(df_inc, "Active Incident")

        df_inc[inc_ageing_col] = pd.to_numeric(df_inc[inc_ageing_col], errors='coerce')
        df_inc = df_inc.dropna(subset=[inc_ageing_col])
        if inc_active_col:
            df_inc = df_inc[df_inc[inc_active_col].astype(str).str.strip().str.lower() == "yes"]
        if inc_status_col:
            df_inc = df_inc[df_inc[inc_status_col].apply(is_active_status)]

        inc_total = len(df_inc)
        inc_gt_90_count = int((df_inc[inc_ageing_col] > 90).sum())
        inc_61_90_count = int(((df_inc[inc_ageing_col] >= 61) & (df_inc[inc_ageing_col] <= 90)).sum())
        inc_31_60_count = int(((df_inc[inc_ageing_col] >= 31) & (df_inc[inc_ageing_col] <= 60)).sum())
        inc_15_30_count = int(((df_inc[inc_ageing_col] >= 15) & (df_inc[inc_ageing_col] <= 30)).sum())
        inc_8_14_count  = int(((df_inc[inc_ageing_col] >= 8)  & (df_inc[inc_ageing_col] <= 14)).sum())
        inc_3_7_count   = int(((df_inc[inc_ageing_col] >= 3)  & (df_inc[inc_ageing_col] <= 7)).sum())
        inc_gt_1_count  = int((df_inc[inc_ageing_col] > 1).sum())
        inc_gt_1_pct = round((inc_gt_1_count / inc_total * 100) if inc_total > 0 else 0, 2)


        # --- History Tracking ---
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
        st.markdown("### Key Metrics Overview")
        st.markdown(f"<p style='color:#64748B; margin-top:-10px;'>Report date: <b style='color:#00A19C;'>{report_date_str}</b></p>", unsafe_allow_html=True)

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1:
            st.metric("Total SR Tickets (Active)", sr_total)
        with m2:
            st.metric("SR Ageing > 30d", sr_gt_30_count)
        with m3:
            st.metric("SR > 1 day %", f"{sr_gt_1_pct}%")
        with m4:
            st.metric("Total INC Tickets (Active)", inc_total)
        with m5:
            st.metric("INC > 1 day %", f"{inc_gt_1_pct}%")

        st.markdown("")

        # ========== GENERATE HTML ==========
        env = Environment(loader=FileSystemLoader(BASE_DIR))
        template = env.get_template("template.html")
        html_output = template.render(
            report_date=report_date_str,
            sr_total=sr_total,
            sr_ageing_more_than_1_day_pct=sr_gt_1_pct,
            sr_ageing_more_than_30_days_pct=sr_gt_30_pct,
            sr_ageing_gt_30_tickets=sr_ageing_gt_30_tickets,
            sr_ageing_15_30_tickets=sr_ageing_15_30_tickets,
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

        tab_preview, tab_source, tab_export = st.tabs(["Email Preview", "HTML Source", "Export Options"])

        with tab_preview:
            st.markdown("""<div style="background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 14px; padding: 8px; margin-top: 10px; box-shadow: 0 4px 12px rgba(0, 0, 0, 0.04);">""", unsafe_allow_html=True)
            st.components.v1.html(html_output, height=900, scrolling=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with tab_source:
            st.code(html_output, language="html")

        with tab_export:
            st.markdown("### Export Actions")
            exp1, exp2, exp3 = st.columns(3)

            with exp1:
                st.download_button(
                    label="Download .html",
                    data=html_output,
                    file_name=f"Weekly_Report_{report_date.strftime('%Y%m%d')}.html",
                    mime="text/html",
                    use_container_width=True
                )

            with exp2:
                if st.button("Generate Snippet", use_container_width=True):
                    st.info("Please use the HTML Source tab to copy the raw text, or Download .html and open it in double click.")
            
            with exp3:
                if sys.platform == 'win32':
                    if st.button("Push to Outlook Draft", use_container_width=True):
                        success = push_to_outlook(html_output, f"Weekly SR & Incident Report - {report_date_str}")
                        if success:
                            st.success("Draft created in Outlook.")
                else:
                    st.button("Outlook (Windows Only)", use_container_width=True, disabled=True)

    except Exception as e:
        st.error(f"Error processing the files: {e}")

else:
    st.markdown("""
    <div style="text-align: center; padding: 70px 40px; background: linear-gradient(180deg, #FFFFFF 0%, #F0FAFA 100%); border: 1px solid #E2E8F0; border-top: 4px solid #00A19C; border-radius: 16px;">
        <h2 style="color: #1A202C !important; font-weight: 800 !important; margin: 0 0 10px 0 !important;">Data Upload Required</h2>
        <p style="color: #718096 !important; max-width: 480px; margin: 0 auto 28px auto; line-height: 1.7;">
            Please upload both your Service Request/Work Order Excel and your Incident Excel using the sidebars.
        </p>
    </div>
    """, unsafe_allow_html=True)
