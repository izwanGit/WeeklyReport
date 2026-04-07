import streamlit as st
import pandas as pd
import json
import os
import datetime
from jinja2 import Environment, FileSystemLoader
import sys
import tempfile
import base64
import traceback

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
MIN_NUMERIC_ROWS = 5  # Sheet must have ≥5 valid numeric ageing rows to qualify

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

def _has_required_columns(df, required_columns):
    """Check if df contains ALL required columns (case-insensitive)."""
    sheet_cols = {str(c).strip().lower() for c in df.columns}
    return all(rc.strip().lower() in sheet_cols for rc in required_columns)

def _count_numeric_rows(df, ageing_col_name):
    """Count how many rows have a valid numeric value in the ageing column."""
    col = find_col(df, ageing_col_name)
    if col is None:
        return 0
    numeric = pd.to_numeric(df[col], errors='coerce')
    return int(numeric.notna().sum())

def detect_valid_sheet(xl, required_columns, ageing_col_name):
    """
    Scan ALL sheets in a pd.ExcelFile. For each sheet:
      1. Load it as a DataFrame
      2. Check it contains ALL required_columns (case-insensitive)
      3. Check that the ageing column has ≥ MIN_NUMERIC_ROWS valid numeric values
    Returns (sheet_name, DataFrame) for the FIRST qualifying sheet, or (None, None).
    Dashboard/image sheets are excluded automatically because they won't have
    enough numeric rows — no name-based filtering needed.
    """
    for sheet_name in xl.sheet_names:
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            if df.empty:
                continue
            if not _has_required_columns(df, required_columns):
                continue
            if _count_numeric_rows(df, ageing_col_name) < MIN_NUMERIC_ROWS:
                continue
            return sheet_name, df
        except Exception:
            continue
    return None, None

def detect_wo_sheet(xl, sr_sheet_name):
    """
    Detect the Work Order detail sheet. It is distinguished from the SR sheet
    by containing 'Work Order ID' AND 'Service Request Ageing Days' columns,
    and it must NOT be the same sheet already identified as the SR sheet.
    """
    wo_required = {"Service Request Ageing Days", "Work Order ID", "Work Order Status"}
    for sheet_name in xl.sheet_names:
        if sheet_name == sr_sheet_name:
            continue
        try:
            df = pd.read_excel(xl, sheet_name=sheet_name)
            if df.empty:
                continue
            if not _has_required_columns(df, wo_required):
                continue
            if _count_numeric_rows(df, "Service Request Ageing Days") < MIN_NUMERIC_ROWS:
                continue
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
    [data-testid="stSidebar"] { border-right: 2px solid #00B1A9 !important; }
    [data-testid="stSidebar"] h1, [data-testid="stSidebar"] h2, [data-testid="stSidebar"] h3 {
        color: #00B1A9 !important; font-weight: 700 !important;
    }
    .stButton > button, .stDownloadButton > button {
        background: #00B1A9 !important;
        color: white !important; border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; transition: all 0.3s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #008C86 !important;
        transform: translateY(-1px) !important; color: white !important;
    }
    [data-testid="stMetric"] {
        background: #FFFFFF !important; border: 1px solid #E2E8F0 !important;
        border-left: 4px solid #00B1A9 !important; border-radius: 12px !important;
        padding: 1.1rem 1.2rem !important; box-shadow: 0 2px 8px rgba(0,0,0,0.04) !important;
    }
    [data-testid="stMetricValue"] { color: #00B1A9 !important; font-weight: 800 !important; font-size: 1.8rem !important; }
    [data-testid="stMetricLabel"] { color: #4A5568 !important; font-weight: 500 !important; }
    .stTabs [data-baseweb="tab"] { font-weight: 500 !important; }
    .stTabs [aria-selected="true"] { color: #00B1A9 !important; font-weight: 700 !important; border-bottom-color: #00B1A9 !important; }
    [data-testid="stFileUploader"] { 
        border: 2px dashed rgba(0, 177, 169, 0.4) !important; 
        border-radius: 12px !important; 
        padding: 16px 20px !important; 
    }
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
<style>
    .banner-title {{ color: #FFFFFF !important; text-transform: uppercase !important; font-weight: 800 !important; text-shadow: 0px 3px 6px rgba(0,0,0,0.3) !important; margin: 0 !important; line-height: 1.1 !important; white-space: nowrap; font-size: clamp(1.4rem, 4vw, 2.2rem) !important; letter-spacing: 0.2px; }}
    .banner-subtitle {{ color: #FFFFFF !important; font-weight: 400 !important; text-shadow: 0px 2px 4px rgba(0,0,0,0.2) !important; margin: 6px 0 0 0 !important; white-space: nowrap; font-size: clamp(0.9rem, 2vw, 1.1rem) !important; opacity: 0.95 !important; }}
</style>
<div style="display: flex; align-items: center; gap: 24px; padding: 28px 40px; background-color: #00B1A9; border-radius: 20px; margin-bottom: 3rem; box-shadow: 0 15px 40px rgba(0, 177, 169, 0.3); overflow: hidden; border: 1px solid rgba(255, 255, 255, 0.15);">
    <img src="{_logo_banner_uri}" style="height: 90px; flex-shrink: 0; filter: drop-shadow(1px 1px 0 white) drop-shadow(-1px -1px 0 white) drop-shadow(1px -1px 0 white) drop-shadow(-1px 1px 0 white);" />
    <div style="min-width: 0;">
        <h1 class="banner-title">Weekly SR &amp; Incident Report Generator</h1>
        <p class="banner-subtitle">Automate your MyGenie Excel exports into production-ready HTML email reports.</p>
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

    st.markdown("---")
    st.markdown("### Open Ticket Counts")
    sr_open_wo = st.number_input("Open WO Tickets", min_value=1, value=1, step=1, help="Total open Work Order ticket count (e.g. 215)")
    inc_open_input = st.number_input("Open INC Tickets", min_value=1, value=1, step=1, help="Total open Incident ticket count (e.g. 7)")

if sr_wo_file and inc_file:
    try:
        # Read each file into ExcelFile ONCE to avoid file-pointer exhaustion
        sr_wo_file.seek(0)
        xl_sr_wo = pd.ExcelFile(sr_wo_file)
        inc_file.seek(0)
        xl_inc = pd.ExcelFile(inc_file)

        # --- Detect sheets by column content + numeric validation ---
        sr_required = {"Service Request Ageing Days", "Service Request ID", "Service Request Status"}
        inc_required = {"Incident Ageing Days", "Incident ID", "Status"}

        sr_sheet_name, df_sr_raw = detect_valid_sheet(xl_sr_wo, sr_required, "Service Request Ageing Days")
        wo_sheet_name, df_wo_raw = detect_wo_sheet(xl_sr_wo, sr_sheet_name) if sr_sheet_name else (None, None)
        inc_sheet_name, df_inc_raw = detect_valid_sheet(xl_inc, inc_required, "Incident Ageing Days")

        # Log all sheets scanned for transparency
        st.info(
            f"SR & WO file sheets: `{xl_sr_wo.sheet_names}`  \n"
            f"Incident file sheets: `{xl_inc.sheet_names}`  \n"
            f"**Detected →** SR: `{sr_sheet_name}`, WO: `{wo_sheet_name}`, INC: `{inc_sheet_name}`"
        )

        if sr_sheet_name is None:
            st.error(f"Could not locate a valid Service Request sheet (need columns {sr_required} with ≥{MIN_NUMERIC_ROWS} numeric ageing rows).")
            st.stop()
        if wo_sheet_name is None:
            st.warning("Work Order detail sheet not found — detail tables will be empty.")
            df_wo_raw = pd.DataFrame()  # empty fallback
        if inc_sheet_name is None:
            st.error(f"Could not locate a valid Incident sheet (need columns {inc_required} with ≥{MIN_NUMERIC_ROWS} numeric ageing rows).")
            st.stop()

        # --- SR Metric Calculations (from SR raw data sheet ONLY) ---
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
        # Use user-provided Open WO ticket count as denominator for >1 day %
        # And use total ageing (>1 day) as denominator for >30 day %
        sr_gt_1_pct  = round((sr_gt_1_count / sr_open_wo * 100) if sr_open_wo > 0 else 0)
        sr_gt_30_pct = round((sr_gt_30_count / sr_gt_1_count * 100) if sr_gt_1_count > 0 else 0)

        # --- SR Details (from WO sheet) ---
        df_wo = df_wo_raw.copy() if not df_wo_raw.empty else pd.DataFrame()
        wo_ageing_col     = find_col(df_wo, "Service Request Ageing Days") if not df_wo.empty else None
        wo_status_col     = find_col(df_wo, "Work Order Status")
        wo_id_col         = find_col(df_wo, "Work Order ID")
        wo_summary_col    = find_col(df_wo, "Work Order Summary")
        wo_customer_col   = find_col(df_wo, "Customer Full Name (Service Request)") or find_col(df_wo, "Customer Full Name")
        wo_reason_col     = find_col(df_wo, "Work Order Status Reason")
        wo_assignee_col   = find_col(df_wo, "Work Order Assignee")

        sr_ageing_gt_30_tickets = []
        sr_ageing_15_30_tickets = []

        if wo_ageing_col is not None and not df_wo.empty:
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

        # --- INC Metric Calculations (from INC raw data sheet ONLY) ---
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
        # Use user-provided Open INC ticket count as denominator for >1 day %
        inc_gt_1_pct = round((inc_gt_1_count / inc_open_input * 100) if inc_open_input > 0 else 0)


        # --- History Tracking ---
        history = load_history()
        short_date = report_date.strftime("%d-%b-%Y")

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

        # Check if record for this exact date already exists
        existing_idx = next((i for i, h in enumerate(history) if h.get("date") == short_date), None)
        
        # Create an in-memory copy for rendering the preview without auto-saving
        render_history = list(history)
        if existing_idx is not None:
            render_history[existing_idx] = new_record
        else:
            render_history.append(new_record)

        def parse_date(date_str):
            try:
                return datetime.datetime.strptime(date_str, "%d-%b-%Y")
            except ValueError:
                return datetime.datetime.min

        render_history.sort(key=lambda x: parse_date(x.get("date", "")))

        if len(render_history) > 4:
            render_history = render_history[-4:]

        # Prepare arrays for template using the in-memory render_history
        trend_dates = [h.get("date", "") for h in render_history]
        sr_trend_gt_30 = [h.get("sr_count_gt_30", 0) for h in render_history]
        sr_trend_15_30 = [h.get("sr_count_15_30", 0) for h in render_history]
        sr_trend_1_14 = [h.get("sr_count_1_14", 0) for h in render_history]
        inc_trend_gt_90 = [h.get("inc_count_gt_90", 0) for h in render_history]
        inc_trend_61_90 = [h.get("inc_count_61_90", 0) for h in render_history]
        inc_trend_31_60 = [h.get("inc_count_31_60", 0) for h in render_history]
        inc_trend_15_30 = [h.get("inc_count_15_30", 0) for h in render_history]
        inc_trend_8_14  = [h.get("inc_count_8_14", 0) for h in render_history]
        inc_trend_3_7   = [h.get("inc_count_3_7", 0) for h in render_history]

        # ========== KPI METRICS CARDS ==========
        st.markdown("### Key Metrics Overview")
        st.markdown(f"<p style='color:#64748B; margin-top:-10px;'>Report date: <b style='color:#00B1A9;'>{report_date_str}</b></p>", unsafe_allow_html=True)

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

        # --- Explicit Save Controls ---
        st.markdown(" ")
        st.markdown(f"<p style='color:#64748B; font-size:12px; margin-bottom: 5px; font-weight:600;'>SNAPSHOT MANAGEMENT</p>", unsafe_allow_html=True)
        if existing_idx is not None:
            c1, c2 = st.columns([1, 4])
            with c1:
                if st.button("Update Saved Snapshot"):
                    history[existing_idx] = new_record
                    history.sort(key=lambda x: parse_date(x.get("date", "")))
                    if len(history) > 4: history = history[-4:]
                    save_history(history)
                    st.rerun()
            with c2:
                st.info(f"✓ The selected date ({short_date}) is already in your History. The table below includes it dynamically.", icon="✅")
        else:
            c1, c2 = st.columns([1, 4])
            with c1:
                if st.button("Save Snapshot to History", type="primary"):
                    history.append(new_record)
                    history.sort(key=lambda x: parse_date(x.get("date", "")))
                    if len(history) > 4: history = history[-4:]
                    save_history(history)
                    st.rerun()
            with c2:
                st.warning(f"This summary is **NOT YET SAVED**. It's temporarily previewed below. Click Save to log {short_date} into history.", icon="⚠️")

        st.markdown("")

        # ========== GENERATE HTML ==========
        env = Environment(loader=FileSystemLoader(BASE_DIR))
        template = env.get_template("template.html")
        html_output = template.render(
            report_date=report_date_str,
            sr_open_wo=sr_open_wo,
            sr_total=sr_total,
            sr_gt_1_count=sr_gt_1_count,
            sr_gt_30_count=sr_gt_30_count,
            sr_ageing_more_than_1_day_pct=sr_gt_1_pct,
            sr_ageing_more_than_30_days_pct=sr_gt_30_pct,
            sr_ageing_gt_30_tickets=sr_ageing_gt_30_tickets,
            sr_ageing_15_30_tickets=sr_ageing_15_30_tickets,
            inc_open_input=inc_open_input,
            inc_total=inc_total,
            inc_gt_1_count=inc_gt_1_count,
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

        tab_preview, tab_source, tab_export, tab_history = st.tabs(["Email Preview", "HTML Source", "Export Options", "Manage History"])

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
                copy_btn_html = f"""
                <html>
                <head>
                <style>
                body {{ margin: 0; padding: 0; display: flex; flex-direction: column; align-items: center; font-family: sans-serif; }}
                button {{
                    background: linear-gradient(135deg, #00A19C 0%, #008C87 100%);
                    background: linear-gradient(135deg, #00B1A9 0%, #008C86 100%);
                    color: white; border: none; border-radius: 10px; font-weight: 600;
                    padding: 0.6rem 1.4rem; font-size: 0.9rem; cursor: pointer;
                    width: 100%; transition: all 0.3s ease;
                }}
                button:hover {{ filter: brightness(1.1); transform: translateY(-1px); }}
                #msg {{ color: #00B1A9; font-size: 0.8rem; font-weight: 600; margin-top: 5px; display: none; }}
                </style>
                </head>
                <body>
                    <button onclick="copyRichText()">Copy Formatted</button>
                    <div id="msg">Copied to clipboard!</div>
                    <div id="source" style="display:none;">{base64.b64encode(html_output.encode('utf-8')).decode('utf-8')}</div>
                    <script>
                    function copyRichText() {{
                        try {{
                            const b64Data = document.getElementById("source").innerText;
                            const html = decodeURIComponent(escape(window.atob(b64Data)));
                            const blobHtml = new Blob([html], {{ type: "text/html" }});
                            const data = [new ClipboardItem({{ ["text/html"]: blobHtml }})];
                            navigator.clipboard.write(data).then(() => {{
                                document.getElementById("msg").style.display = "block";
                                setTimeout(() => document.getElementById("msg").style.display = "none", 3000);
                            }}).catch(err => console.error("Clipboard Error:", err));
                        }} catch (e) {{
                            console.error(e);
                        }}
                    }}
                    </script>
                </body>
                </html>
                """
                st.components.v1.html(copy_btn_html, height=80)
            with exp3:
                if sys.platform == 'win32':
                    if st.button("Push to Outlook Draft", use_container_width=True):
                        success = push_to_outlook(html_output, f"Weekly SR & Incident Report - {report_date_str}")
                        if success:
                            st.success("Draft created in Outlook.")
                else:
                    st.button("Outlook (Windows Only)", use_container_width=True, disabled=True)

        with tab_history:
            st.markdown("### Saved Historical Records")
            saved_data = load_history()
            if not saved_data:
                st.info("No records found in history.json yet. Process and save a snapshot to see it here.")
            else:
                for idx, h in enumerate(saved_data):
                    st.markdown(f"<div style='padding:10px; border:1px solid #E2E8F0; border-radius:8px; margin-bottom:8px;'>", unsafe_allow_html=True)
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        st.markdown(f"**{h.get('date')}** &nbsp;|&nbsp; <span style='color:#718096;'>SR &gt;30d: **{h.get('sr_count_gt_30',0)}**</span> &nbsp;|&nbsp; <span style='color:#718096;'>INC &gt;90d: **{h.get('inc_count_gt_90',0)}**</span>", unsafe_allow_html=True)
                    with col2:
                        if st.button("Delete", key=f"del_{idx}"):
                            saved_data.pop(idx)
                            save_history(saved_data)
                            st.rerun()
                    st.markdown("</div>", unsafe_allow_html=True)
                
                st.markdown("---")
                if st.button("Clear Absolute All History"):
                    save_history([])
                    st.rerun()

    except Exception as e:
        st.error(f"Error processing the files: {e}")
        st.code(traceback.format_exc(), language="text")

else:
    st.markdown("""
    <div style="text-align: center; padding: 70px 40px; background: linear-gradient(180deg, #FFFFFF 0%, #F0FAFA 100%); border: 1px solid #E2E8F0; border-top: 4px solid #00A19C; border-radius: 16px;">
        <h2 style="color: #1A202C !important; font-weight: 800 !important; margin: 0 0 10px 0 !important;">Data Upload Required</h2>
        <p style="color: #718096 !important; max-width: 480px; margin: 0 auto 28px auto; line-height: 1.7;">
            Please upload both your Service Request/Work Order Excel and your Incident Excel using the sidebars.
        </p>
    </div>
    """, unsafe_allow_html=True)
