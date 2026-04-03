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
    # If the application is run as a bundle, the PyInstaller bootloader
    # extends the sys module by a flag frozen=True and sets the app 
    # path into variable _MEIPASS'.
    BASE_DIR = sys._MEIPASS
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Data files
HISTORY_FILE = "history.json"  # Keeps it in the current working directory of the user
TEMPLATE_FILE = os.path.join(BASE_DIR, "template.html")

# Expected Columns
COL_AGEING = "SR Ageing" # Using a generic check since it might be "SR Ageing (days)"
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
    # Try to find the required columns
    ageing_col = get_column_by_prefix(df, "ageing")
    if ageing_col is None:
        st.warning("Could not find 'Ageing' column in the uploaded sheet.")
        return None, None
    
    # Ensure numeric ageing
    df[ageing_col] = pd.to_numeric(df[ageing_col], errors='coerce').fillna(0)
    
    # Create normalized columns for the template
    norm_df = df.copy()
    
    # Function to map columns safely
    def map_col(prefix, default_name):
        col = get_column_by_prefix(df, prefix)
        if col:
            norm_df[default_name] = df[col]
        else:
            norm_df[default_name] = ""
            
    map_col("work order", COL_WO_NO)
    map_col("summary", COL_SUMMARY)
    map_col("user/tsg", COL_USER_TSG)
    map_col("status", COL_STATUS)  # Note: logic might clash if "status" vs "status reason"
    
    # Fine-tuned mappings for status reasoning
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
        mail.Display(True) # True means it opens the draft in a separate window
        return True
    except Exception as e:
        st.error(f"Failed to open Outlook draft: {str(e)}")
        return False

# ----------------------------------------------------
# Main Streamlit App
# ----------------------------------------------------
st.set_page_config(page_title="Weekly Report Auto-Generator", page_icon="📊", layout="wide")

# Petronas styling via markdown
st.markdown("""
<style>
    .reportview-container .main .block-container{
        max-width: 1000px;
    }
    h1, h2, h3 {
        color: #00A19C;
    }
    .stButton>button {
        background-color: #00A19C;
        color: white;
    }
</style>
""", unsafe_allow_html=True)

st.title("📊 Weekly SR & Incident Report Generator")
st.markdown("Automate HTML Email generation from MyGenie Excel exports.")

# Step 1: Inputs
with st.sidebar:
    st.header("Step 1: Setup")
    report_date = st.date_input("Report Date", datetime.date.today())
    report_date_str = report_date.strftime("%d %B %Y")
    
    st.markdown("---")
    uploaded_file = st.file_uploader("Upload MyGenie Excel", type=['xlsx', 'xls'])
    
    # Which sheets are SR and INC?
    if uploaded_file:
        try:
            xl = pd.ExcelFile(uploaded_file)
            sheet_names = xl.sheet_names
            sr_sheet = st.selectbox("Select SR Sheet", sheet_names, index=0)
            inc_sheet = st.selectbox("Select INC Sheet", sheet_names, index=min(1, len(sheet_names)-1))
        except Exception as e:
            st.error(f"Error reading Excelfile: {e}")
            uploaded_file = None

if uploaded_file and sr_sheet and inc_sheet:
    st.header("Step 2: Preview & Export")
    
    try:
        # Load sheets
        df_sr = pd.read_excel(uploaded_file, sheet_name=sr_sheet)
        df_inc = pd.read_excel(uploaded_file, sheet_name=inc_sheet)
        
        # --- SR Processing ---
        norm_sr, sr_ageing = process_tickets(df_sr)
        if norm_sr is not None:
            sr_total = len(norm_sr)
            
            # SR > 30 Days
            sr_gt_30 = norm_sr[norm_sr[COL_AGEING] > 30]
            sr_gt_30_count = len(sr_gt_30)
            
            # SR 15 - 30 Days
            sr_15_30 = norm_sr[(norm_sr[COL_AGEING] >= 15) & (norm_sr[COL_AGEING] <= 30)]
            sr_15_30_count = len(sr_15_30)
            
            # SR 1 - 14 Days
            sr_1_14 = norm_sr[(norm_sr[COL_AGEING] >= 1) & (norm_sr[COL_AGEING] <= 14)]
            sr_1_14_count = len(sr_1_14)
            
            # SR Percents
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
            
            # INC Percents
            inc_gt_1_count = len(norm_inc[norm_inc[COL_AGEING] > 1])
            inc_gt_1_pct = round((inc_gt_1_count / inc_total * 100) if inc_total > 0 else 0, 2)
        
        # --- History Tracking (4 weeks) ---
        history = load_history()
        short_date = report_date.strftime("%d-%b-%Y")
        
        # Check if today is already in history to avoid duplicate appends on reruns
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
                history = history[-4:] # keep only last 4
            save_history(history)
            
        # Prepare arrays for template
        trend_dates = [h["date"] for h in history]
        
        # SR arrays
        sr_trend_gt_30 = [h["sr_count_gt_30"] for h in history]
        sr_trend_15_30 = [h["sr_count_15_30"] for h in history]
        sr_trend_1_14 = [h["sr_count_1_14"] for h in history]
        
        # INC arrays
        inc_trend_gt_90 = [h["inc_count_gt_90"] for h in history]
        inc_trend_61_90 = [h["inc_count_61_90"] for h in history]
        inc_trend_31_60 = [h["inc_count_31_60"] for h in history]
        inc_trend_15_30 = [h["inc_count_15_30"] for h in history]
        inc_trend_8_14  = [h["inc_count_8_14"] for h in history]
        inc_trend_3_7   = [h["inc_count_3_7"] for h in history]
        
        # --- Generate HTML ---
        env = Environment(loader=FileSystemLoader(BASE_DIR))
        try:
            template = env.get_template("template.html")
            html_output = template.render(
                report_date=report_date_str,
                # SR
                sr_total=sr_total,
                sr_ageing_more_than_1_day_pct=sr_gt_1_pct,
                sr_ageing_more_than_30_days_pct=sr_gt_30_pct,
                sr_ageing_gt_30_tickets=sr_gt_30_dict,
                sr_ageing_15_30_tickets=sr_15_30_dict,
                # INC
                inc_total=inc_total,
                inc_ageing_more_than_1_day_pct=inc_gt_1_pct,
                # Trends
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
            
            # Layout
            col1, col2 = st.columns([2, 1])
            with col1:
                st.subheader("Email Preview")
                st.components.v1.html(html_output, height=800, scrolling=True)
                
            with col2:
                st.subheader("Actions")
                
                # 1. Download HTML
                b64 = base64.b64encode(html_output.encode('utf-8')).decode()
                href = f'<a href="data:text/html;base64,{b64}" download="Weekly_Report_{report_date.strftime("%Y%m%d")}.html" style="text-decoration:none;"><button style="width: 100%; border-radius: 4px; padding: 10px; background-color: #00A19C; color: white; border: none; cursor: pointer; font-weight: bold; margin-bottom: 10px;">Export as .html file 💾</button></a>'
                st.markdown(href, unsafe_allow_html=True)
                
                # 2. Push to Outlook
                if sys.platform == 'win32':
                    if st.button("Push to Outlook Draft ✉️", use_container_width=True):
                        success = push_to_outlook(html_output, f"Weekly SR & Incident Report - {report_date_str}")
                        if success:
                            st.success("Draft created in Outlook!")
                else:
                    st.info("ℹ️ Draft push to Outlook is only available on Windows.")
                    
                st.markdown("---")
                
                # 3. View Source / Copy
                st.subheader("Raw HTML Source")
                st.text_area("Copy this text if you wish to paste directly", html_output, height=200)

        except Exception as e:
            st.error(f"Error rendering template: {e}")
            
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        
else:
    st.info("Upload the weekly Excel file to proceed.")
