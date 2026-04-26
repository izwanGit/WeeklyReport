import streamlit as st
import pandas as pd
import json
import os
import datetime
import time
import requests
from jinja2 import Environment, FileSystemLoader
import sys
import tempfile
import base64
import traceback
import io

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
    EXE_DIR  = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    EXE_DIR  = BASE_DIR

HISTORY_FILE  = os.path.join(EXE_DIR,  "history.json")
TEMPLATE_FILE = os.path.join(BASE_DIR, "template.html")

# ----------------------------------------------------
# Dashboard Auto-Fetch Configuration
# ----------------------------------------------------
DASHBOARD_URL = (
    "https://mygenieplus-ir1.onbmc.com/dashboards/api/datasources/proxy"
    "/uid/Uf8LY07Vk/api/arsys/v1.0/report/arsqlquery"
)

DASHBOARD_HEADERS = {
    "accept":             "application/json, text/plain, */*",
    "content-type":       "application/json",
    "x-ar-client-type":  "4021",
    "x-ds-authorization": "IMS-JWT JWT PLACEHOLDER",
    "x-grafana-device-id": "c038f697c5ec05209addd40a9fbf77bb",
    "x-grafana-org-id":  "204007533",
    "x-requested-by":    "undefined",
    "origin":  "https://mygenieplus-ir1.onbmc.com",
    "referer": "https://mygenieplus-ir1.onbmc.com/",
    "user-agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36 Edg/147.0.0.0"
    ),
}

MYGENIE_DOMAIN = "mygenieplus-ir1.onbmc.com"

COOKIE_BRIDGE_URL = "http://localhost:17731"

# ----------------------------------------------------
# Auto Cookie Reader (Cookie Bridge & browser-cookie3)
# ----------------------------------------------------
def get_browser_cookies() -> dict:
    """
    Priority order:
      1. Local Cookie Bridge receiver (extension → localhost server)
      2. browser-cookie3 direct read (original method, fallback)
    Returns a plain dict of cookies, or {} on failure.
    """
    diag_lines = []

    # ----------------------------------------------------------------
    # METHOD 1: Cookie Bridge (extension sends cookies to local server)
    # ----------------------------------------------------------------
    try:
        resp = requests.get(
            f"{COOKIE_BRIDGE_URL}/get",
            timeout=2,
        )
        if resp.status_code == 200:
            data       = resp.json()
            cookies    = data.get("cookies", {})
            saved_at   = data.get("saved_at", "unknown time")
            age        = data.get("age_seconds", 0)

            if cookies:
                age_str = (
                    f"{age // 3600}h {(age % 3600) // 60}m ago"
                    if age > 3600
                    else f"{age // 60}m ago"
                    if age > 60
                    else f"{age}s ago"
                )
                st.session_state['_cookie_source'] = f"extension ({age_str})"
                st.session_state['_cookie_diag']   = (
                    f"[Success] Cookie Bridge: {len(cookies)} cookie(s) received via extension\n"
                    f"   Saved: {saved_at}"
                )
                return cookies
            else:
                diag_lines.append("[Warning] Cookie Bridge: server running but no cookies cached yet.")
        elif resp.status_code == 404:
            diag_lines.append(
                "[Info] Cookie Bridge: server running but no cookies yet — "
                "click the extension button in Edge toolbar."
            )
        else:
            diag_lines.append(f"[Warning] Cookie Bridge: unexpected status {resp.status_code}")

    except requests.exceptions.ConnectionError:
        diag_lines.append(
            "[Info] Cookie Bridge not running. "
            "Start cookie_receiver.py for the best experience."
        )
    except Exception as e:
        diag_lines.append(f"[Warning] Cookie Bridge error: {e}")

    # ----------------------------------------------------------------
    # METHOD 2: browser-cookie3 (original fallback)
    # ----------------------------------------------------------------
    try:
        import browser_cookie3
        diag_lines.append(f"[Status] Trying browser-cookie3 fallback…")

        loaders = [
            ("Edge",   browser_cookie3.edge),
            ("Chrome", browser_cookie3.chrome),
        ]
        for name, loader in loaders:
            try:
                cj      = loader(domain_name=MYGENIE_DOMAIN)
                cookies = {c.name: c.value for c in cj}
                if cookies:
                    diag_lines.append(f"[Success] {name}: got {len(cookies)} cookie(s) via browser-cookie3")
                    st.session_state['_cookie_source'] = f"browser-cookie3 ({name})"
                    st.session_state['_cookie_diag']   = "\n".join(diag_lines)
                    return cookies
                else:
                    diag_lines.append(f"[Info] {name}: 0 cookies found for domain")
            except PermissionError as e:
                diag_lines.append(f"[Warning] {name} locked (browser running): {e}")
            except Exception as e:
                diag_lines.append(f"[Warning] {name}: {type(e).__name__}: {e}")

    except ImportError:
        diag_lines.append("[Warning] browser-cookie3 not installed (pip install browser-cookie3)")

    # ----------------------------------------------------------------
    # All methods failed
    # ----------------------------------------------------------------
    st.session_state['_cookie_source'] = None
    st.session_state['_cookie_diag']   = "\n".join(diag_lines)
    return {}


def _year_start_ms() -> int:
    """Unix-ms for Jan 1 of the current year."""
    year = datetime.date.today().year
    return int(time.mktime(time.strptime(f"{year}-01-01", "%Y-%m-%d"))) * 1000


def _now_ms() -> int:
    """Current time in Unix-ms."""
    return int(time.time() * 1000)


def _post_query(sql: str, cookies: dict):
    """
    POST an AR SQL query to the BMC Helix datasource proxy.
    Returns the integer in rows[0][0], or None on any failure.
    """
    payload = {
        "date_format":      "DD/MM/YYYY",
        "date_time_format": "DD/MM/YYYY HH:MM:SS",
        "output_type":      "Table",
        "sql":              sql,
    }
    try:
        resp = requests.post(
            DASHBOARD_URL,
            headers=DASHBOARD_HEADERS,
            cookies=cookies,
            json=payload,
            timeout=15,
        )
        resp.raise_for_status()
        data = resp.json()
        return int(data[0]["rows"][0][0])
    except Exception:
        return None


def fetch_open_wo(cookies: dict):
    """Fetch open Work Order count (year-to-date, MYCAREERX SUPPORT)."""
    from_ms, to_ms = _year_start_ms(), _now_ms()
    sql = f"""SELECT DISTINCT
COUNT(DISTINCT`SRM:Request`.`Request Number`) AS C1

FROM
`SRM:Request`
INNER JOIN `AR System Schema`.`WOI:WorkOrder`
ON ( `SRM:Request`.`Request Number` = `WOI:WorkOrder`.`SRID`  AND `WOI:WorkOrder`.`Work Order ID` LIKE '%ICT_WO%' AND `WOI:WorkOrder`.`Work Order ID` IS NOT NULL)
LEFT OUTER JOIN (`SLM:Measurement` AS `SLM:Measurement1`)
ON (`WOI:WorkOrder`.`Work Order ID` = `SLM:Measurement1`.`ApplicationUserFriendlyID`)
LEFT OUTER JOIN (`SLM:Measurement` AS `SLM:Measurement2`)
ON (`SRM:Request`.`Request Number` = `SLM:Measurement2`.`ApplicationUserFriendlyID`)
WHERE(
`WOI:WorkOrder`.`ASORG` IN ('ARJDBC6460AC66AB204CA7BE8869BB9AF532F9') 
AND `WOI:WorkOrder`.`ASGRP` IN ('MYCAREERX SUPPORT') 
AND `WOI:WorkOrder`.`Request Assignee` IN ('ARJDBC6460AC66AB204CA7BE8869BB9AF532F9')
AND `SRM:Request`.`Submit Date` between {from_ms}/1000 AND {to_ms}/1000
AND `SRM:Request`.`Status` Not IN ('Cancelled', 'Closed')
)
LIMIT 50000 OFFSET 0"""
    return _post_query(sql, cookies)


def fetch_open_inc(cookies: dict):
    """Fetch open Incident count (year-to-date, MYCAREERX SUPPORT)."""
    from_ms, to_ms = _year_start_ms(), _now_ms()
    sql = f"""SELECT DISTINCT
COUNT(DISTINCT`HPD:Help Desk`.`Incident Number`) AS C1
FROM
`HPD:Help Desk`
INNER JOIN (`SLM:Measurement` )
ON (`HPD:Help Desk`.`Incident Number` = `SLM:Measurement`.`ApplicationUserFriendlyID`
AND `HPD:Help Desk`.`Incident Number` LIKE '%ICT_INC%' AND `HPD:Help Desk`.`Incident Number` IS NOT NULL AND `SLM:Measurement`.`SLACategory` = 'Service Level Agreement' 
AND `HPD:Help Desk`.`Status` IN ('Assigned','Pending','In Progress'))
WHERE 
(
`HPD:Help Desk`.`Reported Date` between {from_ms}/1000 AND {to_ms}/1000
AND `HPD:Help Desk`.`Assigned Group` In ('MYCAREERX SUPPORT')
AND `HPD:Help Desk`.`Assigned Support Organization` IN ('ARJDBC6460AC66AB204CA7BE8869BB9AF532F9')
AND `HPD:Help Desk`.`Assignee` IN ('ARJDBC6460AC66AB204CA7BE8869BB9AF532F9')
)
LIMIT 50000 OFFSET 0"""
    return _post_query(sql, cookies)


# ----------------------------------------------------
# Helper Functions
# ----------------------------------------------------
MIN_NUMERIC_ROWS = 1


def find_col(df, target):
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
    sheet_cols = {str(c).strip().lower() for c in df.columns}
    return all(rc.strip().lower() in sheet_cols for rc in required_columns)


def _count_numeric_rows(df, ageing_col_name):
    col = find_col(df, ageing_col_name)
    if col is None:
        return 0
    numeric = pd.to_numeric(df[col], errors='coerce')
    return int(numeric.notna().sum())


def detect_valid_sheet(xl, required_columns, ageing_col_name):
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
        import pythoncom
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Subject = subject
        mail.HTMLBody = html_body
        mail.Display(True)
        return True
    except Exception as e:
        st.error(f"Failed to open Outlook draft: {str(e)}")
        return False
    finally:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except:
            pass


# ----------------------------------------------------
# Page Config & Premium Styling
# ----------------------------------------------------
st.set_page_config(
    page_title="Weekly Report | PETRONAS",
    page_icon=os.path.join(BASE_DIR, "PETRONAS_LOGO_SQUARE.png"),
    layout="wide",
    initial_sidebar_state="expanded"
)


def _image_to_data_uri(path, mime_type):
    try:
        with open(os.path.join(BASE_DIR, path), 'rb') as f:
            data = f.read()
        return f"data:{mime_type};base64,{base64.b64encode(data).decode()}"
    except:
        return ""


_logo_square_uri   = _image_to_data_uri("PETRONAS_LOGO_SQUARE.png",      "image/png")
_logo_sidebar_uri  = _image_to_data_uri("PETRONAS_LOGO_HORIZONTAL.svg",  "image/svg+xml")

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
    html, body, [data-testid="stAppViewContainer"] {
        font-family: 'Inter', sans-serif !important;
    }
    [data-testid="stDecoration"],
    [data-testid="stStatusWidget"],
    [data-testid="stSidebarNav"],
    .stDeployButton { display: none !important; visibility: hidden !important; }
    header[data-testid="stHeader"] { background: transparent !important; border-bottom: none !important; }
    [data-testid="stAppViewContainer"] > .main { transition: none !important; }
    [data-testid="stSidebar"] { animation: none !important; }
    [data-testid="stSidebarNav"],
    [data-testid="stSidebarNavItems"],
    [data-testid="stSidebarNavSeparator"],
    [data-testid="stStatusWidget"] { display: none !important; visibility: hidden !important; }
    header[data-testid="stHeader"] { background: #F8FAFC !important; }
    [data-testid="stSidebar"] { border-right: none !important; }
    .main .block-container { padding-top: 0.5rem !important; max-width: 1400px !important; }
    [data-testid="stSidebar"] h1,
    [data-testid="stSidebar"] h2,
    [data-testid="stSidebar"] h3 { color: #00B1A9 !important; font-weight: 700 !important; }
    .stButton > button, .stDownloadButton > button {
        background: #00B1A9 !important; color: white !important;
        border: none !important; border-radius: 10px !important;
        font-weight: 600 !important; transition: all 0.3s ease !important;
    }
    .stButton > button:hover, .stDownloadButton > button:hover {
        background: #008C86 !important; transform: translateY(-1px) !important; color: white !important;
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
        border-radius: 12px !important; padding: 16px 20px !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h1,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h3 { margin-bottom: 0px !important; }
    [data-testid="stSidebar"] blockquote { margin-bottom: 0px !important; }
    [data-testid="stSidebar"] .stNumberInput,
    [data-testid="stSidebar"] .stDateInput,
    [data-testid="stSidebar"] .stFileUploader { margin-bottom: -10px !important; }
    #MainMenu { visibility: hidden; }
    footer { visibility: hidden; }
    .stDeployButton, [data-testid="stDeployButton"], [data-testid="stAppDeployButton"] { display: none !important; }
    .genie-link {
        font-size: 0.85rem; font-weight: 500;
        color: #31333F !important; text-decoration: none !important;
        transition: all 0.2s ease !important; cursor: pointer !important;
    }
    .genie-link:hover { color: #00B1A9 !important; text-decoration: none !important; }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<a href="/" target="_self" style="text-decoration:none;display:inline-flex;align-items:center;gap:8px;font-weight:600;color:#64748B;margin-bottom:16px;transition:color 0.2s ease;">
    <svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
        <line x1="19" y1="12" x2="5" y2="12"></line>
        <polyline points="12 19 5 12 12 5"></polyline>
    </svg>
    Back to Hub
</a>
""", unsafe_allow_html=True)

st.markdown(f"""
<style>
.banner-title {{color:#FFFFFF!important;text-transform:uppercase!important;font-weight:800!important;text-shadow:0px 2px 4px rgba(0,0,0,0.3)!important;margin:0!important;line-height:1.1!important;white-space:nowrap;font-size:clamp(1.2rem,3.5vw,1.8rem)!important;letter-spacing:0.1px;}}
.banner-subtitle {{color:#FFFFFF!important;font-weight:400!important;text-shadow:0px 1px 3px rgba(0,0,0,0.2)!important;margin:4px 0 0 0!important;white-space:nowrap;font-size:clamp(0.85rem,2vw,1.0rem)!important;opacity:0.95!important;}}
</style>
<div style="display:flex;align-items:center;gap:24px;padding:22px 32px;background-color:#00B1A9;border-radius:20px;margin-bottom:2rem;box-shadow:0 12px 35px rgba(0,177,169,0.25);overflow:hidden;border:1px solid rgba(255,255,255,0.15);">
<img src="{_logo_square_uri}" style="height:80px;flex-shrink:0;filter:drop-shadow(1px 1px 0 white) drop-shadow(-1px -1px 0 white) drop-shadow(1px -1px 0 white) drop-shadow(-1px 1px 0 white);"/>
<div style="min-width:0;">
<h1 class="banner-title">Weekly SR &amp; Incident Report</h1>
<p class="banner-subtitle">Automate your MyGenie Excel exports into production-ready HTML email reports.</p>
</div>
</div>
""", unsafe_allow_html=True)

# ----------------------------------------------------
# Sidebar
# ----------------------------------------------------
with st.sidebar:
    st.markdown(f"""
<div style="text-align:center;padding:8px 0 20px 0;">
    <a href="/" target="_self" style="display:inline-block;">
        <img src="{_logo_sidebar_uri}" style="height:56px;transition:transform 0.2s;cursor:pointer;"
             onmouseover="this.style.transform='scale(1.05)'"
             onmouseout="this.style.transform='scale(1)'"/>
    </a>
</div>
""", unsafe_allow_html=True)

    # --------------------------------------------------
    # Open Ticket Counts (with optional Auto Sync)
    # --------------------------------------------------
    st.markdown("### Open Ticket Counts")

    if "auto_wo" not in st.session_state:
        st.session_state.auto_wo = None
    if "auto_inc" not in st.session_state:
        st.session_state.auto_inc = None
    if "sync_status" not in st.session_state:
        st.session_state.sync_status = None
    if "sync_error" not in st.session_state:
        st.session_state.sync_error = False

    if st.button("Sync Live Tickets", use_container_width=True):
        with st.spinner("Fetching live counts..."):
            live_cookies = get_browser_cookies()
            if live_cookies:
                st.session_state.auto_wo = fetch_open_wo(live_cookies)
                st.session_state.auto_inc = fetch_open_inc(live_cookies)
                
                src = st.session_state.get('_cookie_source', '')
                if "extension" in src:
                    st.session_state.sync_status = "Connected via Extension"
                else:
                    st.session_state.sync_status = "Connected via Local Browser"
                st.session_state.sync_error = False
            else:
                diag = st.session_state.get('_cookie_diag', '')
                if 'Cookie Bridge not running' in diag:
                    st.session_state.sync_status = "Background server not running"
                else:
                    st.session_state.sync_status = "Please click Extension button first"
                st.session_state.sync_error = True

    st.markdown(
        "<div style='text-align: center; margin-top: -10px; margin-bottom: 12px; font-size: 0.75rem; color: #94A3B8; font-weight: 500;'>"
        "Pulls real-time counts from your active browser session."
        "</div>", 
        unsafe_allow_html=True
    )

    if st.session_state.sync_status:
        if st.session_state.sync_error:
            st.error(f"Sync Failed: {st.session_state.sync_status}")
            with st.expander("Diagnostics", expanded=False):
                st.code(st.session_state.get('_cookie_diag', ''), language="text")
        else:
            st.success(st.session_state.sync_status)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("<a href='https://mygenieplus-ir1.onbmc.com/dashboards/d/aegyhutg26kn4a/f350ff42-68d2-5195-bce1-6a86eeaf6336?orgId=204007533&var-ASORG=All&var-AssignedGroup=MYCAREERX%20SUPPORT&var-assignee=All&var-Status=All' target='_blank' class='genie-link'>Open WO ↗</a>", unsafe_allow_html=True)
        sr_open_wo = st.number_input(
            "Open WO",
            min_value=0,
            value=st.session_state.auto_wo if st.session_state.auto_wo is not None else 1,
            step=1,
            help="Total open Work Order ticket count (e.g. 215)",
            label_visibility="collapsed"
        )
    with c2:
        st.markdown("<a href='https://mygenieplus-ir1.onbmc.com/dashboards/d/beg9bk10a07i8e/39afb7b?orgId=204007533&var-Assigned_Support_Org=All&var-AssignedGroup=MYCAREERX%20SUPPORT&var-Assignee=All&var-SLA=All' target='_blank' class='genie-link'>Open INC ↗</a>", unsafe_allow_html=True)
        inc_open_input = st.number_input(
            "Open INC",
            min_value=0,
            value=st.session_state.auto_inc if st.session_state.auto_inc is not None else 1,
            step=1,
            help="Total open Incident ticket count (e.g. 7)",
            label_visibility="collapsed"
        )

    st.markdown("<div style='margin-top: -30px;'></div>", unsafe_allow_html=True)
    st.markdown("### Report Settings")
    report_date = st.date_input("Report Date", datetime.date.today())
    report_date_str = report_date.strftime("%d %B %Y")
    
    # --------------------------------------------------
    # Data Source (with auto SharePoint detection)
    # --------------------------------------------------
    st.markdown("<div style='margin-top: -25px;'></div>", unsafe_allow_html=True)
    st.markdown("### Data Upload")

    user_profile    = os.environ.get('USERPROFILE', '')
    default_inc_path = os.path.join(
        user_profile, "OneDrive - PETRONAS",
        "SAP HR - HCSM Ticket Monitoring Dashboard",
        "Ticketing Data", "Ageing Incident",
        "Incident Ageing Raw Data (Daily).xlsx",
    )
    default_sr_path = os.path.join(
        user_profile, "OneDrive - PETRONAS",
        "SAP HR - HCSM Ticket Monitoring Dashboard",
        "Ticketing Data", "Ageing Service Request",
        "Service Request Ageing Raw Data [Daily].xlsx",
    )

    sync_active = os.path.exists(default_inc_path) and os.path.exists(default_sr_path)

    if sync_active:
        st.success("Live SharePoint Sync Active")
        st.caption("Auto-using synced data. Upload below to override.")

    st.markdown("<a href='https://mygenieplus-ir1.onbmc.com/dashboards/d/ce3wv282zk1kwd/service-request-and-work-order-ageing-raw-data?orgId=204007533&var-Ownership=All&var-Assignee_Group=MYCAREERX%20SUPPORT&var-Assigned_Support_Org=All' target='_blank' class='genie-link'>SR & WO Excel ↗</a>", unsafe_allow_html=True)
    uploaded_sr_wo = st.file_uploader("SR & WO Excel", type=['xlsx', 'xls'], key="sr_wo", label_visibility="collapsed")
    
    st.markdown("<a href='https://mygenieplus-ir1.onbmc.com/dashboards/d/ddxo5d7th1gqob/incident-ageing-raw-data?orgId=204007533&var-Assignee_Login=All&var-Assigned_Group=MYCAREERX%20SUPPORT&var-Assigned_Support_Org=All&var-enableOverridesForExcel=true' target='_blank' class='genie-link'>Incident Excel ↗</a>", unsafe_allow_html=True)
    uploaded_inc = st.file_uploader("Incident Excel", type=['xlsx', 'xls'], key="inc", label_visibility="collapsed")

    final_sr_wo_file = uploaded_sr_wo if uploaded_sr_wo else (default_sr_path if sync_active else None)
    final_inc_file   = uploaded_inc   if uploaded_inc   else (default_inc_path if sync_active else None)


# ----------------------------------------------------
# Main Processing Logic
# ----------------------------------------------------
if final_sr_wo_file and final_inc_file:
    try:
        def safe_load_excel(file_or_path, name_label):
            if isinstance(file_or_path, str):
                try:
                    with open(file_or_path, "rb") as f:
                        file_bytes = f.read()
                    return pd.ExcelFile(io.BytesIO(file_bytes))
                except PermissionError:
                    st.error(
                        f"**Permission Denied!** The {name_label} file is currently locked.\n\n"
                        f"1. Close Microsoft Excel if you have this file open.\n"
                        f"2. Ensure OneDrive sync is not paused.\n\n"
                        f"**Locked File:** `{file_or_path}`"
                    )
                    st.stop()
            else:
                file_or_path.seek(0)
                return pd.ExcelFile(file_or_path)

        xl_sr_wo = safe_load_excel(final_sr_wo_file, "SR & WO")
        xl_inc   = safe_load_excel(final_inc_file,   "Incident")

        sr_required  = {"Service Request Ageing Days", "Service Request ID", "Service Request Status", "Work Order Assignee Group"}
        inc_required = {"Incident Ageing Days", "Incident ID", "Status", "Assignee Group"}

        sr_sheet_name, df_sr_raw = detect_valid_sheet(xl_sr_wo, sr_required, "Service Request Ageing Days")

        wo_sheet_name, df_wo_raw = None, pd.DataFrame()
        if sr_sheet_name is not None:
            wo_required_inline = {"Service Request Ageing Days", "Work Order ID", "Work Order Status"}
            if _has_required_columns(df_sr_raw, wo_required_inline):
                wo_sheet_name = sr_sheet_name
                df_wo_raw     = df_sr_raw.copy()
            else:
                w_name, w_df = detect_wo_sheet(xl_sr_wo, sr_sheet_name)
                if w_name is not None:
                    wo_sheet_name = w_name
                    df_wo_raw     = w_df.copy()

        inc_sheet_name, df_inc_raw = detect_valid_sheet(xl_inc, inc_required, "Incident Ageing Days")

        status_msg  = "**Data Source:** "
        status_msg += "Live SharePoint Sync\n" if isinstance(final_sr_wo_file, str) else "Manual Upload\n"
        status_msg += f"Detected → SR: `{sr_sheet_name}`, WO: `{wo_sheet_name}`, INC: `{inc_sheet_name}`"
        st.info(status_msg)

        if sr_sheet_name is None:
            st.error("Could not locate a valid Service Request sheet.")
            st.stop()
        if wo_sheet_name is None:
            st.warning("Work Order detail sheet not found — detail tables will be empty.")
            df_wo_raw = pd.DataFrame()
        if inc_sheet_name is None:
            st.error("Could not locate a valid Incident sheet.")
            st.stop()

        # --- SR Metric Calculations ---
        df_sr = df_sr_raw.copy()
        sr_assign_grp_col = find_col(df_sr, "Work Order Assignee Group")
        if sr_assign_grp_col:
            df_sr = df_sr[df_sr[sr_assign_grp_col].astype(str).str.strip().str.upper() == "MYCAREERX SUPPORT"]
        else:
            st.warning(f"⚠️ Filter Skipped: Could not find 'Work Order Assignee Group'. Columns: {', '.join(df_sr.columns.astype(str))}")

        sr_ageing_col = find_col(df_sr, "Service Request Ageing Days")
        sr_status_col = find_col(df_sr, "Service Request Status")
        df_sr[sr_ageing_col] = pd.to_numeric(df_sr[sr_ageing_col], errors='coerce')
        df_sr = df_sr.dropna(subset=[sr_ageing_col])
        df_sr = df_sr[df_sr[sr_status_col].apply(is_active_status)]

        sr_total        = len(df_sr)
        sr_gt_30_count  = int((df_sr[sr_ageing_col] > 30).sum())
        sr_15_30_count  = int(((df_sr[sr_ageing_col] >= 15) & (df_sr[sr_ageing_col] <= 30)).sum())
        sr_1_14_count   = int(((df_sr[sr_ageing_col] >= 1)  & (df_sr[sr_ageing_col] <= 14)).sum())
        sr_gt_1_count   = int((df_sr[sr_ageing_col] > 1).sum())

        sr_gt_1_pct  = round((sr_gt_1_count  / sr_open_wo * 100) if sr_open_wo  > 0 else 0)
        sr_gt_30_pct = round((sr_gt_30_count / sr_gt_1_count * 100) if sr_gt_1_count > 0 else 0)

        # --- SR Details (from WO sheet) ---
        df_wo = df_wo_raw.copy() if not df_wo_raw.empty else pd.DataFrame()
        wo_assign_grp_col = find_col(df_wo, "Work Order Assignee Group")
        if wo_assign_grp_col is not None and not df_wo.empty:
            df_wo = df_wo[df_wo[wo_assign_grp_col].astype(str).str.strip().str.upper() == "MYCAREERX SUPPORT"]

        wo_ageing_col  = find_col(df_wo, "Service Request Ageing Days") if not df_wo.empty else None
        wo_status_col  = find_col(df_wo, "Work Order Status")
        wo_id_col      = find_col(df_wo, "Work Order ID")
        wo_summary_col = find_col(df_wo, "Work Order Summary")
        wo_customer_col = find_col(df_wo, "Customer Full Name (Service Request)") or find_col(df_wo, "Customer Full Name")
        wo_reason_col  = find_col(df_wo, "Work Order Status Reason")
        wo_assignee_col = find_col(df_wo, "Work Order Assignee")

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
                        'SR Ageing':       int(row[wo_ageing_col]),
                        'Work Order No.':  _safe(row, wo_id_col),
                        'Summary':         _safe(row, wo_summary_col),
                        'User/TSG':        _safe(row, wo_customer_col),
                        'WO Status':       _safe(row, wo_status_col),
                        'WO Status Reason': _safe(row, wo_reason_col),
                        'Assignee':        _safe(row, wo_assignee_col),
                    })
                return records

            sr_ageing_gt_30_tickets = extract_wo_records(df_wo[df_wo[wo_ageing_col] > 30])
            sr_ageing_15_30_tickets = extract_wo_records(df_wo[(df_wo[wo_ageing_col] >= 15) & (df_wo[wo_ageing_col] <= 30)])

        # --- INC Metric Calculations ---
        df_inc = df_inc_raw.copy()
        inc_assign_grp_col = find_col(df_inc, "Assignee Group")
        if inc_assign_grp_col:
            df_inc = df_inc[df_inc[inc_assign_grp_col].astype(str).str.strip().str.upper() == "MYCAREERX SUPPORT"]
        else:
            st.warning(f"⚠️ Filter Skipped: Could not find 'Assignee Group'. Columns: {', '.join(df_inc.columns.astype(str))}")

        inc_ageing_col = find_col(df_inc, "Incident Ageing Days")
        inc_status_col = find_col(df_inc, "Status")
        inc_active_col = find_col(df_inc, "Active Incident")

        df_inc[inc_ageing_col] = pd.to_numeric(df_inc[inc_ageing_col], errors='coerce')
        df_inc = df_inc.dropna(subset=[inc_ageing_col])
        if inc_active_col:
            df_inc = df_inc[df_inc[inc_active_col].astype(str).str.strip().str.lower() == "yes"]
        if inc_status_col:
            df_inc = df_inc[df_inc[inc_status_col].apply(is_active_status)]

        inc_total       = len(df_inc)
        inc_gt_90_count = int((df_inc[inc_ageing_col] > 90).sum())
        inc_61_90_count = int(((df_inc[inc_ageing_col] >= 61) & (df_inc[inc_ageing_col] <= 90)).sum())
        inc_31_60_count = int(((df_inc[inc_ageing_col] >= 31) & (df_inc[inc_ageing_col] <= 60)).sum())
        inc_15_30_count = int(((df_inc[inc_ageing_col] >= 15) & (df_inc[inc_ageing_col] <= 30)).sum())
        inc_8_14_count  = int(((df_inc[inc_ageing_col] >= 8)  & (df_inc[inc_ageing_col] <= 14)).sum())
        inc_3_7_count   = int(((df_inc[inc_ageing_col] >= 3)  & (df_inc[inc_ageing_col] <= 7)).sum())
        inc_gt_1_count  = int((df_inc[inc_ageing_col] > 1).sum())

        inc_gt_1_pct = round((inc_gt_1_count / inc_open_input * 100) if inc_open_input > 0 else 0)

        # --- History Tracking ---
        history    = load_history()
        short_date = report_date.strftime("%d-%b-%Y")

        new_record = {
            "date":              short_date,
            "sr_count_gt_30":    sr_gt_30_count,
            "sr_count_15_30":    sr_15_30_count,
            "sr_count_1_14":     sr_1_14_count,
            "inc_count_gt_90":   inc_gt_90_count,
            "inc_count_61_90":   inc_61_90_count,
            "inc_count_31_60":   inc_31_60_count,
            "inc_count_15_30":   inc_15_30_count,
            "inc_count_8_14":    inc_8_14_count,
            "inc_count_3_7":     inc_3_7_count,
        }

        existing_idx = next((i for i, h in enumerate(history) if h.get("date") == short_date), None)

        render_history = list(history)
        if existing_idx is not None:
            render_history[existing_idx] = new_record
        else:
            render_history.append(new_record)

        def parse_date(date_str):
            try:    return datetime.datetime.strptime(date_str, "%d-%b-%Y")
            except: return datetime.datetime.min

        render_history.sort(key=lambda x: parse_date(x.get("date", "")))
        if len(render_history) > 4:
            render_history = render_history[-4:]

        trend_dates    = [h.get("date", "")          for h in render_history]
        sr_trend_gt_30 = [h.get("sr_count_gt_30", 0) for h in render_history]
        sr_trend_15_30 = [h.get("sr_count_15_30", 0) for h in render_history]
        sr_trend_1_14  = [h.get("sr_count_1_14",  0) for h in render_history]
        inc_trend_gt_90 = [h.get("inc_count_gt_90", 0) for h in render_history]
        inc_trend_61_90 = [h.get("inc_count_61_90", 0) for h in render_history]
        inc_trend_31_60 = [h.get("inc_count_31_60", 0) for h in render_history]
        inc_trend_15_30 = [h.get("inc_count_15_30", 0) for h in render_history]
        inc_trend_8_14  = [h.get("inc_count_8_14",  0) for h in render_history]
        inc_trend_3_7   = [h.get("inc_count_3_7",   0) for h in render_history]

        # ========== KPI METRICS CARDS ==========
        st.markdown("<div style='margin-top:-10px;'></div>", unsafe_allow_html=True)
        st.markdown("### Key Metrics Overview")
        st.markdown(
            f"<p style='color:#64748B;margin-top:-15px;'>Report date: "
            f"<b style='color:#00B1A9;'>{report_date_str}</b></p>",
            unsafe_allow_html=True,
        )
        st.markdown("<div style='margin-top:-5px;'></div>", unsafe_allow_html=True)

        m1, m2, m3, m4, m5 = st.columns(5)
        with m1: st.metric("Total SR Tickets (Active)", sr_total)
        with m2: st.metric("SR Ageing > 30d",           sr_gt_30_count)
        with m3: st.metric("SR > 1 day %",              f"{sr_gt_1_pct}%")
        with m4: st.metric("Total INC Tickets (Active)", inc_total)
        with m5: st.metric("INC > 1 day %",             f"{inc_gt_1_pct}%")

        # --- Snapshot Management ---
        st.markdown(" ")
        st.markdown("<p style='color:#64748B;font-size:12px;margin-bottom:5px;font-weight:600;'>SNAPSHOT MANAGEMENT</p>", unsafe_allow_html=True)
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
                st.info(f"✓ {short_date} is already in History. The table below includes it dynamically.", icon="✅")
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
                st.warning(f"Not yet saved. Click Save to log {short_date} into history.", icon="⚠️")

        st.markdown("")

        # ========== GENERATE HTML ==========
        env         = Environment(loader=FileSystemLoader(BASE_DIR))
        template    = env.get_template("template.html")
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
            inc_trend_3_7=inc_trend_3_7,
        )

        email_subject = (
            f"MyCareerX BAU Support Ticket - Ageing Service Request and Incident "
            f"as {report_date.day} {report_date.strftime('%B')}"
        )

        tab_preview, tab_source, tab_export, tab_history = st.tabs(
            ["Email Preview", "HTML Source", "Export Options", "Manage History"]
        )

        with tab_preview:
            st.markdown(f"""
<div style="background-color:#F8FAFC;border:1px solid #E2E8F0;padding:12px 16px;border-radius:8px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,0.05);display:flex;align-items:center;gap:12px;">
    <span style="background-color:#00B1A9;color:white;font-weight:700;font-size:0.75rem;padding:4px 8px;border-radius:4px;text-transform:uppercase;letter-spacing:0.5px;">Subject</span>
    <span style="color:#334155;font-size:0.95rem;font-weight:600;">{email_subject}</span>
</div>
""", unsafe_allow_html=True)
            st.markdown("""<div style="background:#FFFFFF;border:1px solid #E2E8F0;border-radius:14px;padding:8px;box-shadow:0 4px 12px rgba(0,0,0,0.04);">""", unsafe_allow_html=True)
            st.components.v1.html(html_output, height=900, scrolling=True)
            st.markdown("</div>", unsafe_allow_html=True)

        with tab_source:
            st.code(html_output, language="html")

        with tab_export:
            st.markdown("### Export Actions")
            subject_copy_html = f"""
<html><head>
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap');
body{{margin:0;padding:0;font-family:'Inter',sans-serif;}}
.container{{display:flex;align-items:center;background:#F8FAFC;border:1px solid #E2E8F0;border-radius:10px;padding:8px 12px;gap:12px;}}
.text{{flex-grow:1;color:#1E293B;font-size:0.95rem;font-weight:500;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;}}
button{{background:#00B1A9;color:white;border:none;border-radius:6px;padding:6px 14px;font-size:0.8rem;font-weight:600;cursor:pointer;transition:all 0.2s;flex-shrink:0;}}
button:hover{{background:#008C86;transform:translateY(-1px);}}
#msg{{position:absolute;right:80px;color:#00B1A9;font-size:0.75rem;font-weight:700;display:none;}}
</style></head>
<body>
<div style="color:#4A5568;font-size:0.85rem;font-weight:700;margin-bottom:6px;text-transform:uppercase;letter-spacing:0.5px;">Email Subject</div>
<div class="container">
    <div class="text">{email_subject}</div>
    <span id="msg">COPIED!</span>
    <button onclick="copySubject()">COPY</button>
</div>
<script>
function copySubject(){{
    navigator.clipboard.writeText("{email_subject}").then(()=>{{
        const m=document.getElementById("msg");
        m.style.display="inline";
        setTimeout(()=>m.style.display="none",2000);
    }});
}}
</script>
</body></html>"""
            st.components.v1.html(subject_copy_html, height=85)
            st.markdown("<br>", unsafe_allow_html=True)

            exp1, exp2, exp3 = st.columns(3)
            with exp1:
                st.download_button(
                    label="Download .html",
                    data=html_output,
                    file_name=f"Weekly_Report_{report_date.strftime('%Y%m%d')}.html",
                    mime="text/html",
                    use_container_width=True,
                )
            with exp2:
                copy_btn_html = f"""
<html><head>
<style>
body{{margin:0;padding:0;display:flex;flex-direction:column;align-items:center;font-family:sans-serif;}}
button{{background:linear-gradient(135deg,#00B1A9 0%,#008C86 100%);color:white;border:none;border-radius:10px;font-weight:600;padding:0.6rem 1.4rem;font-size:0.9rem;cursor:pointer;width:100%;transition:all 0.3s ease;}}
button:hover{{filter:brightness(1.1);transform:translateY(-1px);}}
#msg{{color:#00B1A9;font-size:0.8rem;font-weight:600;margin-top:5px;display:none;}}
</style></head>
<body>
<button onclick="copyRichText()">Copy Formatted</button>
<div id="msg">Copied to clipboard!</div>
<div id="source" style="display:none;">{base64.b64encode(html_output.encode('utf-8')).decode('utf-8')}</div>
<script>
function copyRichText(){{
    try{{
        const html=decodeURIComponent(escape(window.atob(document.getElementById("source").innerText)));
        navigator.clipboard.write([new ClipboardItem({{"text/html":new Blob([html],{{type:"text/html"}})}})]).then(()=>{{
            const m=document.getElementById("msg");
            m.style.display="block";
            setTimeout(()=>m.style.display="none",3000);
        }}).catch(e=>console.error(e));
    }}catch(e){{console.error(e);}}
}}
</script>
</body></html>"""
                st.components.v1.html(copy_btn_html, height=80)
            with exp3:
                if sys.platform == 'win32':
                    if st.button("Push to Outlook Draft", use_container_width=True):
                        if push_to_outlook(html_output, email_subject):
                            st.success("Draft created in Outlook.")
                else:
                    st.button("Outlook (Windows Only)", use_container_width=True, disabled=True)

        with tab_history:
            st.markdown("### Saved Historical Records")
            saved_data = load_history()
            if not saved_data:
                st.info("No records yet. Process and save a snapshot to see it here.")
            else:
                for idx, h in enumerate(saved_data):
                    st.markdown("<div style='padding:10px;border:1px solid #E2E8F0;border-radius:8px;margin-bottom:8px;'>", unsafe_allow_html=True)
                    col1, col2 = st.columns([5, 1])
                    with col1:
                        st.markdown(
                            f"**{h.get('date')}** &nbsp;|&nbsp; "
                            f"<span style='color:#718096;'>SR &gt;30d: **{h.get('sr_count_gt_30',0)}**</span> &nbsp;|&nbsp; "
                            f"<span style='color:#718096;'>INC &gt;90d: **{h.get('inc_count_gt_90',0)}**</span>",
                            unsafe_allow_html=True,
                        )
                    with col2:
                        if st.button("Delete", key=f"del_{idx}"):
                            saved_data.pop(idx)
                            save_history(saved_data)
                            st.rerun()
                    st.markdown("</div>", unsafe_allow_html=True)

                st.markdown("---")
                if st.button("Clear All History"):
                    save_history([])
                    st.rerun()

    except Exception as e:
        st.error(f"Error processing the files: {e}")
        st.code(traceback.format_exc(), language="text")

else:
    st.markdown("""
<div style="text-align:center;padding:70px 40px;background:linear-gradient(180deg,#FFFFFF 0%,#F0FAFA 100%);border:1px solid #E2E8F0;border-top:4px solid #00A19C;border-radius:16px;">
    <h2 style="color:#1A202C!important;font-weight:800!important;margin:0 0 10px 0!important;">Data Feed Required</h2>
    <p style="color:#718096;max-width:480px;margin:0 auto;line-height:1.6;font-size:0.95rem;">
        Waiting for live SharePoint sync connection, or upload the manual <code>.xlsx</code> exports via the sidebar.
    </p>
</div>
""", unsafe_allow_html=True)