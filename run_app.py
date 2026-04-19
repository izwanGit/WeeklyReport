import os
import sys
import streamlit.web.cli as stcli
import webbrowser
from threading import Timer
import streamlit.runtime.scriptrunner.magic_funcs


def resolve_path(path):
    # Handle pathing for PyInstaller 'onedir' or 'onefile' mode
    if getattr(sys, 'frozen', False):
        curr_dir = sys._MEIPASS
    else:
        curr_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(curr_dir, path)

import subprocess

def open_app_mode(url):
    try:
        if sys.platform == "win32":
            edge_paths = [
                r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
            ]
            for edge in edge_paths:
                if os.path.exists(edge):
                    subprocess.Popen([edge, f"--app={url}"])
                    return True
            
            chrome_paths = [
                r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
            ]
            for chrome in chrome_paths:
                if os.path.exists(chrome):
                    subprocess.Popen([chrome, f"--app={url}"])
                    return True
        elif sys.platform == "darwin":
            try:
                subprocess.Popen(["open", "-n", "-a", "Google Chrome", "--args", f"--app={url}"])
                return True
            except:
                pass
        return False
    except:
        return False

def open_browser():
    url = "http://localhost:8501"
    if not open_app_mode(url):
        webbrowser.open_new(url)

if __name__ == "__main__":
    # Wait 5 seconds after start then open browser automatically
    Timer(5, open_browser).start()
    
    # Run streamlit CLI
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("Report_Hub.py"),
        "--global.developmentMode=false",
        "--server.port=8501",
        "--server.headless=true"
    ]
    sys.exit(stcli.main())
