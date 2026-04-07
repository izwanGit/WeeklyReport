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

def open_browser():
    webbrowser.open_new("http://localhost:8501")

if __name__ == "__main__":
    # Wait 5 seconds after start then open browser automatically
    Timer(5, open_browser).start()
    
    # Run streamlit CLI
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("app.py"),
        "--global.developmentMode=false",
        "--server.port=8501",
        "--server.headless=true"
    ]
    sys.exit(stcli.main())
