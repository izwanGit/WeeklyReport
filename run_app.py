import os
import sys
import streamlit.web.cli as stcli
import webbrowser
from threading import Thread
import time
import subprocess
import signal

def resolve_path(path):
    if getattr(sys, 'frozen', False):
        curr_dir = sys._MEIPASS
    else:
        curr_dir = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(curr_dir, path)

def start_server():
    sys.argv = [
        "streamlit",
        "run",
        resolve_path("Report_Hub.py"),
        "--global.developmentMode=false",
        "--server.port=8501",
        "--server.headless=true",
        "--browser.gatherUsageStats=false"
    ]
    stcli.main()

def get_browser_command(url):
    """Returns a list for subprocess.Popen that opens url in app mode."""
    if sys.platform == "win32":
        edge_paths = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
        ]
        for edge in edge_paths:
            if os.path.exists(edge):
                return [edge, f"--app={url}"]
        
        chrome_paths = [
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe")
        ]
        for chrome in chrome_paths:
            if os.path.exists(chrome):
                return [chrome, f"--app={url}"]
    
    elif sys.platform == "darwin":
        # On Mac, we can use 'open' with application name. 
        # But 'open -W' waits for the app to exit.
        return ["open", "-W", "-n", "-a", "Google Chrome", "--args", f"--app={url}"]
    
    return None

if __name__ == "__main__":
    # 1. Start Streamlit in a background thread
    server_thread = Thread(target=start_server, daemon=True)
    server_thread.start()

    # 2. Wait a moment for server to initialize
    time.sleep(3)

    # 3. Launch browser in App Mode (blocks until closed)
    url = "http://localhost:8501/?splash=true"
    cmd = get_browser_command(url)
    
    if cmd:
        process = subprocess.Popen(cmd)
        process.wait() # This blocks the main script until you close the browser window
    else:
        # Fallback to default browser if no Chrome/Edge (non-blocking, keeps process alive)
        webbrowser.open_new(url)
        while True:
            time.sleep(10) # Keep script alive if we can't wait on browser process

    # 4. Once process.wait() returns, the user closed the window. Exit everything.
    os._exit(0)
