"""
PETRONAS Report Hub — Desktop Launcher
=======================================
Starts Streamlit server in background, waits for it to be ready,
then opens a chromeless browser window (Edge/Chrome App Mode).
Writes a log file next to the .exe for debugging any issues.
"""

import os
import sys
import time
import subprocess
import socket
import traceback
import webbrowser
from threading import Thread

# ────────────────────────────────────────────────────────────
# 0. Resolve frozen vs. dev paths & set working directory
# ────────────────────────────────────────────────────────────
IS_FROZEN = getattr(sys, "frozen", False)

if IS_FROZEN:
    # When bundled by PyInstaller, all data files live in _MEIPASS
    BASE_DIR = sys._MEIPASS
    # The log file should be next to the .exe so the user can find it
    EXE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR = BASE_DIR

# CRITICAL: Streamlit looks for .streamlit/config.toml relative to CWD
os.chdir(BASE_DIR)

LOG_FILE = os.path.join(EXE_DIR, "petronas_hub_debug.log")

# ────────────────────────────────────────────────────────────
# 1. Logger — writes to file AND console
# ────────────────────────────────────────────────────────────
def log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def resolve_path(path):
    return os.path.join(BASE_DIR, path)


# ────────────────────────────────────────────────────────────
# 2. Pre-flight checks  — catch problems BEFORE starting
# ────────────────────────────────────────────────────────────
def preflight():
    """Verify every critical file is reachable."""
    critical_files = [
        "Report_Hub.py",
        "template.html",
        "PETRONAS_LOGO_SQUARE.png",
        "PETRONAS_LOGO_HORIZONTAL.svg",
        "PETRONAS_LOGO_HORIZONTAL_WHITE.svg",
    ]
    critical_dirs = ["pages"]
    ok = True
    for f in critical_files:
        p = resolve_path(f)
        exists = os.path.isfile(p)
        log(f"  [{'OK' if exists else 'MISSING'}] {p}")
        if not exists:
            ok = False
    for d in critical_dirs:
        p = resolve_path(d)
        exists = os.path.isdir(p)
        log(f"  [{'OK' if exists else 'MISSING'}] {p}/")
        if not exists:
            ok = False
    return ok


# ────────────────────────────────────────────────────────────
# 3. Streamlit server — runs in daemon thread
# ────────────────────────────────────────────────────────────
_server_error = None  # will be set if the server crashes


def start_server():
    global _server_error
    try:
        log("Importing streamlit.web.cli ...")
        import streamlit.web.cli as stcli

        sys.argv = [
            "streamlit",
            "run",
            resolve_path("Report_Hub.py"),
            "--global.developmentMode=false",
            "--server.port=8501",
            "--server.headless=true",
            "--server.address=localhost",
            "--browser.gatherUsageStats=false",
            "--browser.serverAddress=localhost",
        ]
        log(f"sys.argv = {sys.argv}")
        log("Calling stcli.main() ...")
        stcli.main()
    except SystemExit:
        # Streamlit sometimes does sys.exit(0) on clean shutdown — ignore
        log("Server exited via SystemExit (normal).")
    except Exception as exc:
        _server_error = traceback.format_exc()
        log(f"SERVER CRASHED:\n{_server_error}")


# ────────────────────────────────────────────────────────────
# 4. Wait for server
# ────────────────────────────────────────────────────────────
def wait_for_server(port=8501, timeout=45):
    """Poll localhost:port until it accepts a TCP connection."""
    start = time.time()
    while time.time() - start < timeout:
        # If the server thread already crashed, bail early
        if _server_error is not None:
            return False
        try:
            with socket.create_connection(("localhost", port), timeout=1):
                return True
        except OSError:
            time.sleep(0.5)
    return False


# ────────────────────────────────────────────────────────────
# 5. Browser launcher
# ────────────────────────────────────────────────────────────
def get_browser_command(url):
    """Returns a subprocess arg list for Edge/Chrome in --app mode."""
    if sys.platform == "win32":
        candidates = [
            r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
            r"C:\Program Files\Google\Chrome\Application\chrome.exe",
            r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
            os.path.expandvars(r"%LocalAppData%\Google\Chrome\Application\chrome.exe"),
        ]
        for path in candidates:
            if os.path.exists(path):
                return [path, f"--app={url}"]
    elif sys.platform == "darwin":
        return ["open", "-W", "-n", "-a", "Google Chrome", "--args", f"--app={url}"]
    return None


# ────────────────────────────────────────────────────────────
# 6. Main entry point
# ────────────────────────────────────────────────────────────
if __name__ == "__main__":
    log("=" * 55)
    log("PETRONAS Report Hub — launcher starting")
    log(f"  frozen   = {IS_FROZEN}")
    log(f"  exe      = {sys.executable}")
    log(f"  BASE_DIR = {BASE_DIR}")
    log(f"  EXE_DIR  = {EXE_DIR}")
    log(f"  CWD      = {os.getcwd()}")
    log(f"  Python   = {sys.version}")
    log("")

    # Pre-flight
    log("Pre-flight checks:")
    if not preflight():
        log("FATAL: One or more critical files are missing. Cannot continue.")
        input("Press Enter to exit ...")
        sys.exit(1)
    log("")

    # Start server
    log("Starting Streamlit server thread ...")
    server_thread = Thread(target=start_server, daemon=True)
    server_thread.start()

    # Wait
    log("Waiting for server on port 8501 (up to 45 s) ...")
    ready = wait_for_server(port=8501, timeout=45)

    if not ready:
        log("FATAL: Server did not start in 45 seconds.")
        if _server_error:
            log(f"Server error:\n{_server_error}")
        else:
            log("No crash was captured — server may still be loading.")
            log("Check if port 8501 is blocked by your firewall / antivirus.")
        input("Press Enter to exit ...")
        sys.exit(1)

    log("Server is LIVE! Opening browser ...")

    # Launch browser
    url = "http://localhost:8501/?splash=true"
    cmd = get_browser_command(url)

    if cmd:
        log(f"Browser command: {cmd}")
        process = subprocess.Popen(cmd)
        process.wait()  # blocks until user closes the window
    else:
        log("No Chrome/Edge found — using default browser.")
        webbrowser.open_new(url)
        while True:
            time.sleep(10)

    # Cleanup
    log("Browser closed. Shutting down.")
    os._exit(0)
