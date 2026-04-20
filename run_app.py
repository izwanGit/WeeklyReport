import os
import sys
import time
import traceback

# ────────────────────────────────────────────────────────────
# ULTRA ROBUST ERROR WRAPPER
# ────────────────────────────────────────────────────────────
def main():
    try:
        import subprocess
        import socket
        import webbrowser
        from threading import Thread

        # Resolve paths
        IS_FROZEN = getattr(sys, "frozen", False)
        if IS_FROZEN:
            BASE_DIR = sys._MEIPASS
            EXE_DIR = os.path.dirname(sys.executable)
        else:
            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            EXE_DIR = BASE_DIR

        # Set working directory to where all our assets are
        os.chdir(BASE_DIR)
        
        LOG_FILE = os.path.join(EXE_DIR, "petronas_hub_debug.log")

        def log(msg):
            line = f"[{time.strftime('%H:%M:%S')}] {msg}"
            print(line, flush=True)
            try:
                with open(LOG_FILE, "a", encoding="utf-8") as f:
                    f.write(line + "\n")
            except:
                pass

        log("=" * 55)
        log("PETRONAS Report Hub — LAUNCHER STARTING")
        log(f"  BASE_DIR: {BASE_DIR}")
        log(f"  EXE_DIR: {EXE_DIR}")
        log(f"  CWD: {os.getcwd()}")
        log(f"  Python: {sys.version}")

        # ────────────────────────────────────────────────────────────
        # 1. Pre-flight checks
        # ────────────────────────────────────────────────────────────
        def check_files():
            critical = ["Report_Hub.py", "template.html", "PETRONAS_LOGO_SQUARE.png"]
            for f in critical:
                path = os.path.join(BASE_DIR, f)
                if not os.path.exists(path):
                    log(f"CRITICAL MISSING FILE: {f} at {path}")
                    return False
            return True

        if not check_files():
            log("FATAL: Critical assets missing from bundle.")
            print("\n" + "!"*50)
            print("ERROR: THE APPLICATION IS MISSING CORE FILES.")
            print("Please ensure you extracted ALL files from the ZIP.")
            print("!"*50 + "\n")
            input("Press Enter to close...")
            return

        # ────────────────────────────────────────────────────────────
        # 2. Server Thread
        # ────────────────────────────────────────────────────────────
        _server_crash = None

        def start_server():
            nonlocal _server_crash
            try:
                log("Loading Streamlit components...")
                import streamlit.web.cli as stcli
                
                # We must use absolute path for the script
                hub_script = os.path.join(BASE_DIR, "Report_Hub.py")
                
                sys.argv = [
                    "streamlit", "run", hub_script,
                    "--global.developmentMode=false",
                    "--server.port=8501",
                    "--server.headless=true",
                    "--browser.gatherUsageStats=false"
                ]
                log(f"Starting server: {hub_script}")
                stcli.main()
            except Exception as e:
                _server_crash = traceback.format_exc()
                log(f"SERVER THREAD CRASHED: {e}")

        server_thread = Thread(target=start_server, daemon=True)
        server_thread.start()

        # ────────────────────────────────────────────────────────────
        # 3. Dynamic Polling (Wait for port 8501)
        # ────────────────────────────────────────────────────────────
        log("Initializing local server (please wait)...")
        ready = False
        for i in range(60): # 30 seconds timeout
            if _server_crash:
                break
            try:
                with socket.create_connection(("127.0.0.1", 8501), timeout=0.5):
                    ready = True
                    break
            except:
                time.sleep(0.5)

        if not ready:
            log("FATAL: Server failed to start.")
            if _server_crash:
                print("\n" + "="*50)
                print("SERVER ERROR DETAILS:")
                print(_server_crash)
                print("="*50 + "\n")
            else:
                log("Server timed out without crash report. Check your firewall.")
            input("Press Enter to close...")
            return

        # ────────────────────────────────────────────────────────────
        # 4. Launch Browser
        # ────────────────────────────────────────────────────────────
        url = "http://127.0.0.1:8501/?splash=true"
        log(f"Server ready at {url}. Launching browser...")

        def get_browser():
            if sys.platform == "win32":
                paths = [
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                ]
                for p in paths:
                    if os.path.exists(p):
                        return [p, f"--app={url}"]
            return None

        cmd = get_browser()
        if cmd:
            log(f"Using App Mode: {cmd[0]}")
            try:
                proc = subprocess.Popen(cmd)
                proc.wait() # KEEP CONSOLE OPEN UNTIL BROWSER CLOSES
            except Exception as e:
                log(f"Browser launch failed: {e}")
                webbrowser.open(url)
                time.sleep(10)
        else:
            log("No Edge/Chrome found. Using default browser.")
            webbrowser.open(url)
            # Stay alive so the server doesn't kill
            while True:
                time.sleep(10)

        log("Application closed by user.")
        os._exit(0)

    except Exception:
        print("\n" + "!"*50)
        print("CRITICAL LAUNCHER ERROR:")
        traceback.print_exc()
        print("!"*50 + "\n")
        input("Press Enter to close...")

if __name__ == "__main__":
    main()
