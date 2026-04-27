import os
import sys
import time
import traceback

# ────────────────────────────────────────────────────────────
# PETRONAS Report Hub — Launcher
# Works in BOTH dev mode (python run_app.py) and frozen .exe
# ────────────────────────────────────────────────────────────

IS_FROZEN = getattr(sys, "frozen", False)

if IS_FROZEN:
    BASE_DIR = sys._MEIPASS          # read-only bundle dir with all assets
    EXE_DIR  = os.path.dirname(sys.executable)  # writable dir next to the .exe
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXE_DIR  = BASE_DIR

# Set working directory so Streamlit finds templates relative to it
os.chdir(BASE_DIR)

LOG_FILE = os.path.join(EXE_DIR, "petronas_hub_debug.log")

def log(msg):
    line = f"[{time.strftime('%H:%M:%S')}] {msg}"
    print(line, flush=True)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def main():
    try:
        import socket
        import subprocess
        import webbrowser
        from threading import Thread

        log("=" * 55)
        log("PETRONAS Report Hub — LAUNCHER STARTING")
        log(f"  IS_FROZEN : {IS_FROZEN}")
        log(f"  BASE_DIR  : {BASE_DIR}")
        log(f"  EXE_DIR   : {EXE_DIR}")
        log(f"  CWD       : {os.getcwd()}")
        log(f"  Python    : {sys.version}")

        # ── 1. Pre-flight checks ──────────────────────────────
        critical = ["Report_Hub.py", "template.html", "PETRONAS_LOGO_SQUARE.png"]
        for fname in critical:
            path = os.path.join(BASE_DIR, fname)
            if not os.path.exists(path):
                log(f"CRITICAL MISSING FILE: {fname}")
                print(f"\n{'!'*50}\nERROR: Missing file: {fname}\n{'!'*50}\n")
                input("Press Enter to close...")
                return

        PORT = 8501
        URL  = f"http://127.0.0.1:{PORT}/?splash=true"

        # ── 2. Browser launcher (runs in background thread) ──
        def open_browser():
            # Poll until server is ready (up to 60 seconds)
            log("Waiting for server to be ready...")
            for _ in range(120):
                try:
                    with socket.create_connection(("127.0.0.1", PORT), timeout=0.5):
                        break
                except OSError:
                    time.sleep(0.5)
            else:
                log("TIMEOUT: Server never became ready.")
                return

            log(f"Server is up. Opening browser at {URL}")

            # Try app-mode (chromeless window)
            browser_cmd = None
            if sys.platform == "win32":
                candidates = [
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
                ]
                for p in candidates:
                    if os.path.exists(p):
                        browser_cmd = [p, f"--app={URL}", "--new-window"]
                        break
            elif sys.platform == "darwin":
                candidates = [
                    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
                    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
                ]
                for p in candidates:
                    if os.path.exists(p):
                        browser_cmd = [p, f"--app={URL}"]
                        break

            if browser_cmd:
                log(f"Launching app-mode browser: {browser_cmd[0]}")
                try:
                    subprocess.Popen(browser_cmd)
                    return
                except Exception as e:
                    log(f"App-mode launch failed ({e}), falling back to default browser")

            log("Using default system browser.")
            webbrowser.open(URL)

        browser_thread = Thread(target=open_browser, daemon=True)
        browser_thread.start()

        # ── 3. Start Streamlit server ─────────────────────────
        # IMPORTANT: In a frozen .exe, sys.executable IS the .exe — we cannot
        # run it with "-m streamlit".  We must call stcli.main() directly.
        # stcli.main() blocks forever (it runs the Tornado event loop).
        hub_script = os.path.join(BASE_DIR, "Report_Hub.py")

        if IS_FROZEN:
            log("Frozen mode: starting Streamlit via stcli.main() in main thread")
            import streamlit.web.cli as stcli
            sys.argv = [
                "streamlit", "run", hub_script,
                "--server.port", str(PORT),
                "--server.headless", "true",
                "--browser.gatherUsageStats", "false",
                "--global.developmentMode", "false",
                "--server.fileWatcherType", "none",  # no file-watcher in frozen env
            ]
            # stcli.main() never returns — that is correct behaviour
            stcli.main()

        else:
            # Dev mode: spawn a subprocess so we can keep the launcher alive
            log("Dev mode: starting Streamlit as subprocess")
            cmd = [
                sys.executable, "-m", "streamlit", "run", hub_script,
                "--server.port", str(PORT),
                "--server.headless", "true",
                "--browser.gatherUsageStats", "false",
                "--global.developmentMode", "false",
            ]
            log(f"Command: {' '.join(cmd)}")
            server_proc = subprocess.Popen(cmd, cwd=BASE_DIR)

            # Wait for the browser thread to finish, then keep alive
            browser_thread.join()

            # Stay alive until the server exits
            try:
                server_proc.wait()
            except KeyboardInterrupt:
                log("Keyboard interrupt — shutting down.")
                server_proc.terminate()

        log("Application exited.")
        os._exit(0)

    except BaseException as e:
        print("\n" + "!"*50)
        print("CRITICAL LAUNCHER ERROR:")
        traceback.print_exc()
        print("!"*50 + "\n")
        try:
            # Also write to log
            log(f"CRITICAL LAUNCHER ERROR:\n{traceback.format_exc()}")
        except Exception:
            pass
        input("Press Enter to close...")


if __name__ == "__main__":
    main()
