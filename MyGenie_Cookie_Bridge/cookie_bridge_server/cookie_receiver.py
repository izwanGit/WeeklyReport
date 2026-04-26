"""
cookie_receiver.py
------------------
A tiny local HTTP server that:
  1. Receives cookies from the Edge extension (POST /cookies)
  2. Saves them to a local file (cookie_cache.json)
  3. Your Streamlit app reads that file instead of the locked DB

Run this ONCE before opening the Streamlit app:
    python cookie_receiver.py

It runs silently in the background on port 17731.
"""

import json
import os
import sys
import datetime
from http.server import BaseHTTPRequestHandler, HTTPServer

PORT       = 17731
CACHE_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cookie_cache.json")


class CookieHandler(BaseHTTPRequestHandler):

    def log_message(self, format, *args):
        # Suppress default request logs — keep terminal clean
        ts = datetime.datetime.now().strftime("%H:%M:%S")
        print(f"[{ts}] {format % args}")

    def _send_json(self, code, data):
        body = json.dumps(data).encode()
        self.send_response(code)
        self.send_header("Content-Type",                "application/json")
        self.send_header("Content-Length",              str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        # Handle CORS preflight from the extension
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin",  "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")
        self.end_headers()

    def do_GET(self):
        if self.path == "/ping":
            self._send_json(200, {"status": "ok", "message": "Cookie Bridge receiver is running."})
        elif self.path == "/get":
            # Streamlit app calls this to retrieve cached cookies
            if os.path.exists(CACHE_FILE):
                try:
                    with open(CACHE_FILE, "r") as f:
                        data = json.load(f)
                    age_seconds = (
                        datetime.datetime.now()
                        - datetime.datetime.fromisoformat(data.get("saved_at", "2000-01-01"))
                    ).total_seconds()
                    self._send_json(200, {
                        "cookies":    data.get("cookies", {}),
                        "domain":     data.get("domain", ""),
                        "saved_at":   data.get("saved_at", ""),
                        "age_seconds": int(age_seconds),
                    })
                except Exception as e:
                    self._send_json(500, {"error": str(e)})
            else:
                self._send_json(404, {"error": "No cookies cached yet. Click the extension button first."})
        else:
            self._send_json(404, {"error": "Unknown path."})

    def do_POST(self):
        if self.path == "/cookies":
            try:
                length  = int(self.headers.get("Content-Length", 0))
                body    = self.rfile.read(length)
                payload = json.loads(body)

                cookies = payload.get("cookies", {})
                domain  = payload.get("domain", "")

                # Save to disk so Streamlit can read it anytime
                cache = {
                    "cookies":  cookies,
                    "domain":   domain,
                    "saved_at": datetime.datetime.now().isoformat(),
                }
                with open(CACHE_FILE, "w") as f:
                    json.dump(cache, f, indent=2)

                print(f"  ✅ Received & cached {len(cookies)} cookie(s) for {domain}")
                self._send_json(200, {"status": "ok", "count": len(cookies)})

            except Exception as e:
                print(f"  ❌ Error: {e}")
                self._send_json(500, {"error": str(e)})
        else:
            self._send_json(404, {"error": "Unknown path."})


def main():
    print("=" * 52)
    print("  MyGenie Cookie Bridge — Local Receiver")
    print("=" * 52)
    print(f"  Listening on http://localhost:{PORT}")
    print(f"  Cookie cache: {CACHE_FILE}")
    print()
    print("  Steps:")
    print("  1. Keep this window open (minimise is fine)")
    print("  2. Open your Streamlit report app")
    print("  3. Click the extension icon in Edge toolbar")
    print("  4. Click 'Send Session to Report App'")
    print("  5. Watch the counts auto-fill in the sidebar!")
    print()
    print("  Press Ctrl+C to stop.")
    print("-" * 52)

    server = HTTPServer(("localhost", PORT), CookieHandler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\n  Stopped.")


if __name__ == "__main__":
    main()
