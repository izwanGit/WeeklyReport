"""
REPLACEMENT for get_browser_cookies() in your Streamlit app.
------------------------------------------------------------
Replace your entire get_browser_cookies() function with this one.
It tries the local Cookie Bridge receiver first (the extension method),
then falls back to the old browser-cookie3 approach.

Also add this helper anywhere before the function:
    COOKIE_BRIDGE_URL = "http://localhost:17731"
"""

COOKIE_BRIDGE_URL = "http://localhost:17731"   # ← add this near your other constants


def get_browser_cookies() -> dict:
    """
    Priority order:
      1. Local Cookie Bridge receiver (extension → localhost server) ← NEW
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
                    f"✅ Cookie Bridge: {len(cookies)} cookie(s) received via extension\n"
                    f"   Saved: {saved_at}"
                )
                return cookies
            else:
                diag_lines.append("⚠️ Cookie Bridge: server running but no cookies cached yet.")
        elif resp.status_code == 404:
            diag_lines.append(
                "ℹ️ Cookie Bridge: server running but no cookies yet — "
                "click the extension button in Edge toolbar."
            )
        else:
            diag_lines.append(f"⚠️ Cookie Bridge: unexpected status {resp.status_code}")

    except requests.exceptions.ConnectionError:
        diag_lines.append(
            "ℹ️ Cookie Bridge not running. "
            "Start cookie_receiver.py for the best experience."
        )
    except Exception as e:
        diag_lines.append(f"⚠️ Cookie Bridge error: {e}")

    # ----------------------------------------------------------------
    # METHOD 2: browser-cookie3 (original fallback)
    # ----------------------------------------------------------------
    try:
        import browser_cookie3
        diag_lines.append(f"🔄 Trying browser-cookie3 fallback…")

        loaders = [
            ("Edge",   browser_cookie3.edge),
            ("Chrome", browser_cookie3.chrome),
        ]
        for name, loader in loaders:
            try:
                cj      = loader(domain_name=MYGENIE_DOMAIN)
                cookies = {c.name: c.value for c in cj}
                if cookies:
                    diag_lines.append(f"✅ {name}: got {len(cookies)} cookie(s) via browser-cookie3")
                    st.session_state['_cookie_source'] = f"browser-cookie3 ({name})"
                    st.session_state['_cookie_diag']   = "\n".join(diag_lines)
                    return cookies
                else:
                    diag_lines.append(f"ℹ️ {name}: 0 cookies found for domain")
            except PermissionError as e:
                diag_lines.append(f"⚠️ {name} locked (browser running): {e}")
            except Exception as e:
                diag_lines.append(f"⚠️ {name}: {type(e).__name__}: {e}")

    except ImportError:
        diag_lines.append("⚠️ browser-cookie3 not installed (pip install browser-cookie3)")

    # ----------------------------------------------------------------
    # All methods failed
    # ----------------------------------------------------------------
    st.session_state['_cookie_source'] = None
    st.session_state['_cookie_diag']   = "\n".join(diag_lines)
    return {}
