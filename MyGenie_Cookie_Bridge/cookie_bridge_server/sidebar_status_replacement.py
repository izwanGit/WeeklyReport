"""
SIDEBAR STATUS SECTION REPLACEMENT
-----------------------------------
In your Streamlit app, find this block in the sidebar:

    if cookie_ok and (auto_wo is not None or auto_inc is not None):
        st.caption("🔑 Browser session detected — counts auto-filled.")
    elif cookie_ok:
        st.caption("⚠️ Session found but API returned no data. Enter counts manually.")
    else:
        st.caption("⚠️ Could not read browser cookies. Enter counts manually.")
        diag = st.session_state.get('_cookie_diag', '')
        if diag:
            with st.expander("🔍 Cookie Diagnostics", expanded=True):
                st.code(diag, language="text")

Replace it with:
"""

    cookie_source = st.session_state.get('_cookie_source')

    if cookie_ok and (auto_wo is not None or auto_inc is not None):
        if cookie_source and "extension" in cookie_source:
            st.success(f"🔗 Connected via Extension — counts auto-filled.")
        else:
            st.caption(f"🔑 Browser session detected — counts auto-filled.")

    elif cookie_ok:
        st.caption("⚠️ Session found but API returned no data. Enter counts manually.")

    else:
        # Check if the bridge server is running but just needs the button click
        bridge_needs_click = st.session_state.get('_cookie_diag', '').startswith('ℹ️ Cookie Bridge: server running')
        bridge_not_running = 'Cookie Bridge not running' in st.session_state.get('_cookie_diag', '')

        if bridge_needs_click:
            st.info(
                "🔗 Cookie Bridge is running!\n\n"
                "Click the **teal extension icon** in your Edge toolbar, "
                "then click **'Send Session to Report App'**.",
                icon="👆"
            )
        elif bridge_not_running:
            st.warning("⚠️ Could not auto-read cookies.")
            with st.expander("How to fix", expanded=True):
                st.markdown(
                    "**Option A (Recommended):**\n"
                    "Run `cookie_receiver.py` then click the Edge extension.\n\n"
                    "**Option B:** Enter counts manually below."
                )
        else:
            st.caption("⚠️ Could not read browser cookies. Enter counts manually.")

        diag = st.session_state.get('_cookie_diag', '')
        if diag:
            with st.expander("🔍 Diagnostics", expanded=False):
                st.code(diag, language="text")
