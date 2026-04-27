import re

file_path = "pages/1_Weekly_Report.py"
with open(file_path, "r", encoding="utf-8") as f:
    content = f.read()

# Add the petronas_alert function right after the imports
imports_end = content.find("import streamlit as st")
if imports_end != -1:
    imports_end = content.find("\n", imports_end) + 1
    
    helper_func = """
def petronas_alert(message: str, type: str = "info", icon: str = None):
    # PETRONAS Brand Colors
    colors = {
        "success": ("rgb(191,215,48)", "rgba(191,215,48,0.15)"), # Lime Green
        "info": ("rgb(0,177,169)", "rgba(0,177,169,0.15)"),       # Teal
        "warning": ("rgb(253,185,36)", "rgba(253,185,36,0.15)"),  # Yellow
        "error": ("rgb(118,63,152)", "rgba(118,63,152,0.15)"),    # Purple
        "blue": ("rgb(32,65,154)", "rgba(32,65,154,0.15)")        # Blue
    }
    border_color, bg_color = colors.get(type, colors["info"])
    icon_html = f"<span style='margin-right: 8px; font-size: 1.1em;'>{icon}</span>" if icon else ""
    html = f'''<div style="background-color: {bg_color}; border-left: 4px solid {border_color}; padding: 12px 16px; border-radius: 4px; margin-bottom: 16px; font-family: sans-serif; color: #1E293B; display: flex; align-items: center;">{icon_html}<div>{message}</div></div>'''
    st.markdown(html, unsafe_allow_html=True)

"""
    content = content[:imports_end] + helper_func + content[imports_end:]

# Replace st.success("...") with petronas_alert("...", type="success", icon="✅")
content = re.sub(r'st\.success\((.*?)\)', r'petronas_alert(\1, type="success", icon="✅")', content)

# Replace st.info("...") with petronas_alert("...", type="info", icon="ℹ️")
# Careful with the one that already has icon="✅"
content = re.sub(r'st\.info\((.*?), icon="✅"\)', r'petronas_alert(\1, type="info", icon="✅")', content)
content = re.sub(r'st\.info\((?!.*?icon=)(.*?)\)', r'petronas_alert(\1, type="info", icon="ℹ️")', content)

# Replace st.error("...") with petronas_alert("...", type="error", icon="🚨")
content = re.sub(r'st\.error\((.*?)\)', r'petronas_alert(\1, type="error", icon="🚨")', content)

with open(file_path, "w", encoding="utf-8") as f:
    f.write(content)

print("Replaced all default Streamlit alerts with PETRONAS alerts.")
