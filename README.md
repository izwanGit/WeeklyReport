# 📊 Weekly SR & Incident Report Generator

This repository contains an automated desktop tool built with **Streamlit** and **Jinja2** that effortlessly converts raw Excel data exported from MyGenie into a cleanly formatted HTML email draft tailored with the Petronas UI aesthetic.

## ✨ Key Features
- **Browser-based UI**: Simple local web interface using Streamlit, requiring zero complex configuration for the end user.
- **Automated Metric Calculations**: Instantly processes Ageing > 30 days, percentages, and status summaries from raw Excel sheets. 
- **Historical Tracker**: Maintains a 4-week revolving historical snapshot stored locally in `history.json` to ensure tracking across emails.
- **Outlook Integration**: Push HTML logs directly into an Outlook draft via `win32com` (Windows exclusively).

## 🗂 File Structure
| File Target | Purpose |
| ----------- | ------- |
| `app.py` | The main Streamlit logic. Contains data ingestion, DataFrame normalizations, and UI bindings. |
| `template.html` | The Jinja2 HTML layout file. This is the exact formatting the eventual email takes. |
| `run_app.py` | The internal launch handler useful for compiling via PyInstaller. |
| `history.json` | An auto-created file on the first run tracking table trends over the past 4 weeks! |

## 🛠 Modifying the Email Format
If the email standard changes in the future, modifications are straightforward:
- Open `template.html`.
- It relies on Jinja2, meaning variables are formatted inside braces like `{{ this }}` and iterations are like `{% for items in list %}`.
- All styles are defined **inline** per `<tr>` row to maximize exact layout behavior within Microsoft Outlook.

## 🚀 Running in Dev Mode (Mac or Windows)
To run or test locally via terminal:
```bash
# 1. Setup a virtual environment
python3 -m venv .venv

# 2. Activate the virtual environment
source .venv/bin/activate  # On Mac/Linux
# or .venv\Scripts\activate on Windows

# 3. Install requirements
pip install -r requirements.txt

# 4. Start the Application
streamlit run app.py
```

## 📦 Building the Standalone `.exe` (For Windows Deployments)
Ensure you are operating on a **Windows machine** when compiling the executable for deployment to other Windows staff members.

1. Ensure Python dependencies are installed (`pip install -r requirements.txt`).
2. Run the provided PyInstaller script specification:
   ```bash
   pyinstaller app.spec
   ```
3. Locate the completed standalone executable in the `/dist` directory!
   > ⚠️ Make sure to bundle `template.html` in the exact same directory alongside your `.exe` file!
