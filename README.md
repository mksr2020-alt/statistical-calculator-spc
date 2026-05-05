# 📊 SPC Statistical Process Calculator

A professional **Statistical Process Capability (SPC)** analysis tool built with Python and Streamlit.
Runs entirely in a browser — no installation needed for end users.

---

## ✨ Features

- **Cp / Cpk** capability analysis (Normal & Lognormal distributions)
- **I-MR Control Charts** with all 8 Nelson Rules
- **Hypothesis Testing** (two-sided, upper, lower)
- **AI Predictive Health** — EWMA + linear regression forecast
- **PDF Export** — professional multi-page capability report
- **Excel Export** — multi-sheet branded report with embedded charts
- **History & Trending** — compare runs over time
- **Sigma Assistant** — built-in chatbot for SPC guidance

---

## 🖥️ Setup — Server PC (One Time Only)

> **Requirement:** Python 3.9 or later must be installed.
> Download from: https://www.python.org/downloads/windows/

### Step 1 — Verify Python is installed
Open **Command Prompt** and run:
```cmd
python --version
pip --version
```
Both should show version numbers. If not, install Python first.

### Step 2 — Download this tool
- Click **Code → Download ZIP** on this page
- Extract the ZIP to a folder, e.g.:
  ```
  C:\Tools\spc-tool\
  ```

### Step 3 — Install dependencies (once only)
Open Command Prompt in the extracted folder:
```cmd
cd "C:\Tools\spc-tool"
pip install -r requirements.txt
```
Takes 3–5 minutes. Only needed once.

> **If pip is blocked by corporate proxy, try:**
> ```cmd
> pip install -r requirements.txt --trusted-host pypi.org --trusted-host files.pythonhosted.org
> ```

### Step 4 — Prevent PC from sleeping
```
Settings → System → Power & Sleep
  → Sleep (plugged in) → Never
  → When I close the lid (plugged in) → Do nothing
```

### Step 5 — Start the server
Double-click **`start_server.bat`**

A window appears showing your network URL:
```
╔══════════════════════════════════════════════════╗
║  YOUR URL:  http://10.x.x.x:8501                ║
║  Share this URL with your colleagues.            ║
╚══════════════════════════════════════════════════╝
```
**Keep this window open** while colleagues are using the tool.

---

## 👥 For Colleagues (Users)

**Nothing to install.** Just open Chrome or Edge and go to:
```
http://10.x.x.x:8501
```
*(Get this URL from the person running the server)*

---

## 📦 Dependencies

| Package | Purpose |
|---------|---------|
| `streamlit` | Web UI framework |
| `pandas` | Data handling |
| `numpy` | Numerical computation |
| `scipy` | Statistical distributions |
| `plotly` | Interactive charts |
| `openpyxl` | Excel export |
| `reportlab` | PDF export |
| `kaleido` | Chart-to-PNG for exports |

---

## 📁 Project Files

```
spc-tool/
├── streamlit_spc.py      → Main application
├── requirements.txt      → Python dependencies
├── start_server.bat      → Double-click to start
└── README.md             → This file
```

---

## 🔄 Updating the Tool

1. Download the new ZIP from GitHub
2. Replace `streamlit_spc.py` in your folder
3. Restart `start_server.bat`

No need to reinstall packages unless `requirements.txt` changes.

---

## 🔐 Data Security

- All data stays on your local machine / office LAN
- No internet connection required after setup
- No data is sent to any external service
- No user accounts or registration needed

---

## ❓ Troubleshooting

| Problem | Solution |
|---------|---------|
| Colleagues cannot open URL | Check `start_server.bat` window is still open |
| `pip` command not found | Reinstall Python and tick **"Add to PATH"** |
| Port 8501 blocked | Ask IT to allow port 8501 on local network |
| Packages fail to install | Use the `--trusted-host` flags shown above |
| Tool opens but shows error | Restart `start_server.bat` |
