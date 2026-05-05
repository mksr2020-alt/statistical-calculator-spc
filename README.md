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

## 🖥️ Running on Windows (No Admin Rights Needed)

### Step 1 — Download WinPython (one time)
> WinPython is a portable Python — no installation, no admin rights required.

1. Go to **https://winpython.github.io/**
2. Download: `WinPython64-3.11.x.x.exe` (dot version, ~100 MB)
3. Run the `.exe` — it just extracts a folder (no install wizard)

### Step 2 — Install dependencies
1. Open `WinPython Command Prompt.exe` (inside the WinPython folder)
2. Navigate to this project folder:
   ```cmd
   cd "C:\path\to\spc-clean-share"
   ```
3. Install packages:
   ```cmd
   pip install -r requirements.txt
   ```
   > First time only. Takes 2–5 minutes.

### Step 3 — Start the server
Double-click **`start_server.bat`**  
OR run in the WinPython Command Prompt:
```cmd
streamlit run streamlit_spc.py --server.address=0.0.0.0 --server.port=8501
```

### Step 4 — Share with colleagues
- Your URL: `http://YOUR-PC-IP:8501`
- Find your IP: open Command Prompt → type `ipconfig` → look for **IPv4 Address**
- Colleagues open this URL in Chrome/Edge — no install needed on their side

---

## 📦 Requirements

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
spc-clean-share/
├── streamlit_spc.py      → Main application (all logic + UI)
├── requirements.txt      → Python dependencies
├── start_server.bat      → Windows one-click server start
└── README.md             → This file
```

---

## 🔐 Data Security

- All data stays on your local machine / office LAN
- No internet connection required to run
- No data is sent to any cloud service
- No user accounts or registration needed

---

## 📞 Support

For technical issues, contact the tool owner directly.
