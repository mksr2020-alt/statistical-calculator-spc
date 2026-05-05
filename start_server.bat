@echo off
REM ══════════════════════════════════════════════════════════════════
REM  SPC Statistical Process Calculator — Server Start Script
REM  Double-click this file to start. Keep the window open.
REM  Requires Python 3.9+ to be installed.
REM ══════════════════════════════════════════════════════════════════

title SPC Calculator — Server

REM ── Move to the folder where this .bat file lives ─────────────────
cd /d "%~dp0"

REM ── Check Python is available ─────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  ERROR: Python not found.
    echo  Install Python from https://www.python.org/downloads/windows/
    echo  Make sure to tick "Add Python to PATH" during install.
    echo.
    pause
    exit /b 1
)

REM ── Check Streamlit is installed ──────────────────────────────────
python -m streamlit --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Streamlit not found. Installing dependencies...
    echo.
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo  Install failed. Try running:
        echo  pip install -r requirements.txt --trusted-host pypi.org --trusted-host files.pythonhosted.org
        echo.
        pause
        exit /b 1
    )
)

REM ── Get local IP address automatically ───────────────────────────
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /C:"IPv4"') do (
    set MY_IP=%%a
    goto :found_ip
)
:found_ip
set MY_IP=%MY_IP: =%

REM ── Display startup banner ────────────────────────────────────────
echo.
echo  ╔══════════════════════════════════════════════════════════╗
echo  ║         SPC Statistical Process Calculator               ║
echo  ╠══════════════════════════════════════════════════════════╣
echo  ║                                                          ║
echo  ║  Status:    Starting...                                  ║
echo  ║                                                          ║
echo  ║  Your URL:  http://%MY_IP%:8501
echo  ║  Local:     http://localhost:8501                        ║
echo  ║                                                          ║
echo  ║  Share "Your URL" with colleagues on the same network.  ║
echo  ║  They open it in Chrome or Edge — nothing to install.   ║
echo  ║                                                          ║
echo  ║  Press Ctrl+C to stop the server.                       ║
echo  ╚══════════════════════════════════════════════════════════╝
echo.

REM ── Start Streamlit ───────────────────────────────────────────────
python -m streamlit run streamlit_spc.py ^
    --server.address=0.0.0.0 ^
    --server.port=8501 ^
    --server.headless=true ^
    --browser.gatherUsageStats=false ^
    --server.enableCORS=false ^
    --server.enableXsrfProtection=true

REM ── If server stops ───────────────────────────────────────────────
echo.
echo  Server stopped. Press any key to close.
pause
