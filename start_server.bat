@echo off
REM ══════════════════════════════════════════════════════════════════
REM  SPC Statistical Process Calculator — Server Start Script
REM  Double-click this file to start the server.
REM  Works with WinPython (no admin rights needed).
REM ══════════════════════════════════════════════════════════════════

title SPC Calculator — Running on port 8501

REM ── Change this to your actual project folder path ────────────────
set APP_DIR=%~dp0
cd /d "%APP_DIR%"

REM ── Find your IP address automatically ───────────────────────────
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /C:"IPv4"') do (
    set MY_IP=%%a
    goto :found_ip
)
:found_ip
set MY_IP=%MY_IP: =%

echo.
echo  ╔══════════════════════════════════════════════════════════╗
echo  ║         SPC Statistical Process Calculator               ║
echo  ╠══════════════════════════════════════════════════════════╣
echo  ║                                                           ║
echo  ║  Server starting...                                       ║
echo  ║                                                           ║
echo  ║  YOUR URL:  http://%MY_IP%:8501              ║
echo  ║  LOCAL:     http://localhost:8501                         ║
echo  ║                                                           ║
echo  ║  Share YOUR URL with colleagues on the same network.     ║
echo  ║  They just open it in Chrome or Edge — nothing to install.║
echo  ║                                                           ║
echo  ║  Press Ctrl+C to stop the server.                        ║
echo  ╚══════════════════════════════════════════════════════════╝
echo.

REM ── Start Streamlit ───────────────────────────────────────────────
python -m streamlit run streamlit_spc.py ^
    --server.address=0.0.0.0 ^
    --server.port=8501 ^
    --server.headless=true ^
    --browser.gatherUsageStats=false ^
    --server.enableCORS=false ^
    --server.enableXsrfProtection=true ^
    --browser.serverAddress=%MY_IP%

REM ── If it stops, show error ───────────────────────────────────────
echo.
echo  Server stopped. Press any key to close.
pause
