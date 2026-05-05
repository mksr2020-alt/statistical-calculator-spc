@echo off
REM ══════════════════════════════════════════════════════════════════
REM  SPC Statistical Process Calculator — Server Start Script
REM  Double-click this file to start. Keep the window open.
REM  Requires Python 3.12 to be installed.
REM ══════════════════════════════════════════════════════════════════

title SPC Calculator — Server

REM ── Move to the folder where this .bat file lives ─────────────────
cd /d "%~dp0"

REM ── Use virtual environment if it exists, else fall back to system Python ──
if exist ".venv\Scripts\python.exe" (
    echo  Using virtual environment...
    set PYTHON=.venv\Scripts\python.exe
    set PIP=.venv\Scripts\pip.exe
) else (
    echo  No virtual environment found. Checking system Python...
    echo  Tip: Run "python -m venv .venv" to create one.
    python --version >nul 2>&1
    if not errorlevel 1 (
        set PYTHON=python
        set PIP=pip
    ) else (
        py --version >nul 2>&1
        if not errorlevel 1 (
            set PYTHON=py
            set PIP=py -m pip
        ) else (
            REM -- Aggressive hunt for common installation paths --
            set PYTHON=python
            set PIP=pip
            for %%P in (
                "%USERPROFILE%\AppData\Local\Programs\Python\Python312\python.exe"
                "%USERPROFILE%\AppData\Local\Programs\Python\Python311\python.exe"
                "C:\Python312\python.exe"
                "C:\Program Files\Python312\python.exe"
                "C:\Program Files\Python311\python.exe"
            ) do (
                if exist %%P (
                    set PYTHON=%%P
                    set PIP=%%P -m pip
                    goto :found_python
                )
            )
            :found_python
            echo. >nul
        )
    )
)

REM ── Check Python is available ─────────────────────────────────────
%PYTHON% --version >nul 2>&1
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
%PYTHON% -m streamlit --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo  Streamlit not found. Installing dependencies...
    echo.
    %PIP% install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo  Install failed. Try:
        echo  %PIP% install -r requirements.txt --trusted-host pypi.org --trusted-host files.pythonhosted.org
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
%PYTHON% -m streamlit run streamlit_spc.py ^
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
