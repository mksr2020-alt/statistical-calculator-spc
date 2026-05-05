@echo off
REM ══════════════════════════════════════════════════════════════════════════════
REM  SPC Calculator — One-Click .EXE Build Script
REM  Run this on your Windows machine to produce the distributable folder.
REM  Output:  dist\SPC Calculator\SPC Calculator.exe
REM
REM  Requirements:
REM    - Python 3.10+ installed and on PATH
REM    - Internet connection for first run (pip installs)
REM    - ~2 GB free disk space during build
REM    - ~400 MB disk space for the final bundle
REM ══════════════════════════════════════════════════════════════════════════════

title SPC Calculator — Build Tool
color 0A

cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════════════════════════════╗
echo  ║     SPC Statistical Process Calculator — Build Script       ║
echo  ╚══════════════════════════════════════════════════════════════╝
echo.

REM ── Detect Python ─────────────────────────────────────────────────────────
set PYTHON=
for %%P in (python python3 py) do (
    %%P --version >nul 2>&1
    if not errorlevel 1 (
        set PYTHON=%%P
        goto :python_found
    )
)
echo  [ERROR] Python not found. Install from https://python.org and add to PATH.
pause & exit /b 1

:python_found
echo  [OK] Using Python: & %PYTHON% --version
echo.

REM ── Step 1: Install / upgrade pip silently ─────────────────────────────────
echo  [1/6] Upgrading pip...
%PYTHON% -m pip install --upgrade pip --quiet
if errorlevel 1 (
    echo  [WARN] pip upgrade failed — continuing anyway.
)

REM ── Step 2: Install app runtime dependencies ──────────────────────────────
echo  [2/6] Installing app dependencies (scipy, plotly, openpyxl, reportlab, kaleido)...
%PYTHON% -m pip install -r requirements.txt --quiet
if errorlevel 1 (
    echo  [ERROR] Failed to install app dependencies. Check your internet connection.
    pause & exit /b 1
)

REM ── Step 3: Install packaging dependencies ─────────────────────────────────
echo  [3/6] Installing packaging tools (pyinstaller, pywebview, Pillow)...
%PYTHON% -m pip install pyinstaller>=6.0.0 pywebview>=4.4.0 Pillow --quiet
if errorlevel 1 (
    echo  [ERROR] Failed to install build tools.
    pause & exit /b 1
)

REM ── Step 4: Generate .ico icon from PNG ───────────────────────────────────
echo  [4/6] Generating Windows icon (icon.png → icon.ico)...
%PYTHON% generate_icon.py
if errorlevel 1 (
    echo  [WARN] Icon generation failed — continuing without custom icon.
)

REM ── Step 5: Clean previous build artifacts ────────────────────────────────
echo  [5/6] Cleaning previous build artifacts...
if exist "dist\SPC Calculator" (
    rmdir /s /q "dist\SPC Calculator"
    echo  Removed old dist folder.
)
if exist "build" (
    rmdir /s /q "build"
    echo  Removed old build folder.
)

REM ── Step 6: Run PyInstaller ───────────────────────────────────────────────
echo  [6/6] Building .exe — this will take 5-15 minutes, please wait...
echo.
%PYTHON% -m PyInstaller SPC.spec --clean --noconfirm

if errorlevel 1 (
    echo.
    echo  ╔══════════════════════════════════════════════════════════════╗
    echo  ║  BUILD FAILED. Review the errors above.                     ║
    echo  ╚══════════════════════════════════════════════════════════════╝
    pause & exit /b 1
)

REM ── Success banner ─────────────────────────────────────────────────────────
echo.
echo  ╔══════════════════════════════════════════════════════════════╗
echo  ║                                                              ║
echo  ║   BUILD COMPLETE!                                            ║
echo  ║                                                              ║
echo  ║   Your app is ready at:                                      ║
echo  ║   dist\SPC Calculator\SPC Calculator.exe                     ║
echo  ║                                                              ║
echo  ║   HOW TO DISTRIBUTE:                                         ║
echo  ║   Zip the entire "dist\SPC Calculator\" folder.              ║
echo  ║   Colleagues unzip and double-click "SPC Calculator.exe"     ║
echo  ║   — nothing to install, no Python needed.                    ║
echo  ║                                                              ║
echo  ╚══════════════════════════════════════════════════════════════╝
echo.

REM ── Optionally open the output folder ─────────────────────────────────────
set /p OPEN_FOLDER="Open output folder now? [Y/N]: "
if /i "%OPEN_FOLDER%"=="Y" (
    explorer "dist\SPC Calculator"
)

pause
