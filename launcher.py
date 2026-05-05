"""
SPC Statistical Process Calculator — Desktop Launcher
======================================================
This script is the entry point for the packaged .exe.
It boots a Streamlit server on a free local port, shows a branded
splash screen, then navigates to the running app inside a native
pywebview desktop window.

Usage (development):
    python launcher.py

Usage (production):
    Double-click SPC Calculator.exe
"""

import os
import sys
import socket
import threading
import time
import logging

# ── Suppress noisy Streamlit / tornado boot messages ─────────────────────────
logging.getLogger("streamlit").setLevel(logging.ERROR)
logging.getLogger("tornado").setLevel(logging.ERROR)

# ── Path resolution (works both in dev and inside PyInstaller bundle) ─────────

def resource_path(relative_path: str) -> str:
    """Return absolute path — handles PyInstaller's _MEIPASS temp directory."""
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, relative_path)


# ── Networking helpers ────────────────────────────────────────────────────────

def find_free_port() -> int:
    """Bind to port 0 and let the OS pick a free one."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        return s.getsockname()[1]


def wait_for_server(port: int, timeout: int = 60) -> bool:
    """Poll localhost until the Streamlit server accepts connections."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(("127.0.0.1", port), timeout=1):
                return True
        except (ConnectionRefusedError, OSError):
            time.sleep(0.35)
    return False


# ── Streamlit boot ────────────────────────────────────────────────────────────

def run_streamlit(port: int) -> None:
    """Start the Streamlit web server programmatically on the given port."""
    import streamlit.web.bootstrap as bootstrap  # lazy — not loaded at module level

    app_path = resource_path("streamlit_spc.py")

    flag_options = {
        "server.port": port,
        "server.address": "127.0.0.1",
        "server.headless": True,
        "server.enableCORS": False,
        "server.enableXsrfProtection": False,
        "server.runOnSave": False,
        "global.developmentMode": False,
        "browser.gatherUsageStats": False,
        "client.toolbarMode": "minimal",  # hide the Streamlit hamburger menu
        "theme.base": "light",
    }

    bootstrap.load_config_options(flag_options=flag_options)
    flag_options["_is_running_with_streamlit"] = True
    bootstrap.run(app_path, "streamlit run", [], flag_options)


# ── Splash screen HTML ────────────────────────────────────────────────────────

SPLASH_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>SPC Calculator — Loading</title>
<style>
  *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }

  body {
    background: #0f172a;
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: center;
    height: 100vh;
    font-family: 'Segoe UI', system-ui, -apple-system, sans-serif;
    color: #f1f5f9;
    overflow: hidden;
    user-select: none;
  }

  /* Animated background grid */
  body::before {
    content: '';
    position: fixed;
    inset: 0;
    background-image:
      linear-gradient(rgba(59,130,246,0.05) 1px, transparent 1px),
      linear-gradient(90deg, rgba(59,130,246,0.05) 1px, transparent 1px);
    background-size: 40px 40px;
    animation: gridMove 20s linear infinite;
    pointer-events: none;
  }
  @keyframes gridMove {
    from { transform: translateY(0); }
    to   { transform: translateY(40px); }
  }

  .card {
    position: relative;
    z-index: 1;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 0;
  }

  /* Sigma icon */
  .icon-ring {
    width: 100px;
    height: 100px;
    border-radius: 28px;
    background: linear-gradient(135deg, #1d4ed8 0%, #7c3aed 100%);
    display: flex;
    align-items: center;
    justify-content: center;
    font-size: 52px;
    font-weight: 700;
    margin-bottom: 28px;
    box-shadow:
      0 0 0 1px rgba(255,255,255,0.08),
      0 8px 32px rgba(59,130,246,0.35),
      0 0 80px rgba(124,58,237,0.2);
    animation: iconPulse 3s ease-in-out infinite;
    letter-spacing: -2px;
  }
  @keyframes iconPulse {
    0%, 100% { box-shadow: 0 0 0 1px rgba(255,255,255,0.08), 0 8px 32px rgba(59,130,246,0.35), 0 0 80px rgba(124,58,237,0.2); }
    50%       { box-shadow: 0 0 0 1px rgba(255,255,255,0.12), 0 8px 48px rgba(59,130,246,0.5),  0 0 120px rgba(124,58,237,0.35); }
  }

  h1 {
    font-size: 32px;
    font-weight: 700;
    letter-spacing: -0.8px;
    background: linear-gradient(90deg, #60a5fa, #a78bfa);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    margin-bottom: 6px;
  }

  .subtitle {
    font-size: 14px;
    color: #64748b;
    letter-spacing: 1.5px;
    text-transform: uppercase;
    margin-bottom: 52px;
    font-weight: 500;
  }

  /* Progress bar */
  .progress-wrap {
    width: 240px;
    height: 3px;
    background: rgba(255,255,255,0.07);
    border-radius: 99px;
    overflow: hidden;
    margin-bottom: 18px;
  }
  .progress-bar {
    height: 100%;
    width: 0%;
    background: linear-gradient(90deg, #3b82f6, #8b5cf6);
    border-radius: 99px;
    animation: progressAnim 4s ease-out forwards;
  }
  @keyframes progressAnim {
    0%   { width: 0%; }
    40%  { width: 55%; }
    80%  { width: 85%; }
    100% { width: 95%; }
  }

  .status-text {
    font-size: 12px;
    color: #475569;
    letter-spacing: 0.3px;
    animation: statusCycle 4s steps(1) infinite;
  }
  @keyframes statusCycle {
    0%   { content: 'Initializing engine...'; }
  }
  .status-text::after {
    content: 'Initializing engine...';
    animation: textCycle 4s steps(1) infinite;
  }
  @keyframes textCycle {
    0%   { content: 'Initializing engine...'; }
    25%  { content: 'Loading statistical modules...'; }
    50%  { content: 'Preparing visualization engine...'; }
    75%  { content: 'Almost ready...'; }
  }

  /* Version tag */
  .version {
    position: fixed;
    bottom: 20px;
    font-size: 11px;
    color: #1e293b;
    letter-spacing: 0.5px;
  }
</style>
</head>
<body>
  <div class="card">
    <div class="icon-ring">Σ</div>
    <h1>SPC Calculator</h1>
    <p class="subtitle">Statistical Process Capability</p>
    <div class="progress-wrap">
      <div class="progress-bar"></div>
    </div>
    <p class="status-text"></p>
  </div>
  <p class="version">v2.0 &nbsp;·&nbsp; Statistical Process Control Suite</p>
</body>
</html>
"""

ERROR_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>Startup Error</title>
<style>
  body {
    background: #0f172a;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    height: 100vh;
    font-family: 'Segoe UI', sans-serif;
    color: #f1f5f9;
  }
  .icon { font-size: 56px; margin-bottom: 20px; }
  h1 { font-size: 22px; color: #f87171; margin-bottom: 12px; }
  p  { font-size: 14px; color: #64748b; max-width: 420px; text-align: center; line-height: 1.6; }
  code { background: #1e293b; padding: 2px 8px; border-radius: 4px; font-size: 12px; color: #94a3b8; }
</style>
</head>
<body>
  <div class="icon">⚠️</div>
  <h1>Application Failed to Start</h1>
  <p>
    The internal server could not be reached within 60 seconds.<br>
    Please close this window and try again.<br><br>
    If the problem persists, check that no other process is blocking the port
    or try running <code>SPC Calculator.exe</code> as Administrator.
  </p>
</body>
</html>
"""


# ── Main entry point ──────────────────────────────────────────────────────────

def main() -> None:
    import webview  # lazy import — not available during build analysis phase

    port = find_free_port()

    # ── 1. Boot Streamlit in a daemon thread (auto-killed on exit) ────────────
    st_thread = threading.Thread(
        target=run_streamlit, args=(port,), daemon=True, name="streamlit-server"
    )
    st_thread.start()

    # ── 2. Create pywebview window with the splash screen ────────────────────
    window = webview.create_window(
        title="SPC Statistical Process Calculator",
        html=SPLASH_HTML,
        width=1366,
        height=870,
        min_size=(1024, 680),
        background_color="#0f172a",
        text_select=True,          # allow copying text from the app
        confirm_close=False,
    )

    # ── 3. Background thread: wait for Streamlit, then navigate ──────────────
    def navigate_when_ready() -> None:
        ok = wait_for_server(port, timeout=60)
        if ok:
            # A brief extra pause ensures the first Streamlit render is ready
            time.sleep(0.6)
            window.load_url(f"http://127.0.0.1:{port}")
        else:
            window.load_html(ERROR_HTML)

    nav_thread = threading.Thread(
        target=navigate_when_ready, daemon=True, name="nav-watch"
    )
    nav_thread.start()

    # ── 4. Start the pywebview event loop (blocks until window is closed) ────
    webview.start(debug=False, private_mode=False)


if __name__ == "__main__":
    main()
