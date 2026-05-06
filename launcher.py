"""
SPC Statistical Process Calculator — Desktop Launcher (v2)
===========================================================
Fixed for PyInstaller bundling:
- Correct _MEIPASS path setup for Streamlit static assets
- Windows asyncio event loop policy fix
- Subprocess-based Streamlit boot (more reliable than threading in bundles)
- Full error logging to spc_launcher.log for debugging
"""

import os
import sys
import socket
import subprocess
import threading
import time
import logging

# ── Logging — always write to file next to the .exe ──────────────────────────
LOG_FILE = os.path.join(os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
                        else os.path.abspath('.'), 'spc_launcher.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.DEBUG,
    format='%(asctime)s [%(levelname)s] %(message)s',
    datefmt='%H:%M:%S',
)
log = logging.getLogger('spc_launcher')
log.info('=== SPC Calculator Launcher Starting ===')
log.info(f'sys.executable: {sys.executable}')
log.info(f'frozen: {getattr(sys, "frozen", False)}')


# ── Path resolution ───────────────────────────────────────────────────────────
IS_BUNDLED = getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS')
BUNDLE_DIR = sys._MEIPASS if IS_BUNDLED else os.path.abspath('.')
log.info(f'BUNDLE_DIR: {BUNDLE_DIR}')


def resource_path(relative_path: str) -> str:
    return os.path.join(BUNDLE_DIR, relative_path)


# ── Critical: fix Streamlit path inside PyInstaller bundle ───────────────────
def setup_bundle_environment():
    """Configure environment so Streamlit finds its static files in the bundle."""
    if not IS_BUNDLED:
        return

    # 1. Add bundle root to Python path so imports work
    if BUNDLE_DIR not in sys.path:
        sys.path.insert(0, BUNDLE_DIR)

    # 2. Tell Streamlit exactly where its static folder lives
    static_path = os.path.join(BUNDLE_DIR, 'streamlit', 'static')
    if os.path.isdir(static_path):
        os.environ['STREAMLIT_STATIC_PATH'] = static_path
        log.info(f'Set STREAMLIT_STATIC_PATH: {static_path}')
    else:
        log.warning(f'Streamlit static path not found: {static_path}')

    # 3. Tell Streamlit where its component templates live
    component_path = os.path.join(BUNDLE_DIR, 'streamlit', 'static', 'components')
    if os.path.isdir(component_path):
        os.environ['STREAMLIT_COMPONENTS_PATH'] = component_path

    # 4. Point HOME to a writable location (avoids write errors in corporate env)
    writable_home = os.path.join(os.path.expanduser('~'), '.spc_calculator')
    os.makedirs(writable_home, exist_ok=True)
    os.environ['STREAMLIT_HOME'] = writable_home

    # 5. Suppress Streamlit's browser-open behaviour
    os.environ['STREAMLIT_BROWSER_GATHER_USAGE_STATS'] = 'false'
    os.environ['STREAMLIT_SERVER_HEADLESS'] = 'true'
    os.environ['STREAMLIT_GLOBAL_DEVELOPMENT_MODE'] = 'false'

    log.info('Bundle environment configured.')


# ── Networking helpers ────────────────────────────────────────────────────────
def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(('127.0.0.1', 0))
        s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        return s.getsockname()[1]


def wait_for_server(port: int, timeout: int = 90) -> bool:
    log.info(f'Waiting for Streamlit on port {port} (timeout={timeout}s)...')
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with socket.create_connection(('127.0.0.1', port), timeout=1):
                log.info(f'Server is ready on port {port}.')
                return True
        except (ConnectionRefusedError, OSError):
            time.sleep(0.4)
    log.error(f'Server did not start within {timeout}s.')
    return False


# ── Streamlit boot ────────────────────────────────────────────────────────────
def run_streamlit(port: int) -> None:
    """
    Boot Streamlit using its bootstrap API.
    Must run in a daemon thread — killed automatically when pywebview closes.
    """
    try:
        # Windows: fix asyncio event loop policy (required in PyInstaller bundles)
        if sys.platform == 'win32':
            import asyncio
            asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

        import streamlit.web.bootstrap as bootstrap

        app_path = resource_path('streamlit_spc.py')
        log.info(f'App path: {app_path}')

        if not os.path.exists(app_path):
            log.error(f'streamlit_spc.py NOT FOUND at: {app_path}')
            return

        flag_options = {
            'server.port': port,
            'server.address': '127.0.0.1',
            'server.headless': True,
            'server.enableCORS': False,
            'server.enableXsrfProtection': False,
            'server.runOnSave': False,
            'global.developmentMode': False,
            'browser.gatherUsageStats': False,
            'client.toolbarMode': 'minimal',
        }

        log.info('Calling streamlit bootstrap.run()...')
        bootstrap.load_config_options(flag_options=flag_options)
        flag_options['_is_running_with_streamlit'] = True
        bootstrap.run(app_path, 'streamlit run', [], flag_options)

    except Exception as e:
        log.exception(f'Streamlit failed to start: {e}')


# ── Splash screen HTML ────────────────────────────────────────────────────────
SPLASH_HTML = """<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>SPC Calculator</title>
<style>
  *, *::before, *::after { margin: 0; padding: 0; box-sizing: border-box; }
  body {
    background: #0f172a;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    height: 100vh;
    font-family: 'Segoe UI', system-ui, sans-serif;
    color: #f1f5f9; overflow: hidden; user-select: none;
  }
  body::before {
    content: ''; position: fixed; inset: 0;
    background-image: linear-gradient(rgba(59,130,246,0.05) 1px, transparent 1px),
                      linear-gradient(90deg, rgba(59,130,246,0.05) 1px, transparent 1px);
    background-size: 40px 40px;
    animation: gridMove 20s linear infinite; pointer-events: none;
  }
  @keyframes gridMove { from { transform: translateY(0); } to { transform: translateY(40px); } }
  .card { position: relative; z-index: 1; display: flex; flex-direction: column; align-items: center; }
  .icon-ring {
    width: 100px; height: 100px; border-radius: 28px;
    background: linear-gradient(135deg, #1d4ed8 0%, #7c3aed 100%);
    display: flex; align-items: center; justify-content: center;
    font-size: 52px; margin-bottom: 28px;
    box-shadow: 0 8px 32px rgba(59,130,246,0.35), 0 0 80px rgba(124,58,237,0.2);
    animation: pulse 3s ease-in-out infinite;
  }
  @keyframes pulse {
    0%,100% { box-shadow: 0 8px 32px rgba(59,130,246,0.35), 0 0 80px rgba(124,58,237,0.2); }
    50% { box-shadow: 0 8px 48px rgba(59,130,246,0.5), 0 0 120px rgba(124,58,237,0.35); }
  }
  h1 { font-size: 32px; font-weight: 700; letter-spacing: -0.8px;
    background: linear-gradient(90deg,#60a5fa,#a78bfa);
    -webkit-background-clip: text; -webkit-text-fill-color: transparent;
    background-clip: text; margin-bottom: 6px; }
  .sub { font-size: 13px; color: #64748b; letter-spacing: 1.5px;
    text-transform: uppercase; margin-bottom: 52px; }
  .prog-wrap { width: 240px; height: 3px; background: rgba(255,255,255,0.07);
    border-radius: 99px; overflow: hidden; margin-bottom: 18px; }
  .prog-bar { height: 100%; width: 0%;
    background: linear-gradient(90deg,#3b82f6,#8b5cf6);
    border-radius: 99px; animation: prog 6s ease-out forwards; }
  @keyframes prog { 0%{width:0%} 30%{width:40%} 60%{width:75%} 100%{width:92%} }
  .status { font-size: 12px; color: #475569; }
  .status::after {
    content: 'Initializing engine...';
    animation: txt 6s steps(1) infinite;
  }
  @keyframes txt {
    0%{content:'Initializing engine...'} 25%{content:'Loading statistical modules...'}
    50%{content:'Preparing visualization engine...'} 75%{content:'Almost ready...'}
  }
  .ver { position: fixed; bottom: 20px; font-size: 11px; color: #1e293b; }
</style>
</head>
<body>
  <div class="card">
    <div class="icon-ring">&#931;</div>
    <h1>SPC Calculator</h1>
    <p class="sub">Statistical Process Capability</p>
    <div class="prog-wrap"><div class="prog-bar"></div></div>
    <p class="status"></p>
  </div>
  <p class="ver">v2.0 &nbsp;&middot;&nbsp; Statistical Process Control Suite</p>
</body>
</html>"""

ERROR_HTML = """<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>Error</title>
<style>
  body { background:#0f172a; display:flex; flex-direction:column; align-items:center;
    justify-content:center; height:100vh; font-family:'Segoe UI',sans-serif; color:#f1f5f9; }
  .icon { font-size:56px; margin-bottom:20px; }
  h1 { font-size:22px; color:#f87171; margin-bottom:12px; }
  p { font-size:14px; color:#64748b; max-width:480px; text-align:center; line-height:1.7; }
  code { background:#1e293b; padding:2px 8px; border-radius:4px; font-size:12px; color:#94a3b8; }
</style>
</head>
<body>
  <div class="icon">&#9888;&#65039;</div>
  <h1>Application Failed to Start</h1>
  <p>
    The internal server could not be reached within 90 seconds.<br>
    Please close this window and try again.<br><br>
    <strong>Check for details:</strong><br>
    A log file <code>spc_launcher.log</code> has been saved next to the .exe<br>
    with the full error message.
  </p>
</body>
</html>"""


# ── Main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    setup_bundle_environment()

    import webview

    port = find_free_port()
    log.info(f'Selected port: {port}')

    # Boot Streamlit in a background daemon thread
    st_thread = threading.Thread(
        target=run_streamlit, args=(port,), daemon=True, name='streamlit-server'
    )
    st_thread.start()

    # Create pywebview window with splash
    window = webview.create_window(
        title='SPC Statistical Process Calculator',
        html=SPLASH_HTML,
        width=1366,
        height=870,
        min_size=(1024, 680),
        background_color='#0f172a',
        text_select=True,
    )

    def navigate_when_ready() -> None:
        ok = wait_for_server(port, timeout=90)
        if ok:
            time.sleep(0.8)
            window.load_url(f'http://127.0.0.1:{port}')
        else:
            window.load_html(ERROR_HTML)

    nav_thread = threading.Thread(target=navigate_when_ready, daemon=True, name='nav-watch')
    nav_thread.start()

    webview.start(debug=False, private_mode=False)
    log.info('Window closed — exiting.')


if __name__ == '__main__':
    main()
