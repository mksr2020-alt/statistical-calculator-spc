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
def _patch_signal_for_thread():
    """
    CRITICAL FIX: Streamlit calls signal.signal(SIGTERM, ...) during startup,
    but signal handlers can only be set from the main thread in Python.
    Since we run Streamlit in a background thread, we patch signal.signal to
    silently skip registration when called from a non-main thread.
    """
    import signal as _sig
    import threading as _thr

    _original_signal = _sig.signal

    def _safe_signal(sig, handler):
        if _thr.current_thread() is _thr.main_thread():
            return _original_signal(sig, handler)
        # Silently ignore — not in main thread, signal registration not possible
        log.debug(f'[signal patch] Skipped signal.signal({sig}) — not in main thread.')

    _sig.signal = _safe_signal
    log.info('Signal patch applied.')


def run_streamlit(port: int) -> None:
    """
    Boot Streamlit using its bootstrap API in a background thread.
    The signal patch above makes this safe to run outside the main thread.
    """
    try:
        # Apply signal patch BEFORE any streamlit import
        _patch_signal_for_thread()

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

        # Detect which bootstrap.run() API version is installed:
        #   Old Streamlit (<1.12): run(main_script_path, is_hello: bool, args, flag_options)
        #   New Streamlit (>=1.12): run(main_script_path, command_line: str, args, flag_options)
        import inspect
        sig_params = list(inspect.signature(bootstrap.run).parameters.keys())
        log.info(f'bootstrap.run signature params: {sig_params}')

        if len(sig_params) > 1 and sig_params[1] == 'is_hello':
            # Old API — second arg must be a bool
            bootstrap.run(app_path, False, [], flag_options)
        else:
            # New API — second arg is command_line string
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
    background: #f8f9fa;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    height: 100vh;
    font-family: 'Segoe UI', system-ui, sans-serif;
    color: #1a1a1a; overflow: hidden; user-select: none;
  }
  .card {
    display: flex; flex-direction: column; align-items: center;
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 16px;
    padding: 48px 64px;
    box-shadow: 0 4px 24px rgba(0,0,0,0.08), 0 1px 4px rgba(0,0,0,0.04);
  }
  .icon-box {
    width: 80px; height: 80px; border-radius: 18px;
    background: #111827;
    color: #ffffff;
    display: flex; align-items: center; justify-content: center;
    font-size: 42px; margin-bottom: 24px;
    box-shadow: 0 4px 16px rgba(0,0,0,0.18);
  }
  h1 {
    font-size: 24px; font-weight: 700;
    color: #111827; letter-spacing: -0.4px;
    margin-bottom: 4px;
  }
  .sub {
    font-size: 12px; color: #9ca3af;
    letter-spacing: 1.8px; text-transform: uppercase;
    margin-bottom: 40px; font-weight: 500;
  }
  .prog-wrap {
    width: 220px; height: 2px;
    background: #f3f4f6;
    border-radius: 99px; overflow: hidden; margin-bottom: 14px;
  }
  .prog-bar {
    height: 100%; width: 0%;
    background: #111827;
    border-radius: 99px;
    animation: prog 5s cubic-bezier(0.4,0,0.2,1) forwards;
  }
  @keyframes prog { 0%{width:0%} 40%{width:55%} 70%{width:80%} 100%{width:93%} }
  .status {
    font-size: 11px; color: #9ca3af; letter-spacing: 0.3px;
  }
  .status::after {
    content: 'Initializing...';
    animation: txt 5s steps(1) infinite;
  }
  @keyframes txt {
    0%{content:'Initializing...'}
    25%{content:'Loading statistical modules...'}
    50%{content:'Preparing visualization engine...'}
    75%{content:'Almost ready...'}
  }
  .footer {
    position: fixed; bottom: 20px;
    font-size: 11px; color: #d1d5db;
    letter-spacing: 0.3px;
  }
  .dot {
    display: inline-block; width: 5px; height: 5px;
    background: #d1d5db; border-radius: 50%;
    margin: 0 3px; vertical-align: middle;
    animation: blink 1.4s ease-in-out infinite;
  }
  .dot:nth-child(2) { animation-delay: 0.2s; }
  .dot:nth-child(3) { animation-delay: 0.4s; }
  @keyframes blink { 0%,80%,100%{opacity:0.2} 40%{opacity:1} }
</style>
</head>
<body>
  <div class="card">
    <div class="icon-box">&#931;</div>
    <h1>SPC Calculator</h1>
    <p class="sub">Statistical Process Capability v1.0</p>
    <div class="prog-wrap"><div class="prog-bar"></div></div>
    <p class="status"></p>
  </div>
  <p class="footer">
    Starting application
    <span class="dot"></span><span class="dot"></span><span class="dot"></span>
  </p>
</body>
</html>"""

ERROR_HTML = """<!DOCTYPE html>
<html lang="en">
<head><meta charset="UTF-8"><title>Error</title>
<style>
  body { background:#f8f9fa; display:flex; flex-direction:column; align-items:center;
    justify-content:center; height:100vh; font-family:'Segoe UI',sans-serif; color:#111827; }
  .card { background:#fff; border:1px solid #e5e7eb; border-radius:16px;
    padding:48px 64px; text-align:center;
    box-shadow:0 4px 24px rgba(0,0,0,0.08); }
  .icon { font-size:48px; margin-bottom:16px; }
  h1 { font-size:20px; color:#dc2626; margin-bottom:10px; font-weight:700; }
  p { font-size:13px; color:#6b7280; max-width:420px; line-height:1.7; }
  code { background:#f3f4f6; padding:2px 8px; border-radius:4px;
    font-size:11px; color:#374151; border:1px solid #e5e7eb; }
</style>
</head>
<body>
  <div class="card">
    <div class="icon">&#9888;&#65039;</div>
    <h1>Application Failed to Start</h1>
    <p>
      The internal server could not be reached within 90 seconds.<br>
      Please close this window and try again.<br><br>
      A log file <code>spc_launcher.log</code> has been saved<br>
      next to the .exe with the full error message.
    </p>
  </div>
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
