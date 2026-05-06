# -*- mode: python ; coding: utf-8 -*-
# ══════════════════════════════════════════════════════════════════════════════
#  SPC Calculator — PyInstaller Specification File
#  Build command:  pyinstaller SPC.spec --clean --noconfirm
# ══════════════════════════════════════════════════════════════════════════════

from PyInstaller.utils.hooks import collect_data_files, collect_submodules, copy_metadata

block_cipher = None

# ── Data files ────────────────────────────────────────────────────────────────
datas = []

# Streamlit: static assets (JS bundles, CSS, fonts, component manifests)
datas += collect_data_files("streamlit", include_py_files=True)
datas += copy_metadata("streamlit")

# Altair: schema JSON files required by Streamlit's built-in chart elements
datas += collect_data_files("altair", include_py_files=True)
datas += copy_metadata("altair")

# Plotly: layout schema, built-in templates, icon sprites
datas += collect_data_files("plotly")

# Openpyxl: built-in Excel templates and style descriptors
datas += collect_data_files("openpyxl")

# ReportLab: fonts (Helvetica, Times, Courier), color database
datas += collect_data_files("reportlab")
datas += collect_data_files("reportlab", subdir="fonts")

# Kaleido: binary renderer for converting Plotly figures to PNG
datas += collect_data_files("kaleido")

# Pywebview: JavaScript bridge files needed by the native window
datas += collect_data_files("webview")

# The main Streamlit app — placed at root of the bundle
datas += [("streamlit_spc.py", ".")]

# App icon asset
import os
if os.path.exists("assets/icon.ico"):
    datas += [("assets/icon.ico", "assets")]
elif os.path.exists("assets/icon.png"):
    datas += [("assets/icon.png", "assets")]

# Explicitly collect Streamlit's static web assets (JS, CSS, fonts)
# This is the most common cause of "server failed to start" in PyInstaller bundles
import streamlit as _st
import pathlib as _pl
_st_pkg_dir = _pl.Path(_st.__file__).parent
_st_static = _st_pkg_dir / "static"
_st_runtime = _st_pkg_dir / "runtime"
if _st_static.exists():
    datas += [(str(_st_static), "streamlit/static")]
if _st_runtime.exists():
    datas += [(str(_st_runtime), "streamlit/runtime")]
# Streamlit component lib (for st.components.v1)
_st_components = _st_pkg_dir / "components"
if _st_components.exists():
    datas += [(str(_st_components), "streamlit/components")]


# ── Hidden imports ────────────────────────────────────────────────────────────
# PyInstaller's static analysis misses dynamically imported sub-modules.
# We list everything our lazy-loading pattern uses.

hiddenimports = []

# Streamlit internals
hiddenimports += collect_submodules("streamlit")
hiddenimports += [
    "streamlit.web.bootstrap",
    "streamlit.runtime.scriptrunner",
    "streamlit.components.v1",
]

# Tornado (Streamlit's async HTTP server)
hiddenimports += collect_submodules("tornado")

# SciPy — C extensions that static analysis cannot detect
hiddenimports += collect_submodules("scipy")
hiddenimports += [
    "scipy.stats",
    "scipy.special",
    "scipy.optimize",
    "scipy._lib.messagestream",
]

# NumPy / Pandas internals
hiddenimports += ["numpy.core._methods", "numpy.lib.format", "pandas.io.formats.style"]

# Plotly runtime
hiddenimports += ["plotly.graph_objects", "plotly.express", "plotly.io"]

# Openpyxl internals (cell writer, styles)
hiddenimports += [
    "openpyxl.cell._writer",
    "openpyxl.styles.fills",
    "openpyxl.styles.borders",
    "openpyxl.styles.fonts",
    "openpyxl.styles.alignment",
    "openpyxl.drawing.image",
    "openpyxl.utils",
]

# ReportLab
hiddenimports += [
    "reportlab.graphics",
    "reportlab.platypus",
    "reportlab.lib.styles",
    "reportlab.lib.colors",
    "reportlab.lib.units",
    "reportlab.lib.pagesizes",
    "reportlab.pdfgen.canvas",
]

# Kaleido image renderer
hiddenimports += ["kaleido.scopes.plotly"]

# PyWebView
hiddenimports += ["webview", "webview.platforms.winforms"]

# Standard library modules sometimes missed
hiddenimports += ["pkg_resources.py2_warn", "packaging.version", "packaging.specifiers"]


# ── Analysis ──────────────────────────────────────────────────────────────────
a = Analysis(
    ["launcher.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Explicitly exclude heavy unused packages to keep bundle lean
    excludes=[
        "matplotlib",
        "tkinter",
        "_tkinter",
        "PyQt5",
        "PyQt6",
        "PySide2",
        "PySide6",
        "wx",
        "IPython",
        "jupyter",
        "notebook",
        "jedi",
        "black",
        "mypy",
        "pylint",
        "pytest",
        "sphinx",
        "docutils",
        "pydoc",
        "test",
        "unittest",
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)


# ── Executable & Collection (--onedir standard method) ───────────────────────
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,     # keep in directory (faster startup vs --onefile)
    name="SPC Calculator",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,                  # compress binaries with UPX (reduces size ~30%)
    console=False,             # no black terminal window — GUI only
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # Icon: prefer .ico; fall back gracefully
    icon="assets/icon.ico" if os.path.exists("assets/icon.ico") else (
          "assets/icon.png" if os.path.exists("assets/icon.png") else None),
    version_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="SPC Calculator",     # → dist/SPC Calculator/ folder
)
