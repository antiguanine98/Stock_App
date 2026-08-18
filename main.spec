# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec for Stock_App v1.34 (Windows one-file, Python 3.12)."""

from PyInstaller.utils.hooks import collect_all, collect_data_files

block_cipher = None

datas = [("viewer.html", ".")]
binaries = []
hiddenimports = [
    "PyQt6.QtWebEngineWidgets",
    "PyQt6.QtWebEngineCore",
    "PyQt6.QtWebChannel",
    "matplotlib.backends.backend_qtagg",
    "mplcursors",
    "openpyxl",
    "pandas",
    "google.genai",
]

for pkg in ("PyQt6.QtWebEngineCore", "PyQt6.QtWebEngineWidgets", "matplotlib"):
    try:
        pkg_datas, pkg_binaries, pkg_hidden = collect_all(pkg)
        datas += pkg_datas
        binaries += pkg_binaries
        hiddenimports += pkg_hidden
    except Exception:
        pass

try:
    datas += collect_data_files("matplotlib")
except Exception:
    pass

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=["hooks/rthook_sanitize_path.py"],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="Stock_App_v1.34",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
