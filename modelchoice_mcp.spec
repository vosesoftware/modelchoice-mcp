# PyInstaller spec for ModelChoice MCP Server.
#
# Produces a single-file Windows .exe that boots the MCP server over
# stdio. Tool/prompt/resource modules are only pulled in via side-effect
# imports in server.py, so they're listed as hidden imports to keep
# PyInstaller's static analyser from stripping them. The COM modules
# (win32com etc.) are used by the bridge's GetActiveObject attach.
#
# Build locally with:
#     uv run pyinstaller modelchoice_mcp.spec --clean

# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

_PACKAGE_MODULES = [
    "modelchoice_mcp.server",
    "modelchoice_mcp.tools",
    "modelchoice_mcp.prompts",
    "modelchoice_mcp.resources",
    "modelchoice_mcp.schemas",
    "modelchoice_mcp.bridge",
    "modelchoice_mcp.store",
    "modelchoice_mcp.tree",
]
_COM_MODULES = [
    "win32com",
    "win32com.client",
    "pywintypes",
    "win32api",
    "xlwings",
]


a = Analysis(
    ["src/modelchoice_mcp/__main__.py"],
    pathex=["src"],
    binaries=[],
    datas=[],
    hiddenimports=_PACKAGE_MODULES + _COM_MODULES,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "tkinter",
        "matplotlib",
        "PySide6",
        "PyQt6",
    ],
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
    name="modelchoice-mcp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
