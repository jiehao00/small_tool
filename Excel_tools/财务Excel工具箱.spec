# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['excel_tool.py'],
    pathex=[],
    binaries=[],
    datas=[('pages', 'pages')],
    hiddenimports=['openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'matplotlib', 'matplotlib.backends.backend_agg', 'fpdf2', 'fpdf'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='财务Excel工具箱',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['ChatGPT-Image-2026年7月28日-14_44_50.ico'],
)
