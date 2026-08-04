# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['excel_tool.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('pages/*.py', 'pages'),
    ],
    hiddenimports=[
        'openpyxl',
        'openpyxl.cell._writer',
        'openpyxl.utils',
        'openpyxl.styles',
        'win32com',
        'win32com.client',
        'fitz',
        'PIL',
        'PIL.Image',
        'PIL.ImageDraw',
        'PIL.ImageFont',
    ],
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
    name='BBC百宝箱',
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
    icon='app.ico',
)
