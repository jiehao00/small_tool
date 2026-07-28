# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_dynamic_libs
from PyInstaller.utils.hooks import collect_submodules

binaries = []
hiddenimports = ['pandas._libs.tslibs', 'pandas._libs.join', 'pandas._libs.ops', 'pandas._libs.parsers', 'openpyxl', 'openpyxl.cell._writer', 'openpyxl.utils', 'openpyxl.styles', 'lxml.etree', 'lxml._elementpath', 'fitz', 'fitz.fitz', 'PIL', 'PIL._imaging', 'PIL._imagingft', 'win32com', 'win32com.client', 'pythoncom', 'tkinter', 'tkinter.filedialog', 'tkinter.messagebox', 'xml', 'xml.etree', 'json', 'zipfile', 're', 'subprocess']
binaries += collect_dynamic_libs('fitz')
binaries += collect_dynamic_libs('lxml')
binaries += collect_dynamic_libs('win32com')
binaries += collect_dynamic_libs('openpyxl')
binaries += collect_dynamic_libs('PIL')
hiddenimports += collect_submodules('fitz')
hiddenimports += collect_submodules('lxml')
hiddenimports += collect_submodules('win32com')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('PIL')


a = Analysis(
    ['generate_salary_slips.py'],
    pathex=[],
    binaries=binaries,
    datas=[],
    hiddenimports=hiddenimports,
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
    name='工资条生成工具',
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
)
