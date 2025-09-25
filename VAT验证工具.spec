# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['vat_validator_gui.py'],
    pathex=[],
    binaries=[],
    datas=[('document_processor.py', '.')],
    hiddenimports=[
        'document_processor',
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.styles',
        'requests',
        'PyQt5.QtCore',
        'PyQt5.QtWidgets',
        'PyQt5.QtGui',
        'concurrent.futures',
        'threading',
        'datetime',
        'json',
        'time',
        'shutil',
        'os',
        'sys'
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
    name='VAT验证工具',
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
    icon='favicon.ico',  # 使用SVG图标文件
)
app = BUNDLE(
    exe,
    name='VAT验证工具.app',
    icon=favicon.ico,
    bundle_identifier=None,
)
