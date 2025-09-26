# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['vat_validator_gui.py'],  # 使用GUI版本作为主入口
    pathex=[],
    binaries=[],
    datas=[
        ('favicon.ico', '.'),  # 包含图标文件
    ],
    hiddenimports=[
        'PyQt5.QtCore',
        'PyQt5.QtGui', 
        'PyQt5.QtWidgets',
        'openpyxl',
        'pandas',
        'docx',
        'requests',
        'Pillow',
        'PIL',
        'document_processor',  # 确保包含文档处理模块
        'vat_validator',       # 确保包含VAT验证模块
        'vat_validator_cli',   # 确保包含CLI模块
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
    console=False,  # 不显示控制台窗口
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='favicon.ico',  # 添加图标文件
)
app = BUNDLE(
    exe,
    name='VAT验证工具.app',
    icon=None,
    bundle_identifier=None,
)
