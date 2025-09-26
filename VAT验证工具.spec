# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['vat_validator_gui.py'],  # 使用GUI版本作为主入口
    pathex=[],
    binaries=[],
    datas=[
        ('favicon.ico', '.'),  # 包含图标文件
    ],
    hiddenimports=[
        # PyQt5 核心模块
        'PyQt5.QtCore',
        'PyQt5.QtGui', 
        'PyQt5.QtWidgets',
        'PyQt5.QtPrintSupport',
        
        # 数据处理模块
        'openpyxl',
        'openpyxl.workbook',
        'openpyxl.worksheet',
        'openpyxl.styles',
        'openpyxl.utils',
        'pandas',
        'pandas.core',
        'pandas.io',
        'pandas.io.excel',
        
        # 文档处理模块
        'docx',
        'docx.shared',
        'docx.enum',
        'docx.enum.text',
        'docx.oxml',
        'docx.oxml.parser',
        'docx.oxml.ns',
        
        # 网络请求模块
        'requests',
        'requests.adapters',
        'requests.auth',
        'requests.cookies',
        'requests.models',
        'requests.sessions',
        'urllib3',
        
        # 图像处理模块
        'Pillow',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        
        # 系统和工具模块
        'json',
        'datetime',
        'threading',
        'concurrent',
        'concurrent.futures',
        'logging',
        'typing',
        'pathlib',
        'shutil',
        'os',
        'sys',
        'time',
        
        # 项目自定义模块
        'document_processor',  # 文档处理模块
        'vat_validator',       # VAT验证模块
        'vat_validator_cli',   # CLI模块
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
