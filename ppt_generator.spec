# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['ppt_generator.py'],
    pathex=[],
    binaries=[],
    datas=[('redbook.ico', '.')],
    hiddenimports=[
        'pandas',
        'PIL',
        'PIL.Image',
        'comtypes',
        'win32com',
        'win32com.client',
        'pythoncom',
        'openpyxl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'numpy',
        'numpy.core._methods',
        'numpy.lib.format',
        'requests',
        'urllib3',
        'charset_normalizer',
        'idna',
        'certifi'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'cv2',
        'PyQt5',
        'PyQt6',
        'PySide2',
        'PySide6',
        'wx',
        'spire',
        'akshare',
        'beautifulsoup4',
        'scrapy',
        'moviepy',
        'openai',
        'pdf2docx',
        'pywinauto',
        'schedule'
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='小红书图文批量制作工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='redbook.ico'
) 