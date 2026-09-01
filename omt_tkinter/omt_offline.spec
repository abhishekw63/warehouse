# PyInstaller spec — OMT Offline
#
#     pyinstaller omt_offline.spec --noconfirm
#
# Produces ONE self-contained dist/OMT Offline.exe. No Python install needed on
# the target machine, no network, no database.
#
# Why the explicit hiddenimports: pandas picks its Excel engine at CALL time
# (openpyxl for .xlsx, xlrd for .xls, pyxlsb for .xlsb). PyInstaller's static
# analysis therefore never sees them, and the .exe would raise
# "Missing optional dependency 'xlrd'" the first time an RK or GT Mass .xls is
# opened — on the user's machine, not ours. Declare them.

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('assets/renee.ico', 'assets'),      # window icon at runtime
    ],
    hiddenimports=[
        'openpyxl',                          # .xlsx read + all workbook writing
        'xlrd',                              # .xls  (RK POItemExport, some GT Mass)
        'pyxlsb',                            # .xlsb (H&B-style sheets)
        'pandas._libs.tslibs.base',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Keep the build lean: the folder still carries the old PyQt6 reference app,
    # and matplotlib/scipy ride in on pandas' coat-tails. None are imported.
    excludes=[
        'PyQt6', 'PyQt5', 'PySide6', 'PySide2',
        'matplotlib', 'scipy', 'IPython', 'jupyter', 'notebook',
        'pytest', 'sphinx', 'django',
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
    name='OMT Offline',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,                           # GUI app — no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='assets/renee.ico',
)
