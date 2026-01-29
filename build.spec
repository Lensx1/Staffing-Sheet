# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for Staffing Sheet Generator
This creates a single-file Windows executable that bundles all dependencies.
"""

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'xlrd',
        'openpyxl.cell._writer',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.np_datetime',
        # Note: pandas._libs.skiplist may not exist in all pandas versions
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
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
    name='StaffingSheetGenerator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Compresses executable (~30% smaller). Set to False if build fails with UPX errors.
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Show console window for progress messages
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon=path/to/icon.ico if you have one
)
