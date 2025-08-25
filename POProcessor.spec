# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for PO Processor
"""

import sys
from pathlib import Path

block_cipher = None

# Get the absolute path to the project
project_dir = Path.cwd()

a = Analysis(
    ['gui_app.py'],
    pathex=[str(project_dir)],
    binaries=[],
    datas=[
        ('src', 'src'),  # Include the src package
    ],
    hiddenimports=[
        'PIL',
        'PIL.Image',
        'pdf2image',
        'pytesseract',
        'pandas',
        'numpy',
        'xlsxwriter',
        'openpyxl',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.ttk',
        'pandas.io.excel._xlsxwriter',
        'pandas.io.excel._openpyxl',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'scipy',
        'jupyter',
        'notebook',
        'pytest',
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
    name='POProcessor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # Don't use UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window for GUI app
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico' if Path('icon.ico').exists() else None,
)

# For macOS, create an app bundle
if sys.platform == 'darwin':
    app = BUNDLE(
        exe,
        name='POProcessor.app',
        icon='icon.icns' if Path('icon.icns').exists() else None,
        bundle_identifier='com.yourcompany.poprocessor',
        info_plist={
            'NSHighResolutionCapable': 'True',
            'CFBundleShortVersionString': '1.0.0',
        },
    )