# compress.spec
# coding: utf-8
from PyInstaller.utils.hooks import collect_submodules
from pathlib import Path


import os
project_root = Path(os.getcwd())

ghostscript_dir = project_root / "tools" / "ghostscript"

a = Analysis(
    ['compress.py'],
    pathex=[],
    binaries=[
        (str(ghostscript_dir / "bin" / "gswin64c.exe"), "tools/ghostscript/bin")
    ],
    datas=[
        (str(project_root / "tools" / "ghostscript"), "tools/ghostscript")
    ],
    hiddenimports=collect_submodules(""),
    hookspath=[],
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=None
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='pdf_compressor',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    name='pdf_compressor'
)
