# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for packaging tab2long_henan.py as a single-file console EXE.
# Usage: pyinstaller tab2long_henan.spec

import os
from PyInstaller.utils.hooks import collect_submodules, collect_data_files

script_path = os.path.abspath("tab2long_henan.py")

# Hidden imports often required by pandas/numpy/openpyxl
hidden = []
hidden += collect_submodules("pandas")
hidden += collect_submodules("numpy")
hidden += collect_submodules("openpyxl")
hidden += collect_submodules("xlrd")

# Data files needed by openpyxl (templates) and pandas
datas = []
datas += collect_data_files("pandas", include_py_files=True)
datas += collect_data_files("openpyxl", include_py_files=True)

a = Analysis(
    [script_path],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=None)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    name="tab2long_henan",
    console=True,
    disable_windowed_traceback=False,
    # You can set icon to a .ico if you add one: icon='app.ico'
    uptodate=True,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="tab2long_henan",
)
