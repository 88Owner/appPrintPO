# -*- mode: python ; coding: utf-8 -*-
# PyInstaller — gọi: pyinstaller AppPrintPO.spec
from PyInstaller.utils.hooks import collect_data_files

datas = collect_data_files("reportlab")

a = Analysis(
    ["main.py"],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=[
        "reportlab.pdfbase._fontdata",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        "sqlalchemy",
        "matplotlib",
        "scipy",
        "pandas.tests",
    ],
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
    name="AppPrintPO",
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
    icon=["assets\\app.ico"],
)
