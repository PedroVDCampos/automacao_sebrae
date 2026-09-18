# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas=[]
binaries=[]
hiddenimports=[]

datas.append(("assets/icone.ico", "assets"))

tmp=collect_all("customtkinter")
datas += tmp[0]
binaries += tmp[1]
hiddenimports += tmp[2]

a=Analysis(
    ["main.pyw"],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["selenium", "webdriver_manager"],
    noarchive=False,
)
pyz=PYZ(a.pure)
exe=EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="RAE_Turbo",
    icon="assets/icone.ico",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
