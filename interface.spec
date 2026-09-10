# -*- mode: python ; coding: utf-8 -*-

import sys

sys.setrecursionlimit(sys.getrecursionlimit() * 5)

from PyInstaller.utils.hooks import (
    collect_data_files,
    collect_dynamic_libs,
    collect_submodules,
)

block_cipher = None

hidden_imports = [
    "iva_process",
    "v0_process",
    "v16_process",
    "resources_rc",
    "header_process",
    "iva_format_process",
    "underline_bank_process",
    "underline_aux_process",
]

# Dependencias complejas
hidden_imports += collect_submodules("torch")
hidden_imports += collect_submodules("torchvision")
hidden_imports += collect_submodules("easyocr")
hidden_imports += collect_submodules("cv2")

datas = []
datas += collect_data_files("easyocr")
datas += collect_data_files("PyQt5")
datas += collect_data_files("qtawesome")
datas += collect_data_files("torch")
datas += collect_data_files("torchvision")

binaries = []
binaries += collect_dynamic_libs("torch")
binaries += collect_dynamic_libs("torchvision")
binaries += collect_dynamic_libs("cv2")

a = Analysis(
    ["src/interface.py"],
    pathex=["src"],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden_imports,
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
    [],
    exclude_binaries=True,
    name="FED",
    debug=True,  # pon True si quieres más diagnóstico
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # mientras depuras, déjalo en True
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon="icons/favicon.ico",
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="FED",
)
