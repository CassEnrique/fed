# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ["src/interface.py"],  # Punto de entrada
    pathex=[
        "src"
    ],  # MUY IMPORTANTE: Le dice a PyInstaller que busque módulos dentro de la carpeta 'src'
    binaries=[],
    datas=[],  # Deja esto vacío para los archivos .py
    hiddenimports=[  # Forzamos la inclusión de tus módulos propios
        "iva_process",
        "v0_process",
        "v16_process",
        "resources_rc",
        "header_process",
        "iva_format_process",
        "underline_bank_process",
        "underline_aux_process",
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
    name="FED",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Si quieres ver errores en consola al probar, cámbialo a True
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=["icons/favicon.ico"],
)
