#!/usr/bin/env python3
"""
Nuitka Builder mejorado: seguro, modular y eficiente
"""

import os
import shutil
import subprocess
import sys
from pathlib import Path

# ================= CONFIGURACIÓN =================
PROJECT_ROOT = Path(__file__).parent
BUILD_DIR = PROJECT_ROOT / "dist"
CACHE_DIR = PROJECT_ROOT / ".nuitka_cache"
ICONS_DIR = PROJECT_ROOT / "icons"
SOURCES_DIR = PROJECT_ROOT / "src"
INTERFACE_FILE = SOURCES_DIR / "interface.py"
RESOURCES_QRC = ICONS_DIR / "sources.qrc"
RESOURCES_PY = SOURCES_DIR / "resources_rc.py"

# Dependencias críticas para verificar
IMPORTS_CHECK = [
    ("PyQt5", "GUI"),
    ("openpyxl", "Manipulación Excel"),
    ("pandas", "Procesamiento de datos"),
    ("cv2", "OpenCV para OCR"),
    ("fitz", "PyMuPDF para PDFs"),
    ("easyocr", "Reconocimiento de texto"),
    ("torch", "Modelos de OCR (CPU)"),
]


def run_command(cmd, desc=""):
    if desc:
        print(f"\n{'='*60}\n🔨 {desc}\n{'='*60}")
    print(f"📌 Ejecutando: {' '.join(cmd)}")

    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"❌ Error: {desc}")
        print(f"STDOUT:\n{result.stdout}")
        print(f"STDERR:\n{result.stderr}")
        sys.exit(1)
    print(f"✅ {desc}")
    return result


def verify_structure():
    print("\n📋 Verificando estructura del proyecto...")
    errors = []
    for path, msg in [
        (ICONS_DIR, "icons/"),
        (SOURCES_DIR, "src/"),
        (INTERFACE_FILE, "src/interface.py"),
        (RESOURCES_QRC, "icons/sources.qrc"),
    ]:
        if not path.exists():
            errors.append(f"No encontrado: {path}")

    if errors:
        for e in errors:
            print(f"❌ {e}")
        sys.exit(1)

    print("✅ Estructura OK")


def compile_resources():
    print("\n🛠️ Compilando recursos Qt...")
    run_command(
        ["pyrcc5", str(RESOURCES_QRC), "-o", str(RESOURCES_PY)],
        "Compilación de QRC"
    )
    print(f"📄 {RESOURCES_PY} actualizado")


def verify_imports():
    print("\n🔍 Verificando imports...")
    failed = []
    for module, desc in IMPORTS_CHECK:
        try:
            __import__(module)
            print(f"   ✓ {module}")
        except ImportError as e:
            failed.append(f"{module} ({desc})")
    if failed:
        print(f"❌ Faltan: {', '.join(failed)}")
        print("📌 Instala con:\n   pip install -r requirements.txt")
        sys.exit(1)
    print("✅ Todos los imports listos")


def clean_previous():
    print("\n🧹 Limpiando builds anteriores...")
    for folder in [BUILD_DIR, CACHE_DIR, "build"]:
        path = PROJECT_ROOT / folder
        if path.exists():
            shutil.rmtree(path)
            print(f"   ✓ {path} eliminado")


def build_nuitka():
    print("\n🚀 Compilando con Nuitka...")
    is_windows = sys.platform.startswith("win")

    cmd = [
        "nuitka",
        "--onefile",
        f"--output-dir={BUILD_DIR}",
        "--remove-output",
        "--jobs=4",
        "--include-data-dir", f"{ICONS_DIR}=icons",
        "--follow-imports",
        "--no-prefer-source-code",

        # Incluir módulos directos (mejor que --include-package)
        "--include-module=fitz",
        "--include-module=cv2",
        "--include-module=easyocr",

        # PyTorch (excluir CUDA)
        "--include-package=torch",
        "--include-package=torchvision",
        "--exclude-module=torch.cuda",
        "--exclude-module=nvidia.*",

        # Otros
        "--include-package=openpyxl",
        "--include-package=pandas",
        "--include-package=rapidfuzz",
        "--include-package=PIL",
    ]

    if is_windows:
        cmd.insert(2, "--plugin-enable=qt-plugins")
        icon_path = ICONS_DIR / "favicon.ico"
        if icon_path.exists():
            cmd.append(f"--windows-icon={icon_path}")

    cmd.append(str(INTERFACE_FILE))

    run_command(cmd, "Nuitka Build")

    # Verificación final
    exe_name = "interface.exe" if is_windows else "interface"
    exe_path = BUILD_DIR / exe_name
    size_mb = exe_path.stat().st_size / (1024 * 1024)
    print(f"\n🎉 ¡Éxito! Ejecutable generado:")
    print(f"📌 {exe_path}")
    print(f"📊 Tamaño: {size_mb:.1f} MB")


if __name__ == "__main__":
    try:
        verify_structure()
        compile_resources()
        verify_imports()
        clean_previous()
        build_nuitka()
    except KeyboardInterrupt:
        print("\n❌ Compilación cancelada.")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ Error: {e}")
        sys.exit(1)
