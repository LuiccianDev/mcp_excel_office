#!/usr/bin/env python3
"""
Scripts comunes para el proyecto MCP Office Excel
"""
import subprocess
import sys
from pathlib import Path


def format_code():
    """Formatear todo el código con black"""
    print("🔧 Formateando código con black...")
    try:
        result = subprocess.run(["uv", "run", "black", "."], check=True)
        print("✅ Código formateado exitosamente!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error al formatear: {e}")
        sys.exit(1)


def check_format():
    """Verificar formato del código"""
    print("🔍 Verificando formato del código...")
    try:
        subprocess.run(["uv", "run", "black", "--check", "."], check=True)
        print("✅ El código está correctamente formateado!")
    except subprocess.CalledProcessError:
        print("⚠️  El código necesita ser formateado")
        sys.exit(1)


def run_server():
    """Ejecutar el servidor MCP"""
    print("🚀 Iniciando servidor MCP...")
    try:
        subprocess.run(["uv", "run", "python", "main.py"], check=True)
    except subprocess.CalledProcessError as e:
        print(f"❌ Error al ejecutar servidor: {e}")
        sys.exit(1)


def install_deps():
    """Instalar dependencias"""
    print("📦 Instalando dependencias...")
    try:
        subprocess.run(["uv", "sync"], check=True)
        print("✅ Dependencias instaladas!")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error al instalar dependencias: {e}")
        sys.exit(1)


def main():
    if len(sys.argv) < 2:
        print("Uso: python scripts.py [comando]")
        print("Comandos disponibles:")
        print("  format     - Formatear código con black")
        print("  check      - Verificar formato del código")
        print("  run        - Ejecutar servidor MCP")
        print("  install    - Instalar dependencias")
        sys.exit(1)

    command = sys.argv[1]

    if command == "format":
        format_code()
    elif command == "check":
        check_format()
    elif command == "run":
        run_server()
    elif command == "install":
        install_deps()
    else:
        print(f"❌ Comando desconocido: {command}")
        sys.exit(1)


if __name__ == "__main__":
    main()
