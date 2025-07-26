#!/usr/bin/env python3
"""
Script para formatear todo el código Python usando black
"""

import subprocess
import sys
from pathlib import Path


def run_black():
    """Ejecuta black en todo el proyecto"""
    print("🔧 Formateando código con black...")

    try:
        # Ejecutar black en el directorio actual
        result = subprocess.run(
            ["uv", "run", "black", "."], capture_output=True, text=True, check=True
        )

        print("✅ Formateo completado exitosamente!")
        if result.stdout:
            print(f"Salida: {result.stdout}")

    except subprocess.CalledProcessError as e:
        print(f"❌ Error al formatear código: {e}")
        if e.stdout:
            print(f"Salida: {e.stdout}")
        if e.stderr:
            print(f"Error: {e.stderr}")
        sys.exit(1)


def check_formatting():
    """Verifica si el código está bien formateado"""
    print("🔍 Verificando formato del código...")

    try:
        result = subprocess.run(
            ["uv", "run", "black", "--check", "."],
            capture_output=True,
            text=True,
            check=True,
        )

        print("✅ El código está correctamente formateado!")
        return True

    except subprocess.CalledProcessError as e:
        print("⚠️  El código necesita ser formateado:")
        if e.stdout:
            print(f"Salida: {e.stdout}")
        return False


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--check":
        check_formatting()
    else:
        run_black()
