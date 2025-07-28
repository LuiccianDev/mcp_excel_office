import json
from pathlib import Path
from typing import Any, cast

CONFIG_PATH_JSON = Path(__file__).resolve().parent / "mcp_excel_config.json"


def load_config() -> dict[str, Any]:
    if not CONFIG_PATH_JSON.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {CONFIG_PATH_JSON}")

    try:
        with open(CONFIG_PATH_JSON, encoding="utf-8") as f:
            return cast(dict[str, Any], json.load(f))
    except json.JSONDecodeError as e:
        raise ValueError(f"Error en el archivo JSON: {e}") from e
    except Exception as e:
        raise Exception(f"Error al cargar la configuración: {e}") from e
