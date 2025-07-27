from pathlib import Path
import json

CONFIG_PATH_JSON = Path(__file__).resolve().parent / "mcp_excel_config.json"

def load_config() -> dict:
    if not CONFIG_PATH_JSON.exists():
        raise FileNotFoundError(f"Archivo no encontrado: {CONFIG_PATH_JSON}")

    try:
        with open(CONFIG_PATH_JSON, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(f"Error en el archivo JSON: {e}")

