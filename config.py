"""
config.py
Manejo de API key y configuracion de la app.
La configuracion se guarda en %APPDATA%/GeneradorAMR/config.json
"""

import json
import os

_CONFIG_DIR  = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "GeneradorAMR")
_CONFIG_FILE = os.path.join(_CONFIG_DIR, "config.json")


def load_config() -> dict:
    if not os.path.isfile(_CONFIG_FILE):
        return {}
    try:
        with open(_CONFIG_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_config(data: dict) -> None:
    os.makedirs(_CONFIG_DIR, exist_ok=True)
    with open(_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def get_api_key() -> str:
    cfg = load_config()
    return cfg.get("claude_api_key", "") or cfg.get("gemini_api_key", "")


def set_api_key(key: str) -> None:
    cfg = load_config()
    cfg["claude_api_key"] = key.strip()
    save_config(cfg)
