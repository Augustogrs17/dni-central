"""
config.py
=========
Gerencia o config.json persistido ao lado do executável (ou no home do
usuário, caso a pasta do .exe não seja gravável).
"""

import os
import sys
import json
from pathlib import Path

CONFIG_FILENAME = "dni_central_config.json"


def _config_path() -> str:
    try:
        if getattr(sys, "frozen", False):
            base = os.path.dirname(sys.executable)
        else:
            base = os.path.dirname(os.path.abspath(__file__))
            base = os.path.dirname(base)  # sobe para raiz do projeto

        # Testa se a pasta é gravável
        teste = os.path.join(base, ".write_test")
        try:
            with open(teste, "w") as f:
                f.write("ok")
            os.remove(teste)
        except Exception:
            base = str(Path.home())
    except Exception:
        base = str(Path.home())

    return os.path.join(base, CONFIG_FILENAME)


def carregar() -> dict:
    path = _config_path()
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def salvar(cfg: dict) -> None:
    path = _config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass
