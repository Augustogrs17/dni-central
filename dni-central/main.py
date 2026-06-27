"""
DNI Central — Metalfrio Solutions
===================================
Autor : Augusto G. Silvestrin (silvestrin)
Área  : Engenharia de Produto / DNI
"""

import multiprocessing
import os
import sys
from pathlib import Path


def _frontend_path() -> str:
    """Retorna o caminho absoluto do index.html."""
    if getattr(sys, "frozen", False):
        # Rodando como .exe (PyInstaller)
        base = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "frontend", "index.html")


def executar_gui() -> None:
    try:
        import webview
    except ImportError:
        _fallback_sem_webview()
        return

    from backend.api import API

    api    = API()
    html   = Path(_frontend_path()).read_text(encoding="utf-8")

    window = webview.create_window(
        title            = "DNI Central — Metalfrio Solutions",
        html             = html,
        js_api           = api,
        width            = 1280,
        height           = 760,
        min_size         = (960, 600),
        background_color = "#141414",
        frameless        = False,
        easy_drag        = False,
    )

    webview.start(debug=False, http_server=False)


def _fallback_sem_webview() -> None:
    try:
        import tkinter as tk
        from tkinter import messagebox
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror(
            "DNI Central",
            "pywebview não encontrado.\n\n"
            "Reconstrua o executável via GitHub Actions.",
        )
        root.destroy()
    except Exception:
        print("ERRO: pywebview não instalado.")


if __name__ == "__main__":
    multiprocessing.freeze_support()
    executar_gui()
