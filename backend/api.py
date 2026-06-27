"""
api.py
======
Bridge Python ↔ JavaScript via pywebview.
Step 1: configurações, seleção de pasta/arquivo e leitura de projetos.
"""

import os
from backend import config as cfg_mod


class API:
    """Métodos expostos ao frontend via window.pywebview.api.*"""

    # ── Config ────────────────────────────────────────────────────────

    def carregar_config(self) -> dict:
        return cfg_mod.carregar()

    def salvar_config(self, dados: dict) -> None:
        cfg_mod.salvar(dados)

    def get_version(self) -> str:
        return "v0.1.0 · DNI Central"

    # ── Diálogos de seleção ───────────────────────────────────────────

    def pick_folder(self) -> str | None:
        try:
            import webview
            dirs = webview.windows[0].create_file_dialog(webview.FOLDER_DIALOG)
            return dirs[0] if dirs else None
        except Exception:
            return None

    def pick_file(self, extensoes: str = "Excel (*.xlsx;*.xlsm)") -> str | None:
        try:
            import webview
            files = webview.windows[0].create_file_dialog(
                webview.OPEN_DIALOG,
                file_types=(extensoes, "All files (*.*)"),
            )
            return files[0] if files else None
        except Exception:
            return None

    # ── Leitura de projetos (Step 1) ──────────────────────────────────

    def listar_projetos(self, pasta: str) -> dict:
        """
        Varre a pasta informada, lê metadados básicos de cada .xlsm
        (sem extrair itens DI ainda) e retorna lista para exibição.
        """
        from backend.extrator import listar_projetos
        try:
            projetos = listar_projetos(pasta)
            return {"ok": True, "projetos": projetos}
        except Exception as e:
            return {"ok": False, "erro": str(e), "projetos": []}

    def gerar_planilha(self, pasta_origem: str, pasta_destino: str) -> dict:
        """
        Lê todos os projetos da pasta, calcula o resumo e gera o Excel.
        """
        from backend.extrator import listar_projetos
        from backend.exportador import calcular_resumo, gerar_excel
        import os
        try:
            if not os.path.isdir(pasta_destino):
                os.makedirs(pasta_destino, exist_ok=True)
            projetos = listar_projetos(pasta_origem)
            if not projetos:
                return {"ok": False, "erro": "Nenhum projeto encontrado na pasta."}
            resumo  = calcular_resumo(projetos)
            caminho = gerar_excel(projetos, resumo, pasta_destino)
            return {
                "ok":      True,
                "caminho": caminho,
                "nome":    os.path.basename(caminho),
                "total":   len(projetos),
            }
        except Exception as e:
            return {"ok": False, "erro": str(e)}
