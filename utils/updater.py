import sys
import subprocess
from pathlib import Path

import requests
from tkinter import messagebox

from utils.logger import configurar_logger
from utils.paths import EXECUTABLE_NAME, executable_dir, update_temp_path, update_bat_path
from version import VERSAO_ATUAL

logger = configurar_logger()

GITHUB_REPO = "PedroVDCampos/automacao_sebrae"
TIMEOUT_REQUEST = 10

def _normalizar_versao(versao: str) -> str:
    versao = str(versao or "").strip()
    if not versao:
        return ""
    return versao if versao.startswith("v") else f"v{versao}"

def _asset_executavel(assets: list[dict]) -> dict | None:
    if not assets:
        return None
    for asset in assets:
        if asset.get("name") == EXECUTABLE_NAME:
            return asset
    executaveis = [a for a in assets if str(a.get("name", "")).lower().endswith(".exe")]
    if len(executaveis) == 1:
        return executaveis[0]
    return None

def verificar_atualizacao():
    try:
        url_api = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resposta = requests.get(url_api, timeout=TIMEOUT_REQUEST)
        resposta.raise_for_status()
        dados = resposta.json()
        versao_mais_recente = _normalizar_versao(dados.get("tag_name"))
        versao_atual = _normalizar_versao(VERSAO_ATUAL)
        if not versao_mais_recente or versao_mais_recente == versao_atual:
            return
        asset = _asset_executavel(dados.get("assets", []))
        if not asset:
            logger.warning("Atualização encontrada, mas nenhum executável válido foi localizado na release.")
            return
        resposta_usuario = messagebox.askyesno(
            "Atualização disponível",
            f"Uma nova versão ({versao_mais_recente}) foi encontrada.\n\nDeseja atualizar agora?"
        )
        if resposta_usuario:
            aplicar_atualizacao(asset["browser_download_url"], asset.get("size"))
    except requests.RequestException as e:
        logger.warning(f"Não foi possível verificar atualizações: {e}")
    except Exception as e:
        logger.error(f"Erro inesperado ao verificar atualizações: {e}")

def aplicar_atualizacao(caminho_novo_exe: str):
    exe_atual = Path(sys.executable).resolve()
    pasta_app = exe_atual.parent
    nome_exe = exe_atual.name

    caminho_novo_exe = Path(caminho_novo_exe).resolve()
    bat_path = pasta_app / "atualizar_rae_turbo.bat"

    conteudo_bat = f"""@echo off
chcp 65001 > nul
cd /d "{pasta_app}"

echo Aguardando o RAE Turbo encerrar...
timeout /t 5 /nobreak > nul

echo Aplicando atualização...
move /Y "{caminho_novo_exe}" "{exe_atual}"

echo Aguardando finalização...
timeout /t 2 /nobreak > nul

echo Reiniciando RAE Turbo...
set PYINSTALLER_RESET_ENVIRONMENT=1
start "" "{exe_atual}"

timeout /t 1 /nobreak > nul
del "%~f0"
"""

    bat_path.write_text(conteudo_bat, encoding="utf-8")

    subprocess.Popen(
        ["cmd", "/c", str(bat_path)],
        creationflags=subprocess.CREATE_NO_WINDOW
    )

    sys.exit(0)
