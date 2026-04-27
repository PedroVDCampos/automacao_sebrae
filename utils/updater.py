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

def aplicar_atualizacao(url_download: str, tamanho_esperado: int | None = None):
    try:
        pasta_execucao = executable_dir()
        nome_executavel_atual = Path(sys.executable).name if getattr(sys, "frozen", False) else EXECUTABLE_NAME
        novo_exe_tmp = update_temp_path()
        with requests.get(url_download, stream=True, timeout=TIMEOUT_REQUEST) as resposta:
            resposta.raise_for_status()
            with open(novo_exe_tmp, "wb") as f:
                for chunk in resposta.iter_content(chunk_size=1024 * 1024):
                    if chunk:
                        f.write(chunk)
        tamanho_baixado = novo_exe_tmp.stat().st_size
        if tamanho_baixado <= 0:
            raise RuntimeError("O arquivo baixado está vazio.")
        if tamanho_esperado and abs(tamanho_baixado - int(tamanho_esperado)) > 1024:
            raise RuntimeError("O tamanho do arquivo baixado não confere com o informado pelo GitHub.")
        script_bat = update_bat_path()
        conteudo_bat = f'''@echo off
cd /d "{pasta_execucao}"
:tentar_deletar
timeout /t 1 /nobreak > NUL
del /f /q "{nome_executavel_atual}"
if exist "{nome_executavel_atual}" goto tentar_deletar

ren "{novo_exe_tmp.name}" "{nome_executavel_atual}"

set _MEIPASS2=
set _MEIPASS=

start "" "{nome_executavel_atual}"
del "%~f0"
'''
        script_bat.write_text(conteudo_bat, encoding="utf-8")
        subprocess.Popen(str(script_bat), shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit()
    except Exception as e:
        logger.error(f"Erro ao aplicar atualização: {e}")
        messagebox.showerror("Erro na atualização", f"Não foi possível atualizar o RAE Turbo.\n\nMotivo: {e}")
