import hashlib
import os
import subprocess
import sys

import requests

from utils.logger import configurar_logger
from utils.paths import caminho_update_script, caminho_update_temporario
from version import VERSAO_ATUAL

logger = configurar_logger()
GITHUB_REPO = "PedroVDCampos/automacao_sebrae"
NOME_EXE = "RAE_Turbo.exe"


def _normalizar(v):
    return str(v or "").strip().lstrip("vV")


def _versao(v):
    partes = []
    for parte in _normalizar(v).split("."):
        numero = ""
        for char in parte:
            if char.isdigit():
                numero += char
            else:
                break
        partes.append(int(numero or 0))
    while len(partes) < 3:
        partes.append(0)
    return tuple(partes[:3])


def _sha256(caminho):
    h = hashlib.sha256()
    with open(caminho, "rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            h.update(bloco)
    return h.hexdigest()


def verificar_atualizacao(callback=None):
    """Consulta a última Release do GitHub sem abrir janelas fora da thread principal."""
    try:
        resposta = requests.get(
            f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",
            timeout=8,
            headers={"Accept": "application/vnd.github+json"},
        )
        resposta.raise_for_status()
        dados = resposta.json()
        remota = dados.get("tag_name")

        if not remota or _versao(remota) <= _versao(VERSAO_ATUAL):
            return None

        asset = next(
            (
                a
                for a in dados.get("assets", [])
                if a.get("name", "").lower() == NOME_EXE.lower()
            ),
            None,
        )
        if not asset or not asset.get("browser_download_url"):
            logger.warning("Release %s sem asset %s", remota, NOME_EXE)
            return None

        info = {
            "versao": remota,
            "url": asset["browser_download_url"],
            "digest": asset.get("digest", ""),
        }
        if callback:
            callback(info)
        return info
    except Exception as erro:
        logger.error("Erro ao verificar atualizações: %s", erro)
        return None


def baixar_atualizacao(info):
    """Baixa a nova versão e valida o SHA-256 publicado pelo GitHub quando disponível."""
    tmp = caminho_update_temporario()
    os.makedirs(os.path.dirname(tmp), exist_ok=True)

    resposta = requests.get(info["url"], stream=True, timeout=120)
    resposta.raise_for_status()

    with open(tmp, "wb") as arquivo:
        for bloco in resposta.iter_content(1024 * 1024):
            if bloco:
                arquivo.write(bloco)

    if os.path.getsize(tmp) < 100 * 1024:
        raise RuntimeError("O arquivo baixado parece inválido ou incompleto.")

    sha_local = _sha256(tmp)
    digest_github = str(info.get("digest") or "")
    if digest_github.startswith("sha256:"):
        sha_esperado = digest_github.split(":", 1)[1].strip().lower()
        if sha_local.lower() != sha_esperado:
            os.remove(tmp)
            raise RuntimeError("A validação SHA-256 da atualização falhou.")

    logger.info("Atualização %s baixada. SHA-256: %s", info["versao"], sha_local)
    return tmp


def instalar_atualizacao(caminho_exe_novo):
    """Troca o executável somente depois que o processo atual for encerrado."""
    if not getattr(sys, "frozen", False):
        raise RuntimeError("A atualização automática funciona no executável compilado.")

    exe = os.path.abspath(sys.executable)
    pasta = os.path.dirname(exe)
    nome = os.path.basename(exe)
    novo = os.path.join(pasta, "_RAE_Turbo_update.exe")

    if os.path.exists(novo):
        os.remove(novo)
    os.replace(caminho_exe_novo, novo)

    script = caminho_update_script()
    bat = f'''@echo off
setlocal
set "APP={nome}"
set "NEW=_RAE_Turbo_update.exe"
cd /d "{pasta}"
:wait
timeout /t 1 /nobreak >nul
tasklist /FI "IMAGENAME eq %APP%" | find /I "%APP%" >nul
if not errorlevel 1 goto wait
if not exist "%NEW%" exit /b 1
move /Y "%NEW%" "%APP%" >nul
start "" "%APP%"
del "%~f0"
'''
    with open(script, "w", encoding="utf-8") as arquivo:
        arquivo.write(bat)

    subprocess.Popen(
        ["cmd", "/c", script],
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    sys.exit(0)
