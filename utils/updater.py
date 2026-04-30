import hashlib
import os
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
TIMEOUT_REQUEST = 30


def _normalizar_versao(versao: str) -> str:
    versao = str(versao or "").strip()
    if not versao:
        return ""
    return versao if versao.startswith("v") else f"v{versao}"


def _asset_por_nome(assets: list[dict], nome: str) -> dict | None:
    for asset in assets or []:
        if asset.get("name") == nome:
            return asset
    return None


def _asset_executavel(assets: list[dict]) -> dict | None:
    if not assets:
        return None

    # Preferência: asset com o nome oficial esperado pelo app.
    asset_oficial = _asset_por_nome(assets, EXECUTABLE_NAME)
    if asset_oficial:
        return asset_oficial

    # Fallback seguro: se houver apenas um .exe na release, usa ele.
    executaveis = [
        asset for asset in assets
        if str(asset.get("name", "")).lower().endswith(".exe")
    ]

    if len(executaveis) == 1:
        return executaveis[0]

    return None


def _baixar_arquivo(url: str, destino: Path, tamanho_esperado: int | None = None) -> bool:
    destino.parent.mkdir(parents=True, exist_ok=True)

    try:
        logger.info(f"Baixando atualização para: {destino}")

        with requests.get(url, stream=True, timeout=TIMEOUT_REQUEST) as resposta:
            resposta.raise_for_status()

            with open(destino, "wb") as arquivo:
                for bloco in resposta.iter_content(chunk_size=1024 * 1024):
                    if bloco:
                        arquivo.write(bloco)

        if not destino.exists() or destino.stat().st_size == 0:
            logger.error("Download da atualização falhou: arquivo não foi criado ou está vazio.")
            return False

        if tamanho_esperado and destino.stat().st_size != int(tamanho_esperado):
            logger.error(
                f"Tamanho da atualização inválido. "
                f"Esperado: {tamanho_esperado} bytes | Baixado: {destino.stat().st_size} bytes"
            )
            return False

        logger.info("Download da atualização concluído com sucesso.")
        return True

    except Exception as e:
        logger.error(f"Erro ao baixar atualização: {e}")
        return False


def _calcular_sha256(caminho: Path) -> str:
    sha256 = hashlib.sha256()

    with open(caminho, "rb") as arquivo:
        for bloco in iter(lambda: arquivo.read(1024 * 1024), b""):
            sha256.update(bloco)

    return sha256.hexdigest().lower()


def _extrair_hash_esperado(conteudo: str) -> str:
    # Aceita tanto arquivo com apenas o hash quanto "HASH  RAE_Turbo.exe".
    partes = str(conteudo or "").strip().split()
    if not partes:
        return ""
    return partes[0].strip().lower()


def _validar_sha256(caminho_exe: Path, hash_esperado: str) -> bool:
    if not hash_esperado:
        return True

    try:
        hash_calculado = _calcular_sha256(caminho_exe)

        if hash_calculado != hash_esperado:
            logger.error(
                f"SHA-256 inválido. Esperado: {hash_esperado} | Calculado: {hash_calculado}"
            )
            return False

        logger.info("SHA-256 da atualização validado com sucesso.")
        return True

    except Exception as e:
        logger.error(f"Erro ao validar SHA-256: {e}")
        return False


def verificar_atualizacao():
    try:
        url_api = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resposta = requests.get(url_api, timeout=TIMEOUT_REQUEST)
        resposta.raise_for_status()

        dados = resposta.json()

        versao_mais_recente = _normalizar_versao(dados.get("tag_name"))
        versao_atual = _normalizar_versao(VERSAO_ATUAL)

        logger.info(f"Versão atual: {versao_atual} | Versão mais recente: {versao_mais_recente}")

        if not versao_mais_recente or versao_mais_recente == versao_atual:
            return

        assets = dados.get("assets", [])
        asset_exe = _asset_executavel(assets)

        if not asset_exe:
            logger.warning("Atualização encontrada, mas nenhum executável válido foi localizado na release.")
            messagebox.showwarning(
                "Atualização",
                "Existe uma nova versão, mas o executável da atualização não foi encontrado na release."
            )
            return

        resposta_usuario = messagebox.askyesno(
            "Atualização disponível",
            f"Uma nova versão ({versao_mais_recente}) foi encontrada.\n\n"
            "Deseja atualizar agora?"
        )

        if not resposta_usuario:
            return

        asset_hash = _asset_por_nome(assets, f"{EXECUTABLE_NAME}.sha256")
        aplicar_atualizacao(asset_exe, asset_hash)

    except requests.RequestException as e:
        logger.warning(f"Não foi possível verificar atualizações: {e}")
    except Exception as e:
        logger.error(f"Erro inesperado ao verificar atualizações: {e}")
        messagebox.showerror(
            "Atualização",
            f"Ocorreu um erro inesperado ao verificar/aplicar atualização.\n\n{e}"
        )


def aplicar_atualizacao(asset_exe: dict, asset_hash: dict | None = None):
    """
    Baixa o novo executável para um arquivo temporário e cria um .bat para:
    1. aguardar o app atual encerrar;
    2. substituir o executável;
    3. reiniciar com PYINSTALLER_RESET_ENVIRONMENT=1.
    """
    try:
        exe_atual = Path(sys.executable).resolve()
        pasta_app = Path(executable_dir()).resolve()
        nome_exe = exe_atual.name

        # Se estiver rodando via python durante desenvolvimento, não tenta substituir python.exe.
        if nome_exe.lower() in {"python.exe", "pythonw.exe"}:
            messagebox.showinfo(
                "Atualização",
                "Atualização automática só funciona no executável final do RAE Turbo.\n\n"
                "Você está rodando pelo Python durante desenvolvimento."
            )
            return

        caminho_temp = Path(update_temp_path()).resolve()
        bat_path = Path(update_bat_path()).resolve()

        url_exe = asset_exe.get("browser_download_url")
        tamanho_exe = asset_exe.get("size")

        if not url_exe:
            raise RuntimeError("URL do executável da atualização não encontrada.")

        logger.info(f"Iniciando atualização. Asset: {asset_exe.get('name')} | URL: {url_exe}")

        if not _baixar_arquivo(url_exe, caminho_temp, tamanho_exe):
            messagebox.showerror(
                "Atualização",
                "Não foi possível baixar a nova versão.\n\n"
                "Verifique sua conexão e tente novamente."
            )
            return

        # Validação opcional do SHA-256, caso a release tenha o asset .sha256.
        if asset_hash and asset_hash.get("browser_download_url"):
            try:
                resposta_hash = requests.get(asset_hash["browser_download_url"], timeout=TIMEOUT_REQUEST)
                resposta_hash.raise_for_status()
                hash_esperado = _extrair_hash_esperado(resposta_hash.text)

                if not _validar_sha256(caminho_temp, hash_esperado):
                    messagebox.showerror(
                        "Atualização",
                        "A nova versão foi baixada, mas falhou na validação de integridade.\n\n"
                        "A atualização foi cancelada por segurança."
                    )
                    try:
                        caminho_temp.unlink(missing_ok=True)
                    except Exception:
                        pass
                    return

            except Exception as e:
                logger.warning(f"Não foi possível validar SHA-256 da atualização: {e}")

        conteudo_bat = f"""@echo off
chcp 65001 > nul
cd /d "{pasta_app}"

echo Aguardando o RAE Turbo encerrar...
timeout /t 5 /nobreak > nul

:aguardar_fechamento
tasklist /FI "IMAGENAME eq {nome_exe}" | find /I "{nome_exe}" > nul
if not errorlevel 1 (
    timeout /t 1 /nobreak > nul
    goto aguardar_fechamento
)

echo Aplicando atualização...
move /Y "{caminho_temp}" "{exe_atual}"
if errorlevel 1 (
    echo Falha ao substituir o executavel.
    pause
    exit /b 1
)

echo Aguardando finalização...
timeout /t 2 /nobreak > nul

echo Reiniciando RAE Turbo...
set PYINSTALLER_RESET_ENVIRONMENT=1
start "" "{exe_atual}"

timeout /t 1 /nobreak > nul
del "%~f0"
"""

        bat_path.write_text(conteudo_bat, encoding="utf-8")
        logger.info(f"Script de atualização criado em: {bat_path}")

        subprocess.Popen(
            ["cmd", "/c", str(bat_path)],
            cwd=str(pasta_app),
            creationflags=subprocess.CREATE_NO_WINDOW,
        )

        logger.info("Encerrando aplicativo atual para aplicar atualização.")
        sys.exit(0)

    except Exception as e:
        logger.error(f"Erro ao aplicar atualização: {e}")
        messagebox.showerror(
            "Atualização",
            f"Não foi possível aplicar a atualização.\n\n{e}"
        )
