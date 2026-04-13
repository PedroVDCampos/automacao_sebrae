import os
import sys
import subprocess
import threading
import requests
from tkinter import messagebox

GITHUB_REPO = "PedroVDCampos/automacao_sebrae"



def verificar_atualizacao():
    """
    Dispara a verificação de atualização em segundo plano,
    sem travar a interface gráfica.
    """
    thread = threading.Thread(target=_checar_versao, daemon=True)
    thread.start()


def _checar_versao():
    from version import VERSAO_ATUAL

    try:
        url = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resposta = requests.get(url, timeout=8)

        if resposta.status_code != 200:
            return  # Silencioso: sem internet ou repo inexistente

        release = resposta.json()
        versao_nova = release["tag_name"].lstrip("v")

        if versao_nova > VERSAO_ATUAL:
            confirmado = messagebox.askyesno(
                "🔄 Atualização Disponível",
                f"Uma nova versão do programa está disponível!\n\n"
                f"  Versão atual:  {VERSAO_ATUAL}\n"
                f"  Nova versão:   {versao_nova}\n\n"
                f"Deseja atualizar agora? O programa será reiniciado automaticamente.",
            )
            if confirmado:
                # Pega o primeiro .exe da release
                asset = next(
                    (a for a in release.get("assets", []) if a["name"].endswith(".exe")),
                    None,
                )
                if asset:
                    _baixar_e_instalar(asset["browser_download_url"], asset["name"])
                else:
                    messagebox.showerror(
                        "Erro",
                        "Não foi encontrado nenhum arquivo .exe nessa release.\n"
                        "Entre em contato com o desenvolvedor.",
                    )

    except requests.exceptions.ConnectionError:
        pass  # Sem internet — ignora silenciosamente
    except Exception as e:
        print(f"[Updater] Erro inesperado ao verificar atualização: {e}")


def _baixar_e_instalar(url_download: str, nome_arquivo: str):
    """
    Baixa o novo .exe e cria um script .bat que:
      1. Aguarda o programa atual fechar
      2. Substitui o .exe antigo pelo novo
      3. Inicia o programa atualizado
      4. Se auto-apaga
    """
    try:
        messagebox.showinfo(
            "Baixando...",
            "A atualização está sendo baixada.\n"
            "O programa será reiniciado automaticamente ao terminar.",
        )

        exe_atual = sys.executable
        pasta_atual = os.path.dirname(exe_atual)
        caminho_novo_exe = os.path.join(pasta_atual, f"_novo_{nome_arquivo}")

        # Download em streaming para não estourar memória
        with requests.get(url_download, stream=True, timeout=60) as r:
            r.raise_for_status()
            with open(caminho_novo_exe, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    f.write(chunk)

        # Script .bat que roda DEPOIS que o processo principal fechar
        conteudo_bat = (
            "@echo off\n"
            "timeout /t 2 /nobreak > nul\n"
            f'move /Y "{caminho_novo_exe}" "{exe_atual}"\n'
            f'start "" "{exe_atual}"\n'
            'del "%~f0"\n'
        )
        caminho_bat = os.path.join(pasta_atual, "_atualizar.bat")
        with open(caminho_bat, "w", encoding="utf-8") as f:
            f.write(conteudo_bat)

        # Inicia o .bat em segundo plano e fecha o programa atual
        subprocess.Popen(
            caminho_bat,
            shell=True,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        sys.exit()

    except Exception as e:
        messagebox.showerror(
            "Erro na Atualização",
            f"Não foi possível concluir a atualização.\n\nMotivo: {e}\n\n"
            "Tente novamente mais tarde ou atualize manualmente.",
        )
