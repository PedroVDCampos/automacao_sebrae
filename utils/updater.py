import os
import sys
import time
import requests
import subprocess
from tkinter import messagebox
from utils.logger import configurar_logger
logger = configurar_logger()

# IMPORTANTE: Este nome deve ser idêntico ao --name do PyInstaller
GITHUB_REPO = "PedroVDCampos/automacao_sebrae" 
VERSAO_ATUAL = "v1.0.10"
NOME_EXE = "RAE Turbo.exe" 

def verificar_atualizacao():
    try:
        url_api = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resposta = requests.get(url_api, timeout=5)
        
        if resposta.status_code == 200:
            dados = resposta.json()
            versao_mais_recente = dados.get("tag_name")
            
            if versao_mais_recente and versao_mais_recente != VERSAO_ATUAL:
                baixar_url = dados["assets"][0]["browser_download_url"]
                
                resposta_usuario = messagebox.askyesno(
                    "Atualização Disponível", 
                    f"Uma nova versão ({versao_mais_recente}) foi encontrada!\nDeseja atualizar agora?"
                )
                
                if resposta_usuario:
                    aplicar_atualizacao(baixar_url)
        else:
            # AGORA ELE VAI TE AVISAR SE DER 404!
            logger.warning(f"O GitHub recusou o acesso ou não achou a release. Código: {resposta.status_code}")
                    
    except Exception as e:
        logger.error(f"Erro ao verificar atualizações: {e}")

def aplicar_atualizacao(url_download):
    try:
        novo_exe_tmp = "RAE_Turbo_Novo.exe"
        
        # Download com stream para não travar a memória
        resposta = requests.get(url_download, stream=True)
        with open(novo_exe_tmp, 'wb') as f:
            for chunk in resposta.iter_content(chunk_size=8192):
                f.write(chunk)
        
        # 2. Cria o script .bat definitivo (Com limpeza de memória PyInstaller)
        script_bat = "atualizar_rae.bat"
        
        conteudo_bat = f"""@echo off
:tentar_deletar
timeout /t 1 /nobreak > NUL
del /f /q "{NOME_EXE}"
if exist "{NOME_EXE}" goto tentar_deletar

ren "{novo_exe_tmp}" "{NOME_EXE}"

:: O SEGREDO SÊNIOR: Limpar a herança do PyInstaller antes de iniciar o novo
set _MEIPASS2=
set _MEIPASS=

start "" "{NOME_EXE}"
del "%~f0"
"""
        with open(script_bat, "w") as f:
            f.write(conteudo_bat)
            
        # Lança o BAT sem abrir janela de comando e fecha o programa
        subprocess.Popen(script_bat, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit()
        
    except Exception as e:
        messagebox.showerror("Erro no Update", f"Não foi possível atualizar: {e}")