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
VERSAO_ATUAL = "v1.1.4"
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
        # Pega o nome exato do arquivo que o usuário clicou (ex: "RAE_Turbo (1).exe")
        nome_executavel_atual = os.path.basename(sys.executable)
        
        # O arquivo novo temporário
        novo_exe_tmp = "update_temporario_download.exe"
        
        # Faz o download
        resposta = requests.get(url_download, stream=True)
        with open(novo_exe_tmp, 'wb') as f:
            for chunk in resposta.iter_content(chunk_size=8192):
                f.write(chunk)
        
        script_bat = "atualizar_rae.bat"
        
        # O script agora usa a variável 'nome_executavel_atual' dinamicamente
        # e renomeia o arquivo novo para ter O MESMO NOME que o usuário já usava.
        conteudo_bat = f"""@echo off
:tentar_deletar
timeout /t 1 /nobreak > NUL
del /f /q "{nome_executavel_atual}"
if exist "{nome_executavel_atual}" goto tentar_deletar

ren "{novo_exe_tmp}" "{nome_executavel_atual}"

set _MEIPASS2=
set _MEIPASS=

start "" "{nome_executavel_atual}"
del "%~f0"
"""
        with open(script_bat, "w") as f:
            f.write(conteudo_bat)
            
        subprocess.Popen(script_bat, shell=True, creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit()
        
    except Exception as e:
        messagebox.showerror("Erro no Update", f"Não foi possível atualizar: {e}")