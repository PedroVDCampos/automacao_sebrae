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
VERSAO_ATUAL = "v1.0.0" 
NOME_EXE = "RAE_Turbo.exe" 

def verificar_atualizacao():
    try:
        url_api = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
        resposta = requests.get(url_api, timeout=5)
        
        if resposta.status_code == 200:
            dados = resposta.json()
            versao_mais_recente = dados.get("tag_name")
            
            # Se a versão do GitHub for diferente da local
            if versao_mais_recente and versao_mais_recente != VERSAO_ATUAL:
                # Pegamos o primeiro arquivo (asset) da release
                baixar_url = dados["assets"][0]["browser_download_url"]
                
                escolha = messagebox.askyesno(
                    "RAE Turbo - Atualização", 
                    f"Nova versão disponível: {versao_mais_recente}\nDeseja atualizar agora?"
                )
                
                if escolha:
                    aplicar_atualizacao(baixar_url)
                    
    except Exception as e:
        logger.error(f"Erro crítico na verificação de update: {e}")

def aplicar_atualizacao(url_download):
    try:
        novo_exe_tmp = "RAE_Turbo_Novo.exe"
        
        # Download com stream para não travar a memória
        resposta = requests.get(url_download, stream=True)
        with open(novo_exe_tmp, 'wb') as f:
            for chunk in resposta.iter_content(chunk_size=8192):
                f.write(chunk)
        
        # O script BAT é o nosso "agente externo" que trabalha enquanto o Python fecha
        script_bat = "atualizar_rae.bat"
        
        # Lógica do BAT:
        # 1. Espera 3 segundos (tempo para o Python dar sys.exit)
        # 2. Tenta apagar o velho (del /f /q força a exclusão)
        # 3. Renomeia o novo para o nome oficial
        # 4. Abre o novo
        # 5. Se apaga sozinho (%~f0)
        conteudo_bat = f"""@echo off
timeout /t 3 /nobreak > NUL
if exist "{NOME_EXE}" del /f /q "{NOME_EXE}"
ren "{novo_exe_tmp}" "{NOME_EXE}"
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