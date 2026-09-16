import os, sys, subprocess, hashlib
import requests
from tkinter import messagebox
from utils.logger import configurar_logger
from utils.paths import caminho_update_temporario, caminho_update_script
from version import VERSAO_ATUAL

logger=configurar_logger()
GITHUB_REPO="PedroVDCampos/automacao_sebrae"
NOME_EXE="RAE_Turbo.exe"

def _normalizar(v): return str(v or "").strip().lstrip("vV")
def _sha256(p):
    h=hashlib.sha256()
    with open(p,"rb") as f:
        for b in iter(lambda:f.read(1024*1024),b""): h.update(b)
    return h.hexdigest()

def verificar_atualizacao():
    try:
        r=requests.get(f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest",timeout=8,headers={"Accept":"application/vnd.github+json"})
        r.raise_for_status(); dados=r.json(); remota=dados.get("tag_name")
        if not remota or _normalizar(remota)==_normalizar(VERSAO_ATUAL): return
        asset=next((a for a in dados.get("assets",[]) if a.get("name","").lower()==NOME_EXE.lower()),None)
        if not asset:
            logger.warning("Release %s sem asset %s",remota,NOME_EXE); return
        if messagebox.askyesno("Atualização disponível",f"Uma nova versão ({remota}) do RAE Turbo foi encontrada.\n\nDeseja atualizar agora?"):
            aplicar_atualizacao(asset["browser_download_url"])
    except Exception as e: logger.error("Erro ao verificar atualizações: %s",e)

def aplicar_atualizacao(url_download):
    try:
        tmp=caminho_update_temporario(); os.makedirs(os.path.dirname(tmp),exist_ok=True)
        r=requests.get(url_download,stream=True,timeout=60); r.raise_for_status()
        with open(tmp,"wb") as f:
            for b in r.iter_content(1024*1024):
                if b:f.write(b)
        if os.path.getsize(tmp)<100*1024: raise RuntimeError("O arquivo baixado parece inválido ou incompleto.")
        logger.info("Atualização baixada. SHA-256: %s",_sha256(tmp))
        if not getattr(sys,"frozen",False):
            return messagebox.showinfo("Atualização","A atualização automática funciona no executável compilado.")
        exe=os.path.abspath(sys.executable); pasta=os.path.dirname(exe); nome=os.path.basename(exe); novo=os.path.join(pasta,"_RAE_Turbo_update.exe")
        if os.path.exists(novo): os.remove(novo)
        os.replace(tmp,novo)
        script=caminho_update_script()
        bat=f'''@echo off
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
        with open(script,"w",encoding="utf-8") as f:f.write(bat)
        subprocess.Popen(["cmd","/c",script],creationflags=subprocess.CREATE_NO_WINDOW)
        sys.exit(0)
    except Exception as e:
        logger.exception("Erro ao aplicar atualização")
        messagebox.showerror("Erro na atualização",f"Não foi possível atualizar o RAE Turbo.\n\n{e}")
