import os
import sys

def resource_path(relative_path):
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    return os.path.join(base, relative_path)

def pasta_appdata():
    p = os.path.join(os.path.expanduser("~"), "AppData", "Local", "RAETurbo")
    os.makedirs(p, exist_ok=True)
    return p

def caminho_historico(): return os.path.join(pasta_appdata(), "historico_execucoes.csv")
def caminho_update_temporario(): return os.path.join(pasta_appdata(), "update_temporario_download.exe")
def caminho_update_script(): return os.path.join(pasta_appdata(), "atualizar_rae_turbo.bat")
