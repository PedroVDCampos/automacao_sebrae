import os
import sys
from pathlib import Path

APP_NAME = "RAETurbo"
EXECUTABLE_NAME = "RAE_Turbo.exe"

def is_frozen() -> bool:
    return getattr(sys, "frozen", False)

def app_root() -> Path:
    if is_frozen():
        return Path(getattr(sys, "_MEIPASS"))
    return Path(__file__).resolve().parents[1]

def executable_dir() -> Path:
    if is_frozen():
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]

def appdata_dir() -> Path:
    base = os.getenv("LOCALAPPDATA")
    if base:
        path = Path(base) / APP_NAME
    else:
        path = Path.home() / "AppData" / "Local" / APP_NAME
    path.mkdir(parents=True, exist_ok=True)
    return path

def resource_path(relative_path: str) -> str:
    return str(app_root() / relative_path)

def data_path(filename: str) -> Path:
    return app_root() / "data" / filename

def config_path() -> Path:
    return appdata_dir() / "config_unidade.json"

def log_path() -> Path:
    return appdata_dir() / "rae_turbo_execucao.log"

def update_temp_path() -> Path:
    return executable_dir() / "update_temporario_download.exe"

def update_bat_path() -> Path:
    return executable_dir() / "atualizar_rae.bat"
