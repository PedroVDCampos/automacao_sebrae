import logging
import os

def configurar_logger():
    pasta = os.path.join(os.path.expanduser("~"), "AppData", "Local", "RAETurbo")
    os.makedirs(pasta, exist_ok=True)
    logger = logging.getLogger("RAETurbo")
    logger.setLevel(logging.INFO)
    if not logger.handlers:
        h = logging.FileHandler(os.path.join(pasta, "rae_turbo_execucao.log"), encoding="utf-8")
        h.setFormatter(logging.Formatter("%(asctime)s - [%(levelname)s] - %(message)s"))
        logger.addHandler(h)
    return logger
