import logging
from utils.paths import log_path

def configurar_logger():
    caminho_log = log_path()
    root_logger = logging.getLogger()
    if not root_logger.handlers:
        logging.basicConfig(
            filename=str(caminho_log),
            level=logging.INFO,
            format="%(asctime)s - [%(levelname)s] - %(message)s",
            encoding="utf-8",
        )
    return logging.getLogger(__name__)
