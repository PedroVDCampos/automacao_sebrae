import logging
import os

def configurar_logger():
    # Evita criar múltiplos manipuladores se a função for chamada duas vezes
    if not logging.getLogger().hasHandlers():
        logging.basicConfig(
            filename='rae_turbo_execucao.log',
            level=logging.INFO,
            format='%(asctime)s - [%(levelname)s] - %(message)s'
        )
    return logging.getLogger(__name__)