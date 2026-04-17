import logging
import os

def configurar_logger():
    # Caminho Sênior: Salva o log na pasta AppData/Local do Windows do usuário
    pasta_appdata = os.path.join(os.path.expanduser("~"), "AppData", "Local", "RAETurbo")
    
    # Cria a pasta invisível se ela não existir
    os.makedirs(pasta_appdata, exist_ok=True)
    
    # Define o arquivo de log lá dentro
    caminho_log = os.path.join(pasta_appdata, "rae_turbo_execucao.log")
    
    if not logging.getLogger().hasHandlers():
        logging.basicConfig(
            filename=caminho_log,
            level=logging.INFO,
            format='%(asctime)s - [%(levelname)s] - %(message)s'
        )
    return logging.getLogger(__name__)