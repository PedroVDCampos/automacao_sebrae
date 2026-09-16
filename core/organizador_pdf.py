import os
import shutil
from pathlib import Path
from utils.logger import configurar_logger
from utils.privacy import nome_seguro_para_pasta

logger = configurar_logger()

def criar_pasta_erro(pasta_origem, categoria):
    pasta = os.path.join(pasta_origem, "_ERROS_ORGANIZACAO", categoria)
    os.makedirs(pasta, exist_ok=True)
    return pasta

def mover_para_erro(caminho_pdf, pasta_origem, categoria):
    pasta = criar_pasta_erro(pasta_origem, categoria)
    nome = os.path.basename(caminho_pdf)
    destino = os.path.join(pasta, nome)
    if os.path.abspath(caminho_pdf) == os.path.abspath(destino):
        return destino
    if os.path.exists(destino):
        p = Path(nome)
        i = 2
        while os.path.exists(destino):
            destino = os.path.join(pasta, f"{p.stem} ({i}){p.suffix}")
            i += 1
    shutil.move(caminho_pdf, destino)
    return destino

def organizar_pdf(caminho_pdf, pasta_destino_raiz, data_arquivo, tipo_atendimento, nome_cliente):
    try:
        pasta = os.path.join(pasta_destino_raiz, str(data_arquivo.year),
                             data_arquivo.strftime("%m"), tipo_atendimento,
                             nome_seguro_para_pasta(nome_cliente))
        os.makedirs(pasta, exist_ok=True)
        destino = os.path.join(pasta, os.path.basename(caminho_pdf))
        if os.path.exists(destino):
            return False, destino, "Arquivo já existente no destino"
        shutil.move(caminho_pdf, destino)
        return True, destino, None
    except Exception as e:
        logger.exception("Erro ao organizar PDF: %s", caminho_pdf)
        return False, "", str(e)
