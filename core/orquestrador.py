import os
import threading
from datetime import datetime
from core.extrator_pdf import identificar_documento
from core.organizador_pdf import organizar_pdf, mover_para_erro
from utils.logger import configurar_logger
from utils.privacy import mascarar_cnpj
from utils.relatorio_execucao import novo_resumo_execucao, finalizar_resumo_execucao
from version import VERSAO_ATUAL

logger = configurar_logger()

def _listar_pdfs(pasta):
    return sorted([n for n in os.listdir(pasta)
                   if n.lower().endswith(".pdf") and os.path.isfile(os.path.join(pasta, n))], key=str.lower)

def _categoria_erro(motivo):
    t = (motivo or "").lower()
    if "data de abertura" in t: return "CCMEI_Data_Abertura_Nao_Encontrada"
    if "cnpj" in t: return "CNPJ_Nao_Identificado"
    if "tipo" in t: return "Tipo_Nao_Identificado"
    if "nome" in t: return "Nome_Nao_Identificado"
    return "Outros"

def _erro(resumo, caminho, origem, motivo):
    resumo["arquivos_com_erro"] += 1
    resumo["erros"].append(f"{os.path.basename(caminho)}: {motivo}")
    try:
        destino = mover_para_erro(caminho, origem, _categoria_erro(motivo))
        logger.warning("PDF enviado para erro: %s", destino)
    except Exception as e:
        logger.error("Falha ao mover PDF para erro: %s", e)

def _contar_tipo(resumo, tipo):
    campos = {"Formalizacao":"formalizacoes", "Alteracao":"alteracoes", "Declaracao":"declaracoes",
              "Boleto_DAS":"boletos_das", "Parcelamento":"parcelamentos", "Baixa":"baixas"}
    if tipo in campos: resumo[campos[tipo]] += 1

def processar_tudo(pasta_origem, pasta_destino_raiz, data_corte_str, evento_cancelar=None,
                   callback_progresso=None, config_unidade=None):
    evento_cancelar = evento_cancelar or threading.Event()
    try:
        data_corte = datetime.strptime(data_corte_str, "%d/%m/%Y")
    except ValueError:
        return {"status":"erro", "msg":"Formato de data inválido. Use DD/MM/AAAA."}
    if not os.path.isdir(pasta_origem):
        return {"status":"erro", "msg":"A pasta de origem não existe."}
    try:
        os.makedirs(pasta_destino_raiz, exist_ok=True)
    except OSError as e:
        return {"status":"erro", "msg":f"Não foi possível criar a pasta de destino:\n{e}"}

    resumo = novo_resumo_execucao(pasta_origem, pasta_destino_raiz, data_corte_str, VERSAO_ATUAL)
    arquivos = _listar_pdfs(pasta_origem)
    resumo["pdfs_encontrados"] = len(arquivos)
    logger.info("=== INÍCIO DA ORGANIZAÇÃO DE PDFs ===")

    for indice, nome_arquivo in enumerate(arquivos, 1):
        if evento_cancelar.is_set():
            resumo_final = finalizar_resumo_execucao(resumo, "cancelado")
            return {"status":"cancelado", "resumo":resumo_final}
        caminho = os.path.join(pasta_origem, nome_arquivo)
        try:
            data_arquivo = datetime.fromtimestamp(os.path.getmtime(caminho))
        except OSError as e:
            _erro(resumo, caminho, pasta_origem, f"Não foi possível obter a data do arquivo: {e}")
            continue
        if data_arquivo < data_corte:
            resumo["pdfs_fora_da_data"] += 1
            if callback_progresso: callback_progresso(indice, len(arquivos), nome_arquivo, resumo)
            continue
        try:
            dados = identificar_documento(caminho, nome_arquivo, data_arquivo)
        except Exception as e:
            _erro(resumo, caminho, pasta_origem, f"Erro na leitura do documento: {e}")
            continue
        tipo, nome, cnpj, motivo = dados.get("tipo"), dados.get("nome_cliente"), dados.get("cnpj"), dados.get("erro")
        if motivo or not tipo:
            resumo["pdfs_sem_dados"] += 1
            _erro(resumo, caminho, pasta_origem, motivo or "Dados insuficientes")
            continue
        if not cnpj:
            resumo["pdfs_sem_dados"] += 1
            _erro(resumo, caminho, pasta_origem, "CNPJ não identificado")
            continue
        if not nome or nome in {"Nome_Nao_Encontrado", "Erro_Leitura"}:
            resumo["pdfs_sem_dados"] += 1
            _erro(resumo, caminho, pasta_origem, "Nome do cliente não identificado")
            continue

        resumo["pdfs_processados"] += 1
        _contar_tipo(resumo, tipo)
        sucesso, destino, erro = organizar_pdf(caminho, pasta_destino_raiz, data_arquivo, tipo, nome)
        if sucesso:
            resumo["arquivos_organizados"] += 1
            logger.info("Organizado %s -> %s | CNPJ %s", nome_arquivo, destino, mascarar_cnpj(cnpj))
        elif erro == "Arquivo já existente no destino":
            resumo["arquivos_ja_existentes"] += 1
        else:
            _erro(resumo, caminho, pasta_origem, erro or "Falha ao mover arquivo")
        if callback_progresso: callback_progresso(indice, len(arquivos), nome_arquivo, resumo)

    status = "cancelado" if evento_cancelar.is_set() else "sucesso"
    resumo_final = finalizar_resumo_execucao(resumo, status)
    logger.info("=== FIM DA ORGANIZAÇÃO: %s ===", status)
    return {"status":status, "arquivos":resumo_final["arquivos_organizados"],
            "erros":resumo_final["erros"], "resumo":resumo_final}
