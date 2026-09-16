import re
from datetime import datetime
import pdfplumber
from utils.logger import configurar_logger

logger = configurar_logger()

def limpar_documento(texto):
    return re.sub(r"\D", "", str(texto or ""))

def _texto_primeira_pagina(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        return (pdf.pages[0].extract_text() or "") if pdf.pages else ""

def ler_pdf_padrao(caminho_pdf, identificador_nome):
    try:
        texto = _texto_primeira_pagina(caminho_pdf)
        linhas = texto.splitlines()
        nome = "Nome_Nao_Encontrado"
        cnpj = ""
        for i, linha in enumerate(linhas):
            if identificador_nome.upper() in linha.upper() and i + 1 < len(linhas):
                nome = re.sub(r"^[\d.\-/]+\s*", "", linhas[i + 1]).strip()
                nome = re.sub(r"\s*[\d.\-/]+$", "", nome).strip()
                break
        match = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", texto)
        if match:
            cnpj = limpar_documento(match.group())
        return nome, cnpj
    except Exception as e:
        logger.error("Erro ao ler PDF padrão: %s", e)
        return "Erro_Leitura", ""

def ler_boleto_parcelamento(caminho_pdf):
    try:
        texto = _texto_primeira_pagina(caminho_pdf)
        linhas = texto.splitlines()
        for i, linha in enumerate(linhas):
            if "CNPJ" in linha.upper():
                bloco = " ".join(linhas[i:i + 4])
                match = re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", bloco)
                if match:
                    return "Cliente_Parcelamento", limpar_documento(match.group())
    except Exception as e:
        logger.error("Erro ao ler boleto de parcelamento: %s", e)
    return "Cliente_Parcelamento", ""

def extrair_data_abertura_ccmei(caminho_pdf):
    try:
        texto = _texto_primeira_pagina(caminho_pdf)
        normalizado = re.sub(r"[ \t]+", " ", texto)
        padroes = [
            r"DATA\s+DE\s+ABERTURA\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s+DE\s+ABERTURA.*?(\d{2}/\d{2}/\d{4})",
        ]
        for padrao in padroes:
            match = re.search(padrao, normalizado, re.I | re.S)
            if match:
                return datetime.strptime(match.group(1), "%d/%m/%Y")
        linhas = texto.splitlines()
        for i, linha in enumerate(linhas):
            if "DATA DE ABERTURA" in linha.upper():
                match = re.search(r"\d{2}/\d{2}/\d{4}", " ".join(linhas[i:i + 4]))
                if match:
                    return datetime.strptime(match.group(), "%d/%m/%Y")
    except Exception as e:
        logger.error("Erro ao extrair Data de Abertura do CCMEI: %s", e)
    return None

def ler_ccmei(caminho_pdf, data_arquivo):
    nome, cnpj = ler_pdf_padrao(caminho_pdf, "NOME CIVIL")
    abertura = extrair_data_abertura_ccmei(caminho_pdf)
    if not abertura:
        return {"tipo": None, "nome_cliente": nome, "cnpj": cnpj,
                "erro": "Data de Abertura não encontrada no CCMEI"}
    tipo = "Formalizacao" if abertura.date() == data_arquivo.date() else "Alteracao"
    return {"tipo": tipo, "nome_cliente": nome, "cnpj": cnpj, "erro": None}

def identificar_documento(caminho_pdf, nome_arquivo, data_arquivo):
    n = nome_arquivo.upper()
    if n.startswith("CCMEI"):
        return ler_ccmei(caminho_pdf, data_arquivo)
    if n.startswith("DASNSIMEI-"):
        nome, cnpj = ler_pdf_padrao(caminho_pdf, "NOME EMPRESARIAL")
        return {"tipo": "Declaracao", "nome_cliente": nome, "cnpj": cnpj,
                "erro": None if cnpj else "CNPJ não encontrado"}
    if n.startswith("DAS-PGMEI-"):
        match = re.search(r"DAS-PGMEI-(\d+)-", nome_arquivo, re.I)
        cnpj = limpar_documento(match.group(1)) if match else ""
        return {"tipo": "Boleto_DAS", "nome_cliente": "Cliente_DAS", "cnpj": cnpj,
                "erro": None if cnpj else "CNPJ não encontrado no nome do arquivo"}
    if n.startswith("EXIBIRDAS-"):
        nome, cnpj = ler_boleto_parcelamento(caminho_pdf)
        return {"tipo": "Parcelamento", "nome_cliente": nome, "cnpj": cnpj,
                "erro": None if cnpj else "CNPJ não encontrado"}
    if "BAIXA" in n:
        nome, cnpj = ler_pdf_padrao(caminho_pdf, "CERTIDÃO DE BAIXA")
        return {"tipo": "Baixa", "nome_cliente": nome, "cnpj": cnpj,
                "erro": None if cnpj else "CNPJ não encontrado"}
    return {"tipo": None, "nome_cliente": "", "cnpj": "",
            "erro": "Tipo de documento não identificado"}
