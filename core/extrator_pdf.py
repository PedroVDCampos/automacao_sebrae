import re
from datetime import datetime
import pdfplumber
from utils.logger import configurar_logger

logger = configurar_logger()


def limpar_documento(texto):
    return re.sub(r'\D', '', texto or "")


def _extrair_texto_pdf(caminho_pdf):
    """Extrai texto de todas as páginas do PDF, ignorando páginas sem texto extraível."""
    textos = []

    with pdfplumber.open(caminho_pdf) as pdf:
        for pagina in pdf.pages:
            texto_pagina = pagina.extract_text() or ""
            if texto_pagina.strip():
                textos.append(texto_pagina)

    return "\n".join(textos)


def _normalizar_texto(texto):
    """Normaliza espaços para facilitar regex em textos extraídos de PDF."""
    return re.sub(r'\s+', ' ', texto or '').strip()


def _converter_data_br(data_str):
    """Converte uma data no formato DD/MM/AAAA para datetime."""
    try:
        return datetime.strptime(data_str.strip(), "%d/%m/%Y")
    except Exception:
        return None


def extrair_data_abertura_ccmei(caminho_pdf):
    """
    Extrai a 'Data de Abertura' do CCMEI.

    Exemplos esperados no PDF:
    - Data de Abertura
      14/05/2026
    - Data de Abertura 14/05/2026
    """
    try:
        texto = _extrair_texto_pdf(caminho_pdf)
        texto_normalizado = _normalizar_texto(texto)

        padroes = [
            r"Data\s+de\s+Abertura\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
            r"DATA\s+DE\s+ABERTURA\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})",
        ]

        for padrao in padroes:
            match = re.search(padrao, texto_normalizado, flags=re.IGNORECASE)
            if match:
                data = _converter_data_br(match.group(1))
                if data:
                    return data

        logger.warning(f"⚠️ Data de abertura não encontrada no CCMEI: {caminho_pdf}")
        return None

    except Exception as e:
        logger.error(f"Erro ao extrair Data de Abertura do CCMEI {caminho_pdf}: {e}")
        return None


def ler_pdf_padrao(caminho_pdf, identificador_nome):
    try:
        texto = _extrair_texto_pdf(caminho_pdf)
        linhas = texto.split('\n')
        nome_limpo = "Nome_Nao_Encontrado"
        cnpj = ""

        for i, linha in enumerate(linhas):
            if identificador_nome in linha.upper():
                if i + 1 < len(linhas):
                    nome_bruto = linhas[i + 1]
                    nome_limpo = re.sub(r'^[\d.\-/]+\s*', '', nome_bruto).strip()
                    nome_limpo = re.sub(r'\s*[\d.\-/]+$', '', nome_limpo).strip()

        match_cnpj = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', texto)
        if match_cnpj:
            cnpj = limpar_documento(match_cnpj.group())

        return nome_limpo, cnpj

    except Exception as e:
        logger.error(f"Erro ao ler PDF padrão {caminho_pdf}: {e}")
        return "Erro_Leitura", ""


def ler_ccmei(caminho_pdf):
    """
    Lê dados do CCMEI e retorna:
    - nome_cliente
    - cnpj_cliente
    - data_abertura

    A data de abertura é usada pelo orquestrador para diferenciar:
    - Formalização: data de abertura igual à data do arquivo/download;
    - Alteração: data de abertura diferente da data do arquivo/download.
    """
    nome_cliente, cnpj_cliente = ler_pdf_padrao(caminho_pdf, "NOME CIVIL")
    data_abertura = extrair_data_abertura_ccmei(caminho_pdf)
    return nome_cliente, cnpj_cliente, data_abertura


def ler_boleto_parcelamento(caminho_pdf):
    try:
        texto = _extrair_texto_pdf(caminho_pdf)
        linhas = texto.split('\n')
        cnpj = ""

        for i, linha in enumerate(linhas):
            if "CNPJ" in linha.upper():
                # Primeiro tenta na linha seguinte, que é o padrão atual.
                if i + 1 < len(linhas):
                    match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', linhas[i + 1])
                    if match:
                        cnpj = limpar_documento(match.group())
                        return "Cliente_Parcelamento", cnpj

                # Fallback: tenta na própria linha.
                match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', linha)
                if match:
                    cnpj = limpar_documento(match.group())
                    return "Cliente_Parcelamento", cnpj

    except Exception as e:
        logger.error(f"Erro ao ler boleto {caminho_pdf}: {e}")

    return "Cliente_Parcelamento", ""
