import re
import pdfplumber
from utils.logger import configurar_logger

logger = configurar_logger()

def limpar_documento(texto):
    return re.sub(r'\D', '', texto)

def ler_pdf_padrao(caminho_pdf, identificador_nome):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            linhas = texto.split('\n')
            nome_limpo = "Nome_Nao_Encontrado"
            cnpj = ""
            
            for i, linha in enumerate(linhas):
                if identificador_nome in linha.upper():
                    nome_bruto = linhas[i+1]
                    nome_limpo = re.sub(r'^[\d.\-/]+\s*', '', nome_bruto).strip()
                    nome_limpo = re.sub(r'\s*[\d.\-/]+$', '', nome_limpo).strip()
            
            match_cnpj = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', texto)
            if match_cnpj:
                cnpj = limpar_documento(match_cnpj.group())
                
            return nome_limpo, cnpj
    except Exception as e:
        logger.error(f"Erro ao ler PDF padrão {caminho_pdf}: {e}")
        return "Erro_Leitura", ""

def ler_boleto_parcelamento(caminho_pdf):
    try:
        with pdfplumber.open(caminho_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            linhas = texto.split('\n')
            cnpj = ""
            
            for i, linha in enumerate(linhas):
                if "CNPJ" in linha.upper():
                    match = re.search(r'\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}', linhas[i+1])
                    if match:
                        cnpj = limpar_documento(match.group())
                        return "Cliente_Parcelamento", cnpj
    except Exception as e:
        logger.error(f"Erro ao ler boleto {caminho_pdf}: {e}")
    return "Cliente_Parcelamento", ""