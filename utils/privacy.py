import re

def mascarar_cnpj(cnpj):
    n = re.sub(r"\D", "", str(cnpj or ""))
    return ("*" * max(0, len(n)-4) + n[-4:]) if n else ""

def nome_seguro_para_pasta(nome):
    nome = str(nome or "").strip()
    nome = re.sub(r'[<>:"/\\|?*]', "_", nome)
    nome = re.sub(r"\s+", " ", nome).strip().rstrip(". ")
    return (nome or "Cliente_Nao_Identificado")[:150]
