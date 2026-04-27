import re

def apenas_digitos(valor: str) -> str:
    return re.sub(r"\D", "", str(valor or ""))

def mascarar_cnpj(cnpj: str) -> str:
    digitos = apenas_digitos(cnpj)
    if len(digitos) != 14:
        return "***"
    return f"{digitos[:2]}.***.***/****-{digitos[-2:]}"

def nome_seguro_para_pasta(nome: str) -> str:
    nome = str(nome or "Cliente").strip()
    nome = re.sub(r'[\\/:*?"<>|]', "_", nome)
    nome = re.sub(r"\s+", " ", nome)
    return nome[:120] or "Cliente"
