import os
import shutil
import re
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from utils.logger import configurar_logger
from core.extrator_pdf import ler_pdf_padrao, ler_boleto_parcelamento
from core.automacao_web import registrar_no_rae, URL_RAE

logger = configurar_logger()

def processar_tudo(pasta_origem, pasta_destino_raiz, data_corte_str, evento_cancelar):
    try:
        data_corte = datetime.strptime(data_corte_str, "%d/%m/%Y")
    except ValueError:
        return {"status": "erro", "msg": "Formato de data inválido. Use DD/MM/AAAA."}
        
    try:
        servico = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=servico)
        driver.maximize_window()
        driver.get(URL_RAE)
    except Exception as e:
        return {"status": "erro_fatal", "msg": f"O robô não conseguiu abrir o Google Chrome.\n\nMotivo Técnico:\n{e}"}
    
    # Pausa intencional para login
    while not evento_cancelar.is_set():
        if "Pesquisa Clientes" in driver.title or "Sebrae" in driver.title: 
            break # Assume que fez login (você pode melhorar essa validação)
            
    arquivos_movidos = 0
    cnpjs_com_erro = []
    logger.info("--- INÍCIO DE NOVA EXECUÇÃO ---")

    for nome_arquivo in os.listdir(pasta_origem):
        if evento_cancelar.is_set():
            logger.info("Operação cancelada pelo usuário.")
            break 
            
        if not nome_arquivo.lower().endswith('.pdf'):
            continue
            
        caminho_completo = os.path.join(pasta_origem, nome_arquivo)
        data_criacao = os.path.getmtime(caminho_completo)
        data_formatada = datetime.fromtimestamp(data_criacao)
        
        if data_formatada < data_corte:
            continue
            
        servico_nome = ""
        nome_cliente = ""
        cnpj_cliente = ""
        palavra_chave = ""
        servico_exato = ""

        if nome_arquivo.startswith("CCMEI-"):
            servico_nome = "Formalizacao"
            palavra_chave = "formalização"
            servico_exato = "MEI - Formalização do MEI"
            nome_cliente, cnpj_cliente = ler_pdf_padrao(caminho_completo, "NOME CIVIL")
        elif nome_arquivo.startswith("CCMEI"):
            servico_nome = "Alteracao"
        elif nome_arquivo.startswith("DASN-"):
            servico_nome = "Declaracao"
            palavra_chave = "dasn"
            servico_exato = "MEI - Declaração Anual do Simples Nacional - DASN - SIMEI"
            nome_cliente, cnpj_cliente = ler_pdf_padrao(caminho_completo, "NOME EMPRESARIAL")
        elif nome_arquivo.startswith("DAS-PGMEI-"):
            servico_nome = "Boleto_DAS"
            palavra_chave = "dasn"
            servico_exato = "MEI - Emissão do DAS"
            nome_cliente = "Cliente_DAS"
            match = re.search(r'DAS-PGMEI-(\d+)-', nome_arquivo)
            if match:
                cnpj_cliente = match.group(1)
        elif nome_arquivo.startswith("ExibirDAS-"):
            servico_nome = "Parcelamento"
            palavra_chave = "parcelamento"
            servico_exato = "MEI - Parcelamento de Débitos"
            nome_cliente, cnpj_cliente = ler_boleto_parcelamento(caminho_completo)
        elif "baixa" in nome_arquivo.lower():
            nome_cliente, cnpj_cliente = ler_pdf_padrao(caminho_completo, "CERTIDÃO DE BAIXA")
            if cnpj_cliente:
                servico_nome = "Baixa"
                palavra_chave = "baixa"
                servico_exato = "Baixa de Inscrição no CNPJ"

        if servico_nome and cnpj_cliente:
            ano = str(data_formatada.year)
            mes = data_formatada.strftime('%m')
            
            nova_pasta = os.path.join(pasta_destino_raiz, ano, mes, servico_nome, nome_cliente)
            os.makedirs(nova_pasta, exist_ok=True)
            
            destino_final = os.path.join(nova_pasta, nome_arquivo)
            if not os.path.exists(destino_final):
                shutil.move(caminho_completo, destino_final)
                arquivos_movidos += 1
                
                dados_atendimento = {
                    'cnpj': cnpj_cliente,
                    'palavra_chave': palavra_chave,
                    'servico_exato': servico_exato,
                    'data_arquivo': data_formatada
                }

                # Executa e checa sucesso
                sucesso = registrar_no_rae(driver, dados_atendimento)
                if not sucesso:
                    cnpjs_com_erro.append(cnpj_cliente)

    driver.quit()
    
    if evento_cancelar.is_set():
        return {"status": "cancelado"}
    
    return {
        "status": "sucesso", 
        "arquivos": arquivos_movidos, 
        "erros": list(set(cnpjs_com_erro))
    }