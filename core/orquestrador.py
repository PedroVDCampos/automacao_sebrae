import os
import sys
import shutil
import re
import time
import subprocess
import winreg
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from utils.logger import configurar_logger
from core.extrator_pdf import ler_pdf_padrao, ler_boleto_parcelamento
from core.automacao_web import registrar_no_rae, URL_RAE

logger = configurar_logger()


def _caminho_chromedriver():
    if getattr(sys, 'frozen', False):
        base = sys._MEIPASS
    else:
        base = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
    return os.path.normpath(os.path.join(base, 'drivers', 'chromedriver.exe'))


def _versao_chrome_instalado() -> str | None:
    chaves = [
        r"SOFTWARE\Google\Chrome\BLBeacon",
        r"SOFTWARE\WOW6432Node\Google\Chrome\BLBeacon",
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Google Chrome",
    ]
    for chave in chaves:
        for raiz in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            try:
                with winreg.OpenKey(raiz, chave) as k:
                    versao, _ = winreg.QueryValueEx(k, "version")
                    return versao
            except (FileNotFoundError, OSError):
                continue
    return None


def _versao_chromedriver(caminho: str) -> str | None:
    try:
        resultado = subprocess.run(
            [caminho, "--version"],
            capture_output=True, text=True, timeout=5,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        match = re.search(r"ChromeDriver\s+([\d.]+)", resultado.stdout)
        if match:
            return match.group(1)
    except Exception:
        pass
    return None


def _major(versao: str) -> int:
    try:
        return int(versao.split(".")[0])
    except (ValueError, IndexError):
        return -1


def verificar_compatibilidade_chrome() -> dict:
    caminho_driver = _caminho_chromedriver()
    versao_chrome  = _versao_chrome_instalado()
    versao_driver  = _versao_chromedriver(caminho_driver)

    logger.info(f"Chrome instalado: {versao_chrome} | ChromeDriver embutido: {versao_driver}")

    if versao_chrome is None:
        return {"status": "aviso", "msg": "⚠️  Não foi possível detectar a versão do Google Chrome.\n\nSe o programa não abrir o navegador, verifique se o Chrome está instalado.", "versao_chrome": None, "versao_driver": versao_driver}

    if versao_driver is None:
        return {"status": "erro", "msg": f"❌  ChromeDriver não encontrado.\n\nCaminho esperado:\n{caminho_driver}\n\nContate o desenvolvedor.", "versao_chrome": versao_chrome, "versao_driver": None}

    if _major(versao_chrome) != _major(versao_driver):
        return {"status": "aviso", "msg": (
            f"⚠️  Chrome ({versao_chrome}) e ChromeDriver ({versao_driver}) com versões diferentes.\n\n"
            f"👉  Acesse https://googlechromelabs.github.io/chrome-for-testing/\n"
            f"     Baixe o ChromeDriver para a versão {_major(versao_chrome)} e solicite atualização ao desenvolvedor."
        ), "versao_chrome": versao_chrome, "versao_driver": versao_driver}

    return {"status": "ok", "msg": "", "versao_chrome": versao_chrome, "versao_driver": versao_driver}


def processar_tudo(pasta_origem, pasta_destino_raiz, data_corte_str, evento_cancelar, callback_login, config_unidade):
    try:
        data_corte = datetime.strptime(data_corte_str, "%d/%m/%Y")
    except ValueError:
        return {"status": "erro", "msg": "Formato de data inválido. Use DD/MM/AAAA."}

    caminho_driver = _caminho_chromedriver()
    if not os.path.exists(caminho_driver):
        return {"status": "erro_fatal", "msg": f"chromedriver.exe não encontrado em:\n{caminho_driver}\n\nContate o desenvolvedor."}

    try:
        opcoes = webdriver.ChromeOptions()
        opcoes.add_experimental_option('excludeSwitches', ['enable-logging'])
        servico = Service(executable_path=caminho_driver)
        servico.creation_flags = subprocess.CREATE_NO_WINDOW
        driver = webdriver.Chrome(service=servico, options=opcoes)
        driver.maximize_window()
        driver.get(URL_RAE)
    except Exception as e:
        return {"status": "erro_fatal", "msg": f"O robô não conseguiu abrir o Google Chrome.\n\nMotivo Técnico:\n{e}"}

    callback_login()

    if evento_cancelar.is_set():
        driver.quit()
        return {"status": "cancelado"}

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
        elif nome_arquivo.startswith("DASNSIMEI-"):
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
                    'cnpj':           cnpj_cliente,
                    'palavra_chave':  palavra_chave,
                    'servico_exato':  servico_exato,
                    'data_arquivo':   data_formatada,
                    'config_unidade': config_unidade,
                }

                sucesso = registrar_no_rae(driver, dados_atendimento)
                if not sucesso:
                    cnpjs_com_erro.append(cnpj_cliente)

    driver.quit()

    if evento_cancelar.is_set():
        return {"status": "cancelado"}

    return {"status": "sucesso", "arquivos": arquivos_movidos, "erros": list(set(cnpjs_com_erro))}