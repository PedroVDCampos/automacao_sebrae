import os
import sys
import shutil
import re
import time
import subprocess
import winreg
import winsound
import ctypes
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from utils.logger import configurar_logger
from utils.paths import resource_path
from utils.privacy import mascarar_cnpj, nome_seguro_para_pasta
from core.extrator_pdf import ler_pdf_padrao, ler_boleto_parcelamento
from core.automacao_web import registrar_no_rae, URL_RAE

logger = configurar_logger()

def _caminho_chromedriver():
    return os.path.normpath(resource_path(os.path.join('drivers', 'chromedriver.exe')))

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
        return {"status": "aviso", "msg": "⚠️ Não foi possível detectar a versão do Google Chrome.\n\nSe o programa não abrir o navegador, verifique se o Chrome está instalado.", "versao_chrome": None, "versao_driver": versao_driver}

    if versao_driver is None:
        return {"status": "erro", "msg": f"❌ ChromeDriver não encontrado.\n\nCaminho esperado:\n{caminho_driver}\n\nContate o desenvolvedor.", "versao_chrome": versao_chrome, "versao_driver": None}

    if _major(versao_chrome) != _major(versao_driver):
        return {"status": "aviso", "msg": (
            f"⚠️ Chrome ({versao_chrome}) e ChromeDriver ({versao_driver}) com versões diferentes.\n\n"
            f"👉 Baixe o ChromeDriver para a versão {_major(versao_chrome)} e solicite atualização ao desenvolvedor."
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
        return {"status": "erro_fatal", "msg": f"O robô não conseguiu abrir o Chrome.\n\nMotivo Técnico:\n{e}"}

    login_autorizado = callback_login()

    if not login_autorizado:
        try:
            driver.quit()
        except Exception:
            pass
        return {"status": "cancelado"}

    if evento_cancelar.is_set():
        driver.quit()
        return {"status": "cancelado"}

    arquivos_movidos = 0
    cnpjs_com_erro = []
    logger.info("--- INÍCIO DE NOVA EXECUÇÃO ---")
    
    # 🧠 MEMÓRIA DE CURTO PRAZO (Evita o problema da Onipresença)
    memoria_atendimentos = set()

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

            nova_pasta = os.path.join(pasta_destino_raiz, ano, mes, servico_nome, nome_seguro_para_pasta(nome_cliente))
            os.makedirs(nova_pasta, exist_ok=True)

            destino_final = os.path.join(nova_pasta, nome_arquivo)
            if not os.path.exists(destino_final):
                shutil.move(caminho_completo, destino_final)
                arquivos_movidos += 1

            # 🛡️ FILTRO DE ONIPRESENÇA: Verifica se já atendeu esse CNPJ hoje para o mesmo serviço
            assinatura_atendimento = f"{cnpj_cliente}_{data_formatada.strftime('%Y-%m-%d')}_{servico_exato}"
            if assinatura_atendimento in memoria_atendimentos:
                logger.info(f"⏭️ Duplicidade barrada: CNPJ {mascarar_cnpj(cnpj_cliente)} já recebeu o serviço '{servico_exato}' hoje. Arquivo apenas organizado.")
                continue

            dados_atendimento = {
                'cnpj':           cnpj_cliente,
                'palavra_chave':  palavra_chave,
                'servico_exato':  servico_exato,
                'data_arquivo':   data_formatada,
                'config_unidade': config_unidade,
            }

            # 🚀 LANÇAMENTO NO RAE
            sucesso = registrar_no_rae(driver, dados_atendimento)
            
            if sucesso == True:
                memoria_atendimentos.add(assinatura_atendimento)
                
            elif sucesso == "nao_encontrado":
                # CNPJ não existe no sistema do Sebrae. Não precisa reiniciar o navegador.
                cnpjs_com_erro.append(f"{mascarar_cnpj(cnpj_cliente)} (Não Encontrado)")
                
            else:
                # 🛑 FALHA CRÔNICA: Se retornou False, o botão prosseguir ou o site travou.
                logger.warning(f"⚠️ Travamento detectado no CNPJ {mascarar_cnpj(cnpj_cliente)}. Iniciando Protocolo de Segurança (Restart)...")
                
                # Dispara Alarme Sonoro de Alerta
                for _ in range(4):
                    winsound.Beep(1500, 400)
                    time.sleep(0.1)
                
                try: driver.quit()
                except: pass

                try:
                    # Recria o navegador do zero para limpar os erros do Sebrae
                    servico_novo = Service(executable_path=caminho_driver)
                    servico_novo.creation_flags = subprocess.CREATE_NO_WINDOW
                    driver = webdriver.Chrome(service=servico_novo, options=opcoes)
                    driver.maximize_window()
                    driver.get(URL_RAE)
                except Exception as e:
                    return {"status": "erro_fatal", "msg": f"Erro crítico ao tentar reiniciar o navegador: {e}"}

                # Pausa Nativa com alerta por cima de todas as janelas
                ctypes.windll.user32.MessageBoxW(
                    0,
                    f"O robô travou no CNPJ: {mascarar_cnpj(cnpj_cliente)}\nO navegador foi reiniciado para limpar o cache do Sebrae.\n\n"
                    "1. Faça o LOGIN novamente.\n"
                    "2. Vá até a tela de 'Pesquisa Clientes'.\n"
                    "3. Clique em OK para o robô tentar o atendimento mais uma vez.",
                    "⚠️ Reinício de Segurança (RAE Turbo)",
                    0x30 | 0x40000
                )
                
                # Segunda Tentativa!
                logger.info(f"🔄 Tentativa de resgate do CNPJ {mascarar_cnpj(cnpj_cliente)}...")
                sucesso_retry = registrar_no_rae(driver, dados_atendimento)
                
                if sucesso_retry == True:
                    logger.info("✅ Resgate bem-sucedido após reinício!")
                    memoria_atendimentos.add(assinatura_atendimento)
                else:
                    logger.error(f"❌ O CNPJ {mascarar_cnpj(cnpj_cliente)} falhou definitivamente.")
                    cnpjs_com_erro.append(mascarar_cnpj(cnpj_cliente))

    try: driver.quit()
    except: pass

    if evento_cancelar.is_set():
        return {"status": "cancelado"}

    return {"status": "sucesso", "arquivos": arquivos_movidos, "erros": list(set(cnpjs_com_erro))}