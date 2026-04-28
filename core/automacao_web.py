import time
import os
import json
import sys
import traceback
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from utils.logger import configurar_logger
from utils.paths import data_path
from utils.privacy import mascarar_cnpj

logger = configurar_logger()
URL_RAE = "https://atendimento.sp.sebrae.com.br/Acesso/Login?ReturnUrl=%2f" 

# ============================================================
# 📁 CARREGAMENTO DINÂMICO (NOVO CÉREBRO)
# ============================================================
def carregar_base_dados(nome_arquivo):
    try:
        caminho = data_path(nome_arquivo)
        with open(caminho, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        logger.warning(f"Aviso: Base {nome_arquivo} não encontrada: {e}")
        return {}

MAPA_UNIDADES = {
    "ER ALTO TIETÊ": "2", "ER ARAÇATUBA": "3", "ER ARARAQUARA": "4", "ER BAIXADA SANTISTA": "5",
    "ER BARRETOS": "6", "ER BAURU": "7", "ER BOTUCATU": "8", "ER CAMPINAS": "9", "ER CAPITAL CENTRO": "10",
    "ER CAPITAL LESTE": "11", "ER CAPITAL LESTE II": "12", "ER CAPITAL NORTE": "13", "ER CAPITAL OESTE": "14",
    "ER CAPITAL SUL": "15", "ER FRANCA": "16", "ER GRANDE ABC": "17", "ER GUARATINGUETÁ": "18",
    "ER GUARULHOS": "19", "ER JUNDIAÍ": "20", "ER MARÍLIA": "21", "ER OSASCO": "22", "ER OURINHOS": "23",
    "ER PIRACICABA": "24", "ER PRESIDENTE PRUDENTE": "25", "ER RIBEIRÃO PRETO": "26", "ER SÃO CARLOS": "27",
    "ER SÃO JOÃO DA BOA VISTA": "28", "ER SÃO JOSÉ DO RIO PRETO": "29", "ER SÃO JOSÉ DOS CAMPOS": "30",
    "ER SOROCABA": "31", "ER SUDOESTE PAULISTA": "32", "ER VALE DO RIBEIRA": "33", "ER VOTUPORANGA": "34",
    "ACESSO A MERCADOS E SERVIÇOS FINANCEIROS": "36", "ADMINISTRAÇÃO": "37", "ATENDIMENTO AO CLIENTE": "38", 
    "ATENDIMENTO REMOTO": "39", "AUDITORIA": "41", "CULTURA EMPREENDEDORA": "44", 
    "DIRETORIA DE ADMINISTRAÇÃO E FINANÇAS": "45", "DIRETORIA TÉCNICA": "46", "ECONOMIA CRIATIVA E STARTUPS": "6226", 
    "ESCOLA DE NEGÓCIOS SEBRAE-SP": "35", "FINANÇAS E CONTROLADORIA": "47", "GESTÃO DE PESSOAS": "48", 
    "GESTÃO DE SOLUÇÕES E TRANSFORMAÇÃO DIGITAL": "49", "GESTÃO ESTRATÉGICA": "50", "JURÍDICO E SECRETARIA GERAL": "52", 
    "MARKETING PUBLICIDADE E PROPAGANDA INSTITUCIONAL": "51", "OUVIDORIA": "53", 
    "POLÍTICAS PÚBLICAS E RELAÇÕES GOVERNAMENTAIS": "3771", "PROJETOS E OBRAS": "54", 
    "RELACIONAMENTO COM CLIENTE": "42", "TECNOLOGIA CORPORATIVA": "56", "TELEMARKETING ATIVO": "57", 
    "TERRITORIAL E SETORIAL": "40", "Aquisições e Credenciamento": "55"
}

def clicar_js(driver, elemento):
    """Rola a tela até o elemento (para você ver) e força o clique via JS"""
    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", elemento)
    time.sleep(0.3) # Micro-pausa para os seus olhos e para a página renderizar
    driver.execute_script("arguments[0].click();", elemento)

def registrar_no_rae(driver, dados):
    wait = WebDriverWait(driver, 15) 
    
    # 🧠 LENDO OS DADOS DINÂMICOS
    base_projetos = carregar_base_dados("base_dados_projetos.json")
    base_acoes = carregar_base_dados("base_dados_acoes.json")
    
    config = dados.get('config_unidade', {})
    id_unidade = MAPA_UNIDADES.get(config.get('unidade_nome', ''), '')
    projetos_da_unidade = base_projetos.get(f"ER_{id_unidade}", {})
    id_projeto = projetos_da_unidade.get(config.get('projeto', ''), config.get('projeto', ''))
    acoes_do_projeto = base_acoes.get(id_projeto, {})
    id_acao = acoes_do_projeto.get(config.get('acao', ''), config.get('acao', ''))
    
    try:
        logger.info(f"Iniciando registro para o CNPJ: {mascarar_cnpj(dados['cnpj'])}")
        
        # 1. PESQUISA DO CNPJ
        aba_pj = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Pessoa Jurídica')]")))
        clicar_js(driver, aba_pj)
        time.sleep(1.5) # Pausa para a aba renderizar

        campo_cnpj = wait.until(EC.visibility_of_element_located((By.ID, "CNPJ")))
        campo_cnpj.clear()
        campo_cnpj.send_keys(dados['cnpj'])
        time.sleep(0.5) 
        campo_cnpj.send_keys(Keys.ENTER) 
        
        # Aguarda a tabela de resultados carregar
        time.sleep(2) 
        
        if "Nenhum registro encontrado" in driver.page_source or "desatualizado" in driver.page_source.lower():
            logger.warning(f"CNPJ {mascarar_cnpj(dados['cnpj'])} não encontrado ou desatualizado.")
            return "nao_encontrado" 

        # 2. EDIÇÃO (O LÁPIS)
        try:
            lapis = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "i.fa-pencil")))
            clicar_js(driver, lapis)
        except:
            return False

        time.sleep(2) # Aguarda a tela de edição abrir completamente
        
        # 3. PROSSEGUIR
        btn_prosseguir = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'PROSSEGUIR COM ATENDIMENTO')]")))
        clicar_js(driver, btn_prosseguir)
        
        time.sleep(1.5)
        
        btn_sim = wait.until(EC.presence_of_element_located((By.ID, "salvarDados")))
        clicar_js(driver, btn_sim)
        
        # 4. PESSOA FÍSICA
        time.sleep(3) # Transição de modal crítica
        try:
            primeira_celula = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "#dtPf tbody tr td")))
            clicar_js(driver, primeira_celula)
            time.sleep(1.5)
            
            btn_individual = wait.until(EC.presence_of_element_located((By.ID, "btnAtendimentoIndividual")))
            clicar_js(driver, btn_individual)
            
        except Exception as e:
            logger.error(f"Erro ao selecionar PF para CNPJ {mascarar_cnpj(dados['cnpj'])}: {e}")
            driver.refresh()
            return False

        # 5. PREENCHIMENTO DO ATENDIMENTO (FINAL)
        time.sleep(3) # Espera a tela pesada final carregar
        
        # CANAL (SELECT2 - Exige clique físico)
        canal_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdCanal-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", canal_combobox)
        time.sleep(0.5)
        canal_combobox.click() # Clique nativo do Selenium (Sem JS)
        time.sleep(1)
            
        caixas_pesquisa = driver.find_elements(By.CSS_SELECTOR, "input.select2-search__field")
        for caixa in caixas_pesquisa:
            if caixa.is_displayed():
                caixa.send_keys(config.get('canal', 'Sebrae Aqui')) 
                time.sleep(0.5)
                caixa.send_keys(Keys.ENTER)
                break
        
        # LOCAL DE EXECUÇÃO (SELECT2 - Exige clique físico)
        time.sleep(1.5)
        local_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdLocalExecucao-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", local_combobox)
        time.sleep(0.5)
        local_combobox.click() # Clique nativo do Selenium (Sem JS)
        time.sleep(1)
            
        caixas_pesquisa = driver.find_elements(By.CSS_SELECTOR, "input.select2-search__field")
        for caixa in caixas_pesquisa:
            if caixa.is_displayed():
                caixa.send_keys(config.get('local_execucao', '')) 
                time.sleep(0.5)
                caixa.send_keys(Keys.ENTER)
                break

        time.sleep(1)

        # BUSCA DO SERVIÇO 
        campo_busca = wait.until(EC.presence_of_element_located((By.ID, "palavra-pesquisa-input")))
        campo_busca.clear()
        campo_busca.send_keys(dados['palavra_chave'])
        time.sleep(0.5)
        campo_busca.send_keys(Keys.ENTER)
        time.sleep(2)  # Aguarda requisição AJAX do site

        try:
            btn_efetuar_busca = wait.until(
                EC.presence_of_element_located(
                    (By.XPATH, "//button[contains(translate(text(), 'efetuar busca', 'EFETUAR BUSCA'), 'EFETUAR BUSCA')]")
                )
            )
            clicar_js(driver, btn_efetuar_busca) 
            time.sleep(1.5)
        except:
            pass 

        servico_exato = dados["servico_exato"].strip()

        caixa_servicos = wait.until(
            EC.presence_of_element_located((By.ID, "ServicosDisponiveis"))
        )

        selecao = Select(caixa_servicos)

        # Garante que nenhuma opção antiga ficou selecionada
        try:
            selecao.deselect_all()
        except Exception:
            pass

        opcao_encontrada = None

        for opcao in selecao.options:
            texto_opcao = opcao.text.strip()

            if texto_opcao == servico_exato:
                opcao_encontrada = opcao
                break

        if opcao_encontrada is None:
            logger.error(f"Serviço exato não encontrado no RAE: {servico_exato}")
            return False

        driver.execute_script(
            "arguments[0].selected = true;"
            "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
            opcao_encontrada
        )

        logger.info(f"Serviço selecionado uma única vez: {servico_exato}")

        time.sleep(1)

        btn_adicionar_servico = wait.until(
            EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'frente')]/parent::*"))
        )
        clicar_js(driver, btn_adicionar_servico)
        time.sleep(1.5)

        # PREENCHIMENTO DOS DADOS FINAIS E PLANO
        campo_necessidade = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//textarea[contains(@id, 'Necessidade') or contains(@name, 'Necessidade')]")
            )
        )

        driver.execute_script(
            "arguments[0].value = arguments[1];"
            "arguments[0].dispatchEvent(new Event('change', { bubbles: true }));",
            campo_necessidade,
            servico_exato
        )

        time.sleep(1)

        btn_plano = wait.until(
            EC.presence_of_element_located(
                (By.XPATH, "//button[contains(translate(text(), 'incluir plano orçamentário', 'INCLUIR PLANO ORÇAMENTÁRIO'), 'INCLUIR PLANO ORÇAMENTÁRIO')]")
            )
        )
        clicar_js(driver, btn_plano)

        time.sleep(2)
        try:
            # INJEÇÃO COM OS DADOS DINÂMICOS
            driver.execute_script(f"$('#UnidadeModal').val('{id_unidade}').trigger('change');")
            time.sleep(0.5)
            driver.execute_script(f"$('#AnoModal').val('{config.get('ano', '2026')}').trigger('change');")
            time.sleep(0.5)
            driver.execute_script(f"$('#PlanoModal').val('{id_projeto}').trigger('change');")
            time.sleep(0.5)
            driver.execute_script(f"$('#AcaoModal').val('{id_acao}').trigger('change');")
        except Exception as e:
            logger.error(f"Falha na injeção do plano orçamentário: {e}")
            
        time.sleep(1)
        btn_selecionar_plano = wait.until(EC.presence_of_element_located((By.ID, "btnSalvarPlano")))
        clicar_js(driver, btn_selecionar_plano)
        time.sleep(1.5)
        
        # LANÇAMENTO RETROATIVO E SALVAMENTO
        data_hoje = datetime.now().date()
        data_arq = dados['data_arquivo']
        
        if data_arq.date() != data_hoje:
            str_data = data_arq.strftime("%d/%m/%Y")
            str_hora_ini = data_arq.strftime("%H:%M")
            str_hora_fim = (data_arq + timedelta(minutes=2)).strftime("%H:%M")
            
            try:
                aba_retroativo = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'panel-heading') and contains(text(), 'Lançamento Retroativo')]")))
                clicar_js(driver, aba_retroativo)
                time.sleep(1)
            except:
                pass
            
            campo_data = wait.until(EC.presence_of_element_located((By.ID, "LancamentoRetroativoOT_Data")))
            campo_data.clear()
            campo_data.send_keys(str_data)
            
            campo_hora_ini = wait.until(EC.presence_of_element_located((By.ID, "LancamentoRetroativoOT_HorarioInicial")))
            campo_hora_ini.clear()
            campo_hora_ini.send_keys(str_hora_ini)
            
            campo_hora_fim = wait.until(EC.presence_of_element_located((By.ID, "LancamentoRetroativoOT_HorarioFinal")))
            campo_hora_fim.clear()
            campo_hora_fim.send_keys(str_hora_fim)
            
            campo_motivo = wait.until(EC.presence_of_element_located((By.ID, "LancamentoRetroativoOT_Motivo")))
            campo_motivo.clear()
            campo_motivo.send_keys("Problema no sistema")

        time.sleep(1)
        btn_salvar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'salvar atendimento', 'SALVAR ATENDIMENTO'), 'SALVAR ATENDIMENTO')]")))
        clicar_js(driver, btn_salvar)

        time.sleep(3) # Tempo crítico para o servidor salvar
        btn_finalizar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'finalizar', 'FINALIZAR'), 'FINALIZAR')]")))
        clicar_js(driver, btn_finalizar)

        time.sleep(3)
        try:
            btn_voltar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'voltar', 'VOLTAR'), 'VOLTAR') or @title='Voltar']")))
            clicar_js(driver, btn_voltar)
        except:
            driver.get(URL_RAE)
            
        logger.info(f"CNPJ {mascarar_cnpj(dados['cnpj'])} registrado com sucesso.")
        return True
        
    except Exception as e:
        logger.error(f"Erro ao processar o CNPJ {mascarar_cnpj(dados.get('cnpj', ''))}. Motivo: {e}")
        logger.error(traceback.format_exc()) # Grava a linha exata do erro no log
        driver.get(URL_RAE) 
        time.sleep(3)
        return False