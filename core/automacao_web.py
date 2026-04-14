import time
from datetime import datetime, timedelta
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support import expected_conditions as EC
from utils.logger import configurar_logger

logger = configurar_logger()
URL_RAE = "https://atendimento.sp.sebrae.com.br/Acesso/Login?ReturnUrl=%2f" 

def registrar_no_rae(driver, dados):
    wait = WebDriverWait(driver, 15) 
    
    try:
        logger.info(f"Iniciando registro para o CNPJ: {dados['cnpj']}")
        
        aba_pj = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Pessoa Jurídica')]")))
        aba_pj.click()
        time.sleep(1)

        campo_cnpj = wait.until(EC.visibility_of_element_located((By.ID, "CNPJ")))
        campo_cnpj.clear()
        campo_cnpj.send_keys(dados['cnpj'])
        time.sleep(0.5) 
        campo_cnpj.send_keys(Keys.ENTER) 
        
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table, .alert, div.resultados"))) 

        if "Nenhum registro encontrado" in driver.page_source or "desatualizado" in driver.page_source.lower():
            logger.warning(f"CNPJ {dados['cnpj']} não encontrado ou desatualizado.")
            return False 

        try:
            lapis = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa-pencil")))
            lapis.click()
        except:
            return False

        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        
        btn_prosseguir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'PROSSEGUIR COM ATENDIMENTO')]")))
        btn_prosseguir.click()
        
        btn_sim = wait.until(EC.presence_of_element_located((By.ID, "salvarDados")))
        driver.execute_script("arguments[0].click();", btn_sim)
        
        try:
            primeira_celula = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dtPf tbody tr td")))
            primeira_celula.click()
            time.sleep(1)
            
            btn_individual = wait.until(EC.presence_of_element_located((By.ID, "btnAtendimentoIndividual")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_individual)
            driver.execute_script("arguments[0].click();", btn_individual)
            
        except Exception as e:
            logger.error(f"Erro ao selecionar PF para CNPJ {dados['cnpj']}: {e}")
            driver.refresh()
            return False

        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
        
        canal_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdCanal-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", canal_combobox)
        driver.execute_script("arguments[0].click();", canal_combobox)
            
        caixas_pesquisa = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input.select2-search__field")))
        for caixa in caixas_pesquisa:
            if caixa.is_displayed():
                caixa.send_keys("Sebrae Aqui") 
                time.sleep(0.5)
                caixa.send_keys(Keys.ENTER)
                break
        
        local_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdLocalExecucao-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", local_combobox)
        driver.execute_script("arguments[0].click();", local_combobox)
            
        caixas_pesquisa = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "input.select2-search__field")))
        for caixa in caixas_pesquisa:
            if caixa.is_displayed():
                caixa.send_keys("SEBRAE AQUI - SÃO MIGUEL ARCANJO") 
                time.sleep(0.5)
                caixa.send_keys(Keys.ENTER)
                break

        campo_busca = wait.until(EC.presence_of_element_located((By.ID, "palavra-pesquisa-input")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_busca)
        
        campo_busca.clear()
        campo_busca.send_keys(dados['palavra_chave'])
        campo_busca.send_keys(Keys.ENTER)
        time.sleep(1) 
        
        try:
            btn_efetuar_busca = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'efetuar busca', 'EFETUAR BUSCA'), 'EFETUAR BUSCA')]")))
            driver.execute_script("arguments[0].click();", btn_efetuar_busca) 
        except:
            pass 
        
        xpath_opcao = f"//select[@id='ServicosDisponiveis']//option[contains(text(), '{dados['servico_exato']}')]"
        wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcao)))
        
        caixa_servicos = driver.find_element(By.ID, "ServicosDisponiveis")
        selecao = Select(caixa_servicos)
        selecao.select_by_visible_text(dados['servico_exato'])
        
        btn_adicionar_servico = wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'frente')]/parent::*")))
        driver.execute_script("arguments[0].click();", btn_adicionar_servico)
        
        campo_necessidade = wait.until(EC.presence_of_element_located((By.XPATH, "//textarea[contains(@id, 'Necessidade') or contains(@name, 'Necessidade')]")))
        driver.execute_script(f"arguments[0].value = '{dados['servico_exato']}';", campo_necessidade)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", campo_necessidade)
        
        btn_plano = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'incluir plano orçamentário', 'INCLUIR PLANO ORÇAMENTÁRIO'), 'INCLUIR PLANO ORÇAMENTÁRIO')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_plano)
        driver.execute_script("arguments[0].click();", btn_plano)
        
        time.sleep(1.5) 
        try:
            driver.execute_script("$('#UnidadeModal').val('31').trigger('change');")
            driver.execute_script("$('#AnoModal').val('2026').trigger('change');")
            driver.execute_script("$('#PlanoModal').val('82').trigger('change');")
            driver.execute_script("$('#AcaoModal').val('260395547').trigger('change');")
        except Exception as e:
            logger.error(f"Falha na injeção do plano orçamentário: {e}")
            
        btn_selecionar_plano = wait.until(EC.presence_of_element_located((By.ID, "btnSalvarPlano")))
        driver.execute_script("arguments[0].click();", btn_selecionar_plano)
        
        data_hoje = datetime.now().date()
        data_arq = dados['data_arquivo']
        
        if data_arq.date() != data_hoje:
            logger.info("Atendimento retroativo detectado.")
            str_data = data_arq.strftime("%d/%m/%Y")
            str_hora_ini = data_arq.strftime("%H:%M")
            str_hora_fim = (data_arq + timedelta(minutes=2)).strftime("%H:%M")
            
            try:
                aba_retroativo = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'panel-heading') and contains(text(), 'Lançamento Retroativo')]")))
                driver.execute_script("arguments[0].click();", aba_retroativo)
            except:
                pass
            
            campo_data = wait.until(EC.presence_of_element_located((By.ID, "LancamentoRetroativoOT_Data")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_data)
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

        btn_salvar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'salvar atendimento', 'SALVAR ATENDIMENTO'), 'SALVAR ATENDIMENTO')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
        driver.execute_script("arguments[0].click();", btn_salvar)

        btn_finalizar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'finalizar', 'FINALIZAR'), 'FINALIZAR')]")))
        driver.execute_script("arguments[0].click();", btn_finalizar)

        try:
            btn_voltar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'voltar', 'VOLTAR'), 'VOLTAR') or @title='Voltar']")))
            driver.execute_script("arguments[0].click();", btn_voltar)
        except:
            driver.get(URL_RAE)
            
        logger.info(f"CNPJ {dados['cnpj']} registrado com sucesso.")
        return True
        
    except Exception as e:
        logger.error(f"Erro ao processar o CNPJ {dados['cnpj']}. Motivo: {e}")
        driver.get(URL_RAE) 
        return False