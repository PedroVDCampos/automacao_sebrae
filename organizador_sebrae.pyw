import os
import sys
import shutil
import re
import time
import pdfplumber
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
from tkinter import filedialog, messagebox
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.action_chains import ActionChains
from datetime import datetime, timedelta

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


    # CONFIGURAÇÕES INICIAIS

URL_RAE = "https://atendimento.sp.sebrae.com.br/Acesso/Login?ReturnUrl=%2f" 
cnpjs_com_erro = []


    # 1. FUNÇÕES DE EXTRAÇÃO (PDF E NOME)

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
    except:
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
    except:
        pass
    return "Cliente_Parcelamento", ""


    # 2. ROBÔ WEB (SELENIUM)

def registrar_no_rae(driver, dados):
    wait = WebDriverWait(driver, 10)
    
    try:
        aba_pj = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Pessoa Jurídica')]")))
        aba_pj.click()
        time.sleep(1)

        campo_cnpj = wait.until(EC.visibility_of_element_located((By.ID, "CNPJ")))
        campo_cnpj.clear()
        campo_cnpj.send_keys(dados['cnpj'])
        time.sleep(0.5) 
        campo_cnpj.send_keys(Keys.ENTER) 
        time.sleep(2)

        if "Nenhum registro encontrado" in driver.page_source or "desatualizado" in driver.page_source.lower():
            cnpjs_com_erro.append(dados['cnpj'])
            return 

        try:
            lapis = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa-pencil")))
            lapis.click()
        except:
            cnpjs_com_erro.append(dados['cnpj']) 
            return

        time.sleep(2)

        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1) 

        btn_prosseguir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'PROSSEGUIR COM ATENDIMENTO')]")))
        btn_prosseguir.click()
        
        time.sleep(1.5)
        
        btn_sim = wait.until(EC.presence_of_element_located((By.ID, "salvarDados")))
        driver.execute_script("arguments[0].click();", btn_sim)
        time.sleep(2.5)
        
     # PASSO 5: SELEÇÃO DA PESSOA FÍSICA 
        try:
            time.sleep(2) 
            primeira_celula = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "#dtPf tbody tr td")))
            primeira_celula.click()
            time.sleep(5) 
            
            btn_individual = wait.until(EC.presence_of_element_located((By.ID, "btnAtendimentoIndividual")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_individual)
            time.sleep(1) 
            driver.execute_script("window.scrollBy(0, 200);")
            time.sleep(1)
            driver.execute_script("arguments[0].click();", btn_individual)
            
        except Exception as e:
            print(f"Erro ao selecionar pessoa física ou iniciar atendimento: {e}")
            cnpjs_com_erro.append(dados['cnpj'])
            driver.refresh()
            time.sleep(3)
            return

        
        # TELA DE PREENCHIMENTO DO ATENDIMENTO (FINAL)
        
        time.sleep(3)
        
        # CANAL
        canal_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdCanal-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", canal_combobox)
        time.sleep(1)
        
        try:
            canal_combobox.click()
        except:
            driver.execute_script("arguments[0].click();", canal_combobox)
            
        time.sleep(1.5)
        
        caixas_pesquisa = driver.find_elements(By.CSS_SELECTOR, "input.select2-search__field")
        for caixa in caixas_pesquisa:
            if caixa.is_displayed():
                caixa.send_keys("Sebrae Aqui") 
                time.sleep(1)
                caixa.send_keys(Keys.ENTER)
                break
        
        # LOCAL DE EXECUÇÃO
        time.sleep(2)
        
        local_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdLocalExecucao-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", local_combobox)
        time.sleep(0.5)
        
        clicou_local = False
        for tentativa in range(5):
            try:
                local_combobox.click()
                clicou_local = True
                break 
            except:
                time.sleep(1.5)
                
        if not clicou_local:
             driver.execute_script("arguments[0].click();", local_combobox)
             
        time.sleep(1.5)
        
        caixas_pesquisa = driver.find_elements(By.CSS_SELECTOR, "input.select2-search__field")
        for caixa in caixas_pesquisa:
            if caixa.is_displayed():
                caixa.send_keys("SEBRAE AQUI - SÃO MIGUEL ARCANJO") 
                time.sleep(1)
                caixa.send_keys(Keys.ENTER)
                break
        
        time.sleep(1)

        # BUSCA DO SERVIÇO 
        campo_busca = wait.until(EC.presence_of_element_located((By.ID, "palavra-pesquisa-input")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_busca)
        time.sleep(0.5)
        
        campo_busca.clear()
        campo_busca.send_keys(dados['palavra_chave'])
        time.sleep(0.5) 
        campo_busca.send_keys(Keys.ENTER)
        time.sleep(1) 
        
        try:
            btn_efetuar_busca = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'efetuar busca', 'EFETUAR BUSCA'), 'EFETUAR BUSCA')]")))
            try:
                btn_efetuar_busca.click() 
            except:
                driver.execute_script("arguments[0].click();", btn_efetuar_busca) 
        except:
            pass 
        
        xpath_opcao = f"//select[@id='ServicosDisponiveis']//option[contains(text(), '{dados['servico_exato']}')]"
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcao)))
        except:
            raise Exception(f"Demorou demais! O serviço '{dados['servico_exato']}' não carregou.")
        
        caixa_servicos = driver.find_element(By.ID, "ServicosDisponiveis")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", caixa_servicos)
        time.sleep(0.5)
        
        selecao = Select(caixa_servicos)
        selecao.select_by_visible_text(dados['servico_exato'])
        
        btn_adicionar_servico = wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'frente')]/parent::*")))
        driver.execute_script("arguments[0].click();", btn_adicionar_servico)
        time.sleep(2) 
        
        #  PREENCHIMENTO DOS DADOS FINAIS 
        driver.execute_script("window.scrollBy(0, 400);")
        time.sleep(1)
        
        # Necessidade do Cliente
        campo_necessidade = wait.until(EC.presence_of_element_located((By.XPATH, "//textarea[contains(@id, 'Necessidade') or contains(@name, 'Necessidade')]")))
        texto_necessidade = dados['servico_exato']
        driver.execute_script(f"arguments[0].value = '{texto_necessidade}';", campo_necessidade)
        driver.execute_script("arguments[0].dispatchEvent(new Event('change'));", campo_necessidade)
        time.sleep(1)
        
        #  PLANO ORÇAMENTÁRIO 
        driver.execute_script("window.scrollBy(0, 400);")
        time.sleep(1)
        
        btn_plano = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'incluir plano orçamentário', 'INCLUIR PLANO ORÇAMENTÁRIO'), 'INCLUIR PLANO ORÇAMENTÁRIO')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_plano)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", btn_plano)
        
        time.sleep(2)
        try:
            driver.execute_script("$('#UnidadeModal').val('31').trigger('change');")
            time.sleep(2) 
            driver.execute_script("$('#AnoModal').val('2026').trigger('change');")
            time.sleep(1)
            driver.execute_script("$('#PlanoModal').val('82').trigger('change');")
            time.sleep(2) 
            driver.execute_script("$('#AcaoModal').val('260395547').trigger('change');")
            time.sleep(1.5)
        except Exception as e:
            print(f"Atenção: Falha na injeção direta do plano orçamentário. Erro: {e}")
            
        btn_selecionar_plano = wait.until(EC.presence_of_element_located((By.ID, "btnSalvarPlano")))
        driver.execute_script("arguments[0].click();", btn_selecionar_plano)
          
        # LÓGICA FINAL: LANÇAMENTO RETROATIVO E SALVAMENTO
        
        driver.execute_script("window.scrollBy(0, 600);")
        time.sleep(1)
        
        data_hoje = datetime.now().date()
        data_arq = dados['data_arquivo']
        
        if data_arq.date() != data_hoje:
            print("Atendimento retroativo detectado. Preenchendo horários...")
            
            str_data = data_arq.strftime("%d/%m/%Y")
            str_hora_ini = data_arq.strftime("%H:%M")
            str_hora_fim = (data_arq + timedelta(minutes=2)).strftime("%H:%M")
            
            try:
                aba_retroativo = wait.until(EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'panel-heading') and contains(text(), 'Lançamento Retroativo')]")))
                driver.execute_script("arguments[0].click();", aba_retroativo)
                time.sleep(1)
            except:
                pass
            
            campo_data = wait.until(EC.presence_of_element_located((By.ID, "LancamentoRetroativoOT_Data")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_data)
            time.sleep(0.5)
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
            
            time.sleep(1.5) 

        btn_salvar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'salvar atendimento', 'SALVAR ATENDIMENTO'), 'SALVAR ATENDIMENTO')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", btn_salvar)

        time.sleep(3) 
        btn_finalizar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'finalizar', 'FINALIZAR'), 'FINALIZAR')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_finalizar)
        time.sleep(1)
        driver.execute_script("arguments[0].click();", btn_finalizar)

        time.sleep(3)
        try:
            btn_voltar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(translate(text(), 'voltar', 'VOLTAR'), 'VOLTAR') or @title='Voltar']")))
            driver.execute_script("arguments[0].click();", btn_voltar)
        except:
            driver.get(URL_RAE)
        
    except Exception as e:
        print(f"Erro ao processar o CNPJ {dados['cnpj']}. Motivo: {e}")
        cnpjs_com_erro.append(dados['cnpj'])
        driver.get(URL_RAE) 
        time.sleep(3)


# 3. LÓGICA PRINCIPAL (ORQUESTRAÇÃO)

def processar_tudo(pasta_origem, pasta_destino_raiz, data_corte_str):
    try:
        data_corte = datetime.strptime(data_corte_str, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")
        return

    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico)
    driver.maximize_window()
    driver.get(URL_RAE)
    
    messagebox.showinfo("Ação Necessária", "1. O navegador foi aberto.\n2. Faça o login no RAE.\n3. Quando estiver na tela de 'Pesquisa Clientes', clique em OK nesta janela para o robô começar.")
    
    arquivos_movidos = 0
    cnpjs_com_erro.clear()
    
    for nome_arquivo in os.listdir(pasta_origem):
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
            
        else:
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
                
                # Monta os dados para o robô
                dados_atendimento = {
                    'cnpj': cnpj_cliente,
                    'palavra_chave': palavra_chave,
                    'servico_exato': servico_exato,
                    'data_arquivo': data_formatada
                }

                registrar_no_rae(driver, dados_atendimento)

    driver.quit()
    
    # Relatório Final
    if len(cnpjs_com_erro) > 0:
        lista_erros = "\n".join(set(cnpjs_com_erro))
        messagebox.showwarning("Atenção: Cadastros Pendentes", f"Os seguintes CNPJs tiveram erro, não foram encontrados ou estão desatualizados:\n\n{lista_erros}")
    else:
        messagebox.showinfo("Sucesso!", f"Processo concluído!\n\n{arquivos_movidos} arquivo(s) organizado(s) e registrado(s).")


# 4. INTERFACE GRÁFICA (CUSTOM TKINTER)

ctk.set_appearance_mode("Dark") 
ctk.set_default_color_theme("blue") 

def selecionar_origem():
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_origem.delete(0, ctk.END)
        entrada_origem.insert(0, pasta)

def selecionar_destino():
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_destino.delete(0, ctk.END)
        entrada_destino.insert(0, pasta)

def iniciar():
    origem = entrada_origem.get()
    destino = entrada_destino.get()
    data_corte = entrada_data.get()

    if not origem or not destino or not data_corte:
        messagebox.showwarning("Atenção", "Preencha todas as pastas e a data de corte!")
        return

    btn_iniciar.configure(text="Robô Trabalhando...", state="disabled")
    janela.update() 
    
    processar_tudo(origem, destino, data_corte)
    
    btn_iniciar.configure(text="Iniciar Automação", state="normal")

# Criação da Janela Principal
janela = ctk.CTk()
janela.title("Super Organizador e Robô RAE - Sebrae")

janela.iconbitmap(resource_path("icone.ico"))

# Frame principal para dar aquele espaçamento elegante nas bordas
frame_principal = ctk.CTkFrame(janela, fg_color="transparent")
frame_principal.grid(row=0, column=0, padx=30, pady=30, sticky="nsew")
frame_principal.grid_columnconfigure(0, weight=1)

# Título
titulo = ctk.CTkLabel(frame_principal, text="Robô de Atendimentos", font=ctk.CTkFont(size=24, weight="bold"))
titulo.grid(row=0, column=0, pady=(0, 30))

#  Grupo: Origem 
lbl_origem = ctk.CTkLabel(frame_principal, text="Origem (Downloads):", font=ctk.CTkFont(size=14))
lbl_origem.grid(row=1, column=0, sticky="w", pady=(0, 5))

frame_origem = ctk.CTkFrame(frame_principal, fg_color="transparent")
frame_origem.grid(row=2, column=0, sticky="ew", pady=(0, 15))
frame_origem.grid_columnconfigure(0, weight=1)

entrada_origem = ctk.CTkEntry(frame_origem, height=35, placeholder_text="Caminho da pasta de downloads...")
entrada_origem.grid(row=0, column=0, sticky="ew", padx=(0, 10))

btn_origem = ctk.CTkButton(frame_origem, text="Procurar", width=100, height=35, fg_color="gray50", hover_color="gray40", command=selecionar_origem)
btn_origem.grid(row=0, column=1)

#  Grupo: Destino 
lbl_destino = ctk.CTkLabel(frame_principal, text="Destino (Sebrae_Organizados):", font=ctk.CTkFont(size=14))
lbl_destino.grid(row=3, column=0, sticky="w", pady=(0, 5))

frame_destino = ctk.CTkFrame(frame_principal, fg_color="transparent")
frame_destino.grid(row=4, column=0, sticky="ew", pady=(0, 20))
frame_destino.grid_columnconfigure(0, weight=1)

entrada_destino = ctk.CTkEntry(frame_destino, height=35, placeholder_text="Caminho da pasta destino...")
entrada_destino.grid(row=0, column=0, sticky="ew", padx=(0, 10))

btn_destino = ctk.CTkButton(frame_destino, text="Procurar", width=100, height=35, fg_color="gray50", hover_color="gray40", command=selecionar_destino)
btn_destino.grid(row=0, column=1)

#  Grupo: Data 
lbl_data = ctk.CTkLabel(frame_principal, text="Processar apenas arquivos a partir de (DD/MM/AAAA):", font=ctk.CTkFont(size=14))
lbl_data.grid(row=5, column=0, sticky="w", pady=(0, 5))

entrada_data = ctk.CTkEntry(frame_principal, width=150, height=35)
entrada_data.insert(0, "01/04/2026")
entrada_data.grid(row=6, column=0, sticky="w", pady=(0, 30))

#  Botão Iniciar 
btn_iniciar = ctk.CTkButton(frame_principal, text="Iniciar Automação", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=iniciar)
btn_iniciar.grid(row=7, column=0, sticky="ew")

janela.mainloop()