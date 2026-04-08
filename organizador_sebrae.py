import os
import shutil
import re
import time
from datetime import datetime, timedelta
import pdfplumber
import tkinter as tk
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

# ==========================================
# CONFIGURAÇÕES INICIAIS
# ==========================================
URL_RAE = "https://atendimento.sp.sebrae.com.br/Acesso/Login?ReturnUrl=%2f" # Confirme se é este o link exato de login
cnpjs_com_erro = []

# ==========================================
# 1. FUNÇÕES DE EXTRAÇÃO (PDF E NOME)
# ==========================================
def limpar_documento(texto):
    # Remove tudo que não for número do CNPJ/CPF
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
            
            # Tenta achar um CNPJ no texto todo
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

# ==========================================
# 2. ROBÔ WEB (SELENIUM)
# ==========================================
def registrar_no_rae(driver, dados):
    wait = WebDriverWait(driver, 10)
    
    try:
        # 1. Expandir a aba Pessoa Jurídica
        aba_pj = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[contains(text(), 'Pessoa Jurídica')]")))
        aba_pj.click()
        time.sleep(1)

        # 2. Pesquisa Cliente pelo CNPJ e aperta ENTER
        campo_cnpj = wait.until(EC.visibility_of_element_located((By.ID, "CNPJ")))
        campo_cnpj.clear()
        campo_cnpj.send_keys(dados['cnpj'])
        time.sleep(0.5) 
        campo_cnpj.send_keys(Keys.ENTER) 
        time.sleep(2)
        
        # Verifica erro de cadastro 
        if "Nenhum registro encontrado" in driver.page_source or "desatualizado" in driver.page_source.lower():
            cnpjs_com_erro.append(dados['cnpj'])
            return 
        
        # 3. Clica no Lápis (Alterar)
        try:
            lapis = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "i.fa-pencil")))
            lapis.click()
        except:
            cnpjs_com_erro.append(dados['cnpj']) 
            return

        time.sleep(2)
        
        # Rolar a página para baixo
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(1) 

        # 4. Prosseguir e Confirmar
        btn_prosseguir = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'PROSSEGUIR COM ATENDIMENTO')]")))
        btn_prosseguir.click()
        
        time.sleep(1.5)
        
        btn_sim = wait.until(EC.presence_of_element_located((By.ID, "salvarDados")))
        driver.execute_script("arguments[0].click();", btn_sim)
        time.sleep(2.5)
        
        # === PASSO 5: SELEÇÃO DA PESSOA FÍSICA ===
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

        # ==========================================
        # TELA DE PREENCHIMENTO DO ATENDIMENTO (FINAL)
        # ==========================================
        # Pausa inicial para a tela carregar
        time.sleep(3)
        
        # --- CANAL ---
        # Foca no combobox (que é a estrutura que aceita o clique corretamente)
        canal_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdCanal-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", canal_combobox)
        time.sleep(1)
        
        # Tenta o clique padrão, se a barra bloquear, usa o JS
        try:
            canal_combobox.click()
        except:
            driver.execute_script("arguments[0].click();", canal_combobox)
            
        time.sleep(1) # Espera a barrinha de pesquisa descer
        
        # Usamos visibility_of_element_located para garantir que a barra já desceu
        busca_canal = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        busca_canal.send_keys("Sebrae Aqui") # AJUSTE AQUI SE O CANAL FOR OUTRO
        time.sleep(1)
        busca_canal.send_keys(Keys.ENTER)
        
        # --- LOCAL DE EXECUÇÃO ---
        local_combobox = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@aria-labelledby='select2-CanalRelacionado_IdLocalExecucao-container']")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", local_combobox)
        time.sleep(1)
        
        try:
            local_combobox.click()
        except:
            driver.execute_script("arguments[0].click();", local_combobox)
            
        time.sleep(1)
        
        busca_local = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        busca_local.send_keys("SEBRAE AQUI - SÃO MIGUEL ARCANJO") 
        time.sleep(1)
        busca_local.send_keys(Keys.ENTER)
        
        time.sleep(1)

        # --- BUSCA DO SERVIÇO ---
        campo_busca = wait.until(EC.presence_of_element_located((By.ID, "palavra-pesquisa-input")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_busca)
        time.sleep(0.5)
        campo_busca.clear()
        campo_busca.send_keys(dados['palavra_chave'])
        
        btn_efetuar_busca = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'EFETUAR BUSCA')]")))
        driver.execute_script("arguments[0].click();", btn_efetuar_busca)
        
        xpath_opcao = f"//select[@id='ServicosDisponiveis']//option[contains(text(), '{dados['servico_exato']}')]"
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, xpath_opcao)))
        except:
            raise Exception(f"Demorou demais! O serviço '{dados['servico_exato']}' não carregou na caixa de seleção.")
        
        caixa_servicos = driver.find_element(By.ID, "ServicosDisponiveis")
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", caixa_servicos)
        time.sleep(0.5)
        
        selecao = Select(caixa_servicos)
        selecao.select_by_visible_text(dados['servico_exato'])
        
        btn_adicionar_servico = wait.until(EC.presence_of_element_located((By.XPATH, "//i[contains(@class, 'frente')]/parent::*")))
        driver.execute_script("arguments[0].click();", btn_adicionar_servico)
        time.sleep(1) 
        
        # --- PREENCHIMENTO DOS DADOS FINAIS ---
        campo_necessidade = wait.until(EC.presence_of_element_located((By.XPATH, "//textarea[contains(@id, 'necessidade') or contains(@name, 'necessidade')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", campo_necessidade)
        time.sleep(0.5)
        campo_necessidade.clear()
        campo_necessidade.send_keys(dados['servico_exato'])
        
        btn_plano = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'INCLUIR PLANO')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_plano)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", btn_plano)
        
        select_unidade = wait.until(EC.presence_of_element_located((By.ID, "select2-UnidadeModal-container")))
        driver.execute_script("arguments[0].click();", select_unidade)
        time.sleep(0.5)
        
        busca_unidade = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "input.select2-search__field")))
        busca_unidade.send_keys("Sorocaba")
        time.sleep(1)
        busca_unidade.send_keys(Keys.ENTER)
        
        btn_selecionar_plano = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'SELECIONAR')]")))
        driver.execute_script("arguments[0].click();", btn_selecionar_plano)
        
        time.sleep(1)
        btn_salvar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'SALVAR ATENDIMENTO')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_salvar)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", btn_salvar)
        
        time.sleep(2) 
        btn_finalizar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'FINALIZAR')]")))
        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn_finalizar)
        time.sleep(0.5)
        driver.execute_script("arguments[0].click();", btn_finalizar)
        
        time.sleep(3)
        # Tenta voltar usando vários mapeamentos comuns
        try:
            btn_voltar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(text(), 'VOLTAR') or contains(text(), 'Voltar') or @title='Voltar']")))
            driver.execute_script("arguments[0].click();", btn_voltar)
        except:
            # Se não achar o botão, volta na "marra" pela URL para garantir que o próximo rode
            driver.get(URL_RAE)
        
    except Exception as e:
        print(f"Erro ao processar o CNPJ {dados['cnpj']}. Motivo: {e}")
        cnpjs_com_erro.append(dados['cnpj'])
        driver.get(URL_RAE) # Retorna para a tela inicial em caso de erro grave
        time.sleep(3)

# ==========================================
# 3. LÓGICA PRINCIPAL (ORQUESTRAÇÃO)
# ==========================================
def processar_tudo(pasta_origem, pasta_destino_raiz, data_corte_str):
    try:
        data_corte = datetime.strptime(data_corte_str, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("Erro", "Formato de data inválido. Use DD/MM/AAAA.")
        return

    # Inicia o Navegador
    servico = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=servico)
    driver.maximize_window()
    driver.get(URL_RAE)
    
    # Pausa o código para o humano logar
    messagebox.showinfo("Ação Necessária", "1. O navegador foi aberto.\n2. Faça o login no RAE.\n3. Quando estiver na tela de 'Pesquisa Clientes', clique em OK nesta janela para o robô começar.")
    
    arquivos_movidos = 0
    cnpjs_com_erro.clear()
    
    for nome_arquivo in os.listdir(pasta_origem):
        if not nome_arquivo.lower().endswith('.pdf'):
            continue
            
        caminho_completo = os.path.join(pasta_origem, nome_arquivo)
        
        # Filtro de Data
        data_criacao = os.path.getmtime(caminho_completo)
        data_formatada = datetime.fromtimestamp(data_criacao)
        
        if data_formatada < data_corte:
            continue
            
        servico_nome = ""
        nome_cliente = ""
        cnpj_cliente = ""
        palavra_chave = ""
        servico_exato = ""

        # Identificação e Extração
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
            palavra_chave = "das"
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

        # Execução: Organiza e Registra
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
                
                # Manda o robô registrar
                registrar_no_rae(driver, dados_atendimento)

    driver.quit()
    
    # Relatório Final
    if len(cnpjs_com_erro) > 0:
        lista_erros = "\n".join(set(cnpjs_com_erro))
        messagebox.showwarning("Atenção: Cadastros Pendentes", f"Os seguintes CNPJs tiveram erro, não foram encontrados ou estão desatualizados:\n\n{lista_erros}")
    else:
        messagebox.showinfo("Sucesso!", f"Processo concluído!\n\n{arquivos_movidos} arquivo(s) organizado(s) e registrado(s).")

# ==========================================
# 4. INTERFACE GRÁFICA
# ==========================================
def selecionar_origem():
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_origem.delete(0, tk.END)
        entrada_origem.insert(0, pasta)

def selecionar_destino():
    pasta = filedialog.askdirectory()
    if pasta:
        entrada_destino.delete(0, tk.END)
        entrada_destino.insert(0, pasta)

def iniciar():
    origem = entrada_origem.get()
    destino = entrada_destino.get()
    data_corte = entrada_data.get()

    if not origem or not destino or not data_corte:
        messagebox.showwarning("Atenção", "Preencha todas as pastas e a data de corte!")
        return

    btn_iniciar.config(text="Robô Trabalhando...", state=tk.DISABLED)
    janela.update() 
    
    processar_tudo(origem, destino, data_corte)
    
    btn_iniciar.config(text="🚀 Iniciar Automação", state=tk.NORMAL)

janela = tk.Tk()
janela.title("Super Organizador e Robô RAE - Sebrae")
janela.geometry("550x380")
janela.configure(padx=20, pady=20)

tk.Label(janela, text="Robô de Atendimentos", font=("Arial", 14, "bold")).pack(pady=(0, 20))

frame_origem = tk.Frame(janela)
frame_origem.pack(fill="x", pady=5)
tk.Label(frame_origem, text="Origem (Downloads):").pack(anchor="w")
entrada_origem = tk.Entry(frame_origem, width=50)
entrada_origem.pack(side="left", padx=(0, 10))
tk.Button(frame_origem, text="Procurar", command=selecionar_origem).pack(side="left")

frame_destino = tk.Frame(janela)
frame_destino.pack(fill="x", pady=15)
tk.Label(frame_destino, text="Destino (Sebrae_Organizados):").pack(anchor="w")
entrada_destino = tk.Entry(frame_destino, width=50)
entrada_destino.pack(side="left", padx=(0, 10))
tk.Button(frame_destino, text="Procurar", command=selecionar_destino).pack(side="left")

frame_data = tk.Frame(janela)
frame_data.pack(fill="x", pady=5)
tk.Label(frame_data, text="Processar apenas arquivos a partir de (DD/MM/AAAA):").pack(anchor="w")
entrada_data = tk.Entry(frame_data, width=20)
entrada_data.insert(0, "01/04/2026")
entrada_data.pack(anchor="w")

btn_iniciar = tk.Button(janela, text="🚀 Iniciar Automação", font=("Arial", 12, "bold"), bg="#005b9f", fg="white", command=iniciar, pady=10)
btn_iniciar.pack(fill="x", pady=(20, 0))

janela.mainloop()