import time
import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# A lista completa
lista_ids_er = [
    "36", "37", "2", "55", "3", "4", "38", "39", "41", "5", 
    "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", 
    "24292", "44", "45", "46", "6226", "35", "24707", "47", 
    "16", "48", "49", "50", "17", "18", "19", "20", "52", 
    "21", "51", "22", "23", "53", "24", "3771", "25", "54", 
    "42", "26", "27", "28", "29", "30", "31", "32", "56", 
    "57", "40", "33", "34"
]

def injetar_id(driver, id_select, valor_id, tempo_espera):
    # Versão mais segura que verifica se a caixa existe e está pronta
    script = f"""
    var el = $('#{id_select}');
    if(el.length) {{
        el.val('{valor_id}').trigger('change');
        if (el.hasClass('select2-hidden-accessible')) {{
            el.trigger('change.select2');
        }}
    }}
    """
    driver.execute_script(script)
    time.sleep(tempo_espera)

def rodar_mapeador():
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    wait = WebDriverWait(driver, 15)
    
    driver.get("https://atendimento.sp.sebrae.com.br/Acesso/Login?ReturnUrl=%2f")
    print("FAÇA O LOGIN. O robô vai aguardar você chegar na tela de atendimento...")
    
    while True:
        try:
            if driver.find_elements(By.ID, "UnidadeModal"):
                break
        except:
            pass
        time.sleep(1)
        
    print("Iniciando Mapeamento em Massa. Vá tomar um café!")
    
    base_de_dados = {}

    for id_er in lista_ids_er:
        print(f"Mapeando Unidade ID: {id_er}...")
        try:
            # Se a página estiver fora do lugar, recarrega
            if not driver.find_elements(By.ID, "UnidadeModal"):
                driver.refresh()
                time.sleep(4)
            
            injetar_id(driver, "UnidadeModal", id_er, 2.5)
            injetar_id(driver, "AnoModal", "2026", 3.5)
            
            opcoes_projeto = driver.find_elements(By.CSS_SELECTOR, "#PlanoModal option")
            
            projetos_desta_unidade = {}
            for opt in opcoes_projeto:
                valor = opt.get_attribute("value")
                texto = opt.text.strip()
                if valor and texto and texto != "Selecione...":
                    projetos_desta_unidade[texto] = valor
                    
            base_de_dados[f"ER_{id_er}"] = projetos_desta_unidade
            
            # SALVAMENTO INCREMENTAL: Salva no HD a cada volta!
            with open("base_dados_projetos.json", "w", encoding="utf-8") as f:
                json.dump(base_de_dados, f, ensure_ascii=False, indent=4)
                
        except Exception as e:
            print(f"⚠️ Erro no site ao mapear a Unidade {id_er}. O robô vai pular, dar F5 e continuar...")
            driver.refresh()
            time.sleep(5)
        
    print("Mapeamento Concluído! Arquivo 'base_dados_projetos.json' gerado com sucesso.")
    driver.quit()

if __name__ == "__main__":
    rodar_mapeador()