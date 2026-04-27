from pathlib import Path

RAIZ_PROJETO = Path(__file__).resolve().parents[1]
PASTA_DATA = RAIZ_PROJETO / "data"
PASTA_DATA.mkdir(exist_ok=True)

import time
import json
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

def injetar_id(driver, id_select, valor_id, tempo_espera):
    script = f"var el = $('#{id_select}'); if(el.length) {{ el.val('{valor_id}').trigger('change'); if (el.hasClass('select2-hidden-accessible')) {{ el.trigger('change.select2'); }} }}"
    driver.execute_script(script)
    time.sleep(tempo_espera)

def rodar_mapeador_acoes():
    try:
        with open(PASTA_DATA / "base_dados_projetos.json", "r", encoding="utf-8") as f:
            dados_projetos = json.load(f)
    except Exception:
        print("❌ Arquivo 'base_dados_projetos.json' não encontrado.")
        return

    projetos_unicos = {} 
    for er_chave, projetos in dados_projetos.items():
        id_er = er_chave.replace("ER_", "")
        for nome_proj, id_proj in projetos.items():
            if id_proj not in projetos_unicos:
                projetos_unicos[id_proj] = {"nome": nome_proj, "er_id": id_er}

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))
    driver.get("https://atendimento.sp.sebrae.com.br/Acesso/Login?ReturnUrl=%2f")
    print("FAÇA O LOGIN e vá para a tela de atendimento...")

    while True:
        try:
            if driver.find_elements(By.ID, "UnidadeModal"): break
        except: pass
        time.sleep(1)

    base_de_acoes_hierarquica = {}
    contador, total = 1, len(projetos_unicos)

    for id_proj, info in projetos_unicos.items():
        nome_proj, id_er = info["nome"], info["er_id"]
        print(f"[{contador}/{total}] Raspando Ações do Projeto: {nome_proj}...")
        
        try:
            if contador % 15 == 0:
                driver.refresh()
                time.sleep(4)

            injetar_id(driver, "UnidadeModal", id_er, 1.5)
            injetar_id(driver, "AnoModal", "2026", 2)
            injetar_id(driver, "PlanoModal", id_proj, 3.5) 

            opcoes_acao = driver.find_elements(By.CSS_SELECTOR, "#AcaoModal option")
            
            # A MÁGICA AQUI: Salva as ações DENTRO da gaveta deste projeto específico!
            acoes_deste_projeto = {}
            for opt in opcoes_acao:
                valor = opt.get_attribute("value")
                texto = opt.text.strip()
                if valor and texto and texto != "Selecione...":
                    acoes_deste_projeto[texto] = valor

            base_de_acoes_hierarquica[id_proj] = acoes_deste_projeto

            with open(PASTA_DATA / "base_dados_acoes.json", "w", encoding="utf-8") as f:
                json.dump(base_de_acoes_hierarquica, f, ensure_ascii=False, indent=4)

        except Exception as e:
            print(f"⚠️ Erro ao raspar o projeto '{nome_proj}'. Pulando...")
            driver.refresh()
            time.sleep(4)
            
        contador += 1

    print(f"\n✅ Concluído! Novo 'base_dados_acoes.json' gerado com hierarquia.")
    driver.quit()

if __name__ == "__main__":
    rodar_mapeador_acoes()