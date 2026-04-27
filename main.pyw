import os
import sys
import json
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from core.orquestrador import processar_tudo, verificar_compatibilidade_chrome
from core.automacao_web import MAPA_UNIDADES, carregar_base_dados
from utils.updater import verificar_atualizacao

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

def _caminho_config():
    base = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base, "config_unidade.json")

def carregar_config() -> dict:
    try:
        with open(_caminho_config(), "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}

def salvar_config(dados: dict):
    try:
        with open(_caminho_config(), "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print(f"Aviso: não foi possível salvar configurações: {e}")

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class DialogConfiguracoes(ctk.CTkToplevel):
    def __init__(self, parent, config_atual: dict):
        super().__init__(parent)
        self.title("Configuração da Unidade")
        self.resizable(False, False)
        self.grab_set()
        self.resultado = None
        self.base_projetos = carregar_base_dados("base_dados_projetos.json")
        self.base_acoes = carregar_base_dados("base_dados_acoes.json")
        
        self.after(50, self._centralizar, parent)
        self._montar_ui(config_atual)

    def _montar_ui(self, config_atual):
        pad = {"padx": 16, "pady": (0, 12)}

        ctk.CTkLabel(self, text="Configuração da Unidade", font=ctk.CTkFont(size=18, weight="bold")).pack(padx=20, pady=(20, 16))

        frame_canal = ctk.CTkFrame(self)
        frame_canal.pack(fill="x", padx=20, pady=(0, 12))

        ctk.CTkLabel(frame_canal, text="Canal e local de realização", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(12, 8))
        
        ctk.CTkLabel(frame_canal, text="Canal:", anchor="w").pack(fill="x", padx=16)
        self.entrada_canal = ctk.CTkEntry(frame_canal, height=35)
        self.entrada_canal.pack(fill="x", **pad)
        self.entrada_canal.insert(0, config_atual.get("canal", ""))

        ctk.CTkLabel(frame_canal, text="Local de execução:", anchor="w").pack(fill="x", padx=16)
        self.entrada_local = ctk.CTkEntry(frame_canal, height=35)
        self.entrada_local.pack(fill="x", **pad)
        self.entrada_local.insert(0, config_atual.get("local_execucao", ""))

        frame_plano = ctk.CTkFrame(self)
        frame_plano.pack(fill="x", padx=20, pady=(0, 4))

        ctk.CTkLabel(frame_plano, text="Plano orçamentário", font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=16, pady=(12, 8))

        # CASCATA 1: Unidade
        ctk.CTkLabel(frame_plano, text="Escritório Regional ou Unidade:", anchor="w").pack(fill="x", padx=16)
        self.combo_unidade = ctk.CTkComboBox(frame_plano, values=list(MAPA_UNIDADES.keys()), height=35, command=self._ao_alterar_unidade, state="readonly")
        self.combo_unidade.pack(fill="x", **pad)
        
        # CASCATA 2: Projeto
        ctk.CTkLabel(frame_plano, text="Projeto:", anchor="w").pack(fill="x", padx=16)
        self.combo_projeto = ctk.CTkComboBox(frame_plano, values=["Selecione a unidade..."], height=35, command=self._ao_alterar_projeto, state="readonly")
        self.combo_projeto.pack(fill="x", **pad)

        # CASCATA 3: Ação
        ctk.CTkLabel(frame_plano, text="Ação:", anchor="w").pack(fill="x", padx=16)
        self.combo_acao = ctk.CTkComboBox(frame_plano, values=["Selecione o projeto..."], height=35, state="readonly")
        self.combo_acao.pack(fill="x", **pad)

        ctk.CTkLabel(frame_plano, text="Ano:", anchor="w").pack(fill="x", padx=16)
        self.entrada_ano = ctk.CTkEntry(frame_plano, height=35)
        self.entrada_ano.pack(fill="x", **pad)
        self.entrada_ano.insert(0, config_atual.get("ano", "2026"))

        if config_atual.get("unidade_nome"):
            self.combo_unidade.set(config_atual["unidade_nome"])
            self._ao_alterar_unidade(config_atual["unidade_nome"])
            if config_atual.get("projeto"): 
                self.combo_projeto.set(config_atual["projeto"])
                self._ao_alterar_projeto(config_atual["projeto"])
                if config_atual.get("acao"): self.combo_acao.set(config_atual["acao"])

        frame_btns = ctk.CTkFrame(self, fg_color="transparent")
        frame_btns.pack(fill="x", padx=20, pady=(8, 20))
        
        ctk.CTkButton(frame_btns, text="Cancelar", fg_color="gray40", command=self.destroy).pack(side="left", expand=True, padx=(0, 5))
        ctk.CTkButton(frame_btns, text="Confirmar e Salvar", font=ctk.CTkFont(weight="bold"), command=self._confirmar).pack(side="left", expand=True, padx=(5, 0))

    def _ao_alterar_unidade(self, escolha):
        id_er = MAPA_UNIDADES.get(escolha, "")
        projetos_dict = self.base_projetos.get(f"ER_{id_er}", {})
        lista_nomes = list(projetos_dict.keys())
        
        if lista_nomes:
            self.combo_projeto.configure(values=lista_nomes)
            self.combo_projeto.set(lista_nomes[0])
            self._ao_alterar_projeto(lista_nomes[0]) # Dispara o próximo nível da cascata
        else:
            self.combo_projeto.configure(values=["Nenhum projeto encontrado"])
            self.combo_projeto.set("Nenhum projeto encontrado")
            self.combo_acao.configure(values=["Nenhuma ação encontrada"])
            self.combo_acao.set("Nenhuma ação encontrada")

    def _ao_alterar_projeto(self, escolha_projeto):
        id_er = MAPA_UNIDADES.get(self.combo_unidade.get(), "")
        projetos_dict = self.base_projetos.get(f"ER_{id_er}", {})
        id_projeto = projetos_dict.get(escolha_projeto, "")

        acoes_dict = self.base_acoes.get(id_projeto, {})
        lista_acoes = list(acoes_dict.keys())

        if lista_acoes:
            self.combo_acao.configure(values=lista_acoes)
            self.combo_acao.set(lista_acoes[0])
        else:
            self.combo_acao.configure(values=["Nenhuma ação encontrada"])
            self.combo_acao.set("Nenhuma ação encontrada")

    def _centralizar(self, parent):
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")

    def _confirmar(self):
        self.resultado = {
            "canal": self.entrada_canal.get().strip(),
            "local_execucao": self.entrada_local.get().strip(),
            "unidade_nome": self.combo_unidade.get(),
            "ano": self.entrada_ano.get().strip(),
            "projeto": self.combo_projeto.get(),
            "acao": self.combo_acao.get(),
        }
        self.destroy()

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RAE Turbo")
        self.geometry("500x450")
        
        try: self.iconbitmap(resource_path("assets/icone.ico"))
        except: pass

        self.evento_cancelar = threading.Event()
        self._montar_ui()
        self.after(1000, verificar_atualizacao)

    def _montar_ui(self):
        self.grid_columnconfigure(0, weight=1)
        
        ctk.CTkLabel(self, text="RAE Turbo", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, pady=20)

        self._criar_campo_pasta("Origem:", "entrada_origem", 1)
        self._criar_campo_pasta("Destino:", "entrada_destino", 3)

        ctk.CTkLabel(self, text="Processar a partir de (DD/MM/AAAA):").grid(row=5, column=0, sticky="w", padx=30)
        self.entrada_data = ctk.CTkEntry(self, width=150)
        self.entrada_data.insert(0, "01/01/2026")
        self.entrada_data.grid(row=6, column=0, sticky="w", padx=30, pady=(0, 20))

        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Automação", height=50, font=ctk.CTkFont(weight="bold"), command=self.iniciar)
        self.btn_iniciar.grid(row=7, column=0, padx=30, pady=5, sticky="ew")

        self.btn_cancelar = ctk.CTkButton(self, text="Cancelar", height=40, fg_color="#C0392B", hover_color="#A93226", state="disabled", command=self.cancelar)
        self.btn_cancelar.grid(row=8, column=0, padx=30, pady=5, sticky="ew")

    def _criar_campo_pasta(self, label, var_name, row):
        ctk.CTkLabel(self, text=label).grid(row=row, column=0, sticky="w", padx=30)
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row+1, column=0, sticky="ew", padx=30, pady=(0, 15))
        frame.columnconfigure(0, weight=1)
        
        ent = ctk.CTkEntry(frame)
        ent.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        setattr(self, var_name, ent)
        
        ctk.CTkButton(frame, text="Abrir", width=80, command=lambda: self._selecionar_pasta(ent)).grid(row=0, column=1)

    def _selecionar_pasta(self, entrada):
        pasta = filedialog.askdirectory()
        if pasta:
            entrada.delete(0, ctk.END)
            entrada.insert(0, pasta)

    def iniciar(self):
        origem, destino, data = self.entrada_origem.get(), self.entrada_destino.get(), self.entrada_data.get()
        if not all([origem, destino, data]):
            messagebox.showwarning("Atenção", "Preencha todos os campos!")
            return

        comp = verificar_compatibilidade_chrome()
        if comp["status"] == "erro":
            messagebox.showerror("Erro", comp["msg"])
            return

        dialogo = DialogConfiguracoes(self, carregar_config())
        self.wait_window(dialogo)

        if dialogo.resultado:
            salvar_config(dialogo.resultado)
            self.evento_cancelar.clear()
            self.btn_iniciar.configure(text="Trabalhando...", state="disabled")
            self.btn_cancelar.configure(state="normal")
            
            threading.Thread(target=self.rodar_background, args=(origem, destino, data, dialogo.resultado), daemon=True).start()

    def rodar_background(self, o, d, dt, cfg):
        def pausa():
            self.after(0, lambda: self.btn_iniciar.configure(text="Continuar", state="normal"))

        res = processar_tudo(o, d, dt, self.evento_cancelar, pausa, cfg)
        self.after(0, lambda: self.finalizar(res))

    def finalizar(self, r):
        # 1. Reseta os botões caso o programa continue aberto
        self.btn_iniciar.configure(text="Iniciar Automação", state="normal")
        self.btn_cancelar.configure(text="Cancelar", state="disabled")
        
        # 2. Avalia o resultado que veio do robô
        status = r.get("status")
        
        if status == "sucesso":
            messagebox.showinfo("Sucesso", f"Processo concluído!\nErros: {len(r.get('erros', []))}")
            
        elif status == "erro":
            messagebox.showerror("Erro", r.get("msg", "Ocorreu um erro durante o processo."))
            
        elif status == "cancelado":
            messagebox.showinfo("Cancelado", "Automação interrompida com sucesso. O programa será fechado.")
            self.destroy() 

    def cancelar(self):
        self.evento_cancelar.set()
        self.btn_cancelar.configure(text="Cancelando...")

if __name__ == "__main__":
    app = App()
    app.mainloop()