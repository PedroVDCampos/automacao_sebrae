import os
import sys
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from core.orquestrador import processar_tudo
from utils.updater import verificar_atualizacao

# Configuração da UI
ctk.set_appearance_mode("Dark") 
ctk.set_default_color_theme("blue") 

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("RAE Turbo")
        
        try:
            self.iconbitmap(resource_path("assets/icone.ico"))
        except:
            pass

        # Criação do evento de cancelamento (Thread-safe)
        self.evento_cancelar = threading.Event()

        # Layout
        self.frame_principal = ctk.CTkFrame(self, fg_color="transparent")
        self.frame_principal.grid(row=0, column=0, padx=30, pady=30, sticky="nsew")
        self.frame_principal.grid_columnconfigure(0, weight=1)

        self.titulo = ctk.CTkLabel(self.frame_principal, text="RAE Turbo", font=ctk.CTkFont(size=24, weight="bold"))
        self.titulo.grid(row=0, column=0, pady=(0, 30))

        # Inputs (Origem, Destino, Data)
        self.lbl_origem = ctk.CTkLabel(self.frame_principal, text="Origem (Downloads):", font=ctk.CTkFont(size=14))
        self.lbl_origem.grid(row=1, column=0, sticky="w", pady=(0, 5))

        self.frame_origem = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        self.frame_origem.grid(row=2, column=0, sticky="ew", pady=(0, 15))
        self.frame_origem.grid_columnconfigure(0, weight=1)

        self.entrada_origem = ctk.CTkEntry(self.frame_origem, height=35, placeholder_text="Caminho da pasta de downloads...")
        self.entrada_origem.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.btn_origem = ctk.CTkButton(self.frame_origem, text="Procurar", width=100, height=35, fg_color="gray50", hover_color="gray40", command=self.selecionar_origem)
        self.btn_origem.grid(row=0, column=1)

        self.lbl_destino = ctk.CTkLabel(self.frame_principal, text="Destino (Sebrae_Organizados):", font=ctk.CTkFont(size=14))
        self.lbl_destino.grid(row=3, column=0, sticky="w", pady=(0, 5))

        self.frame_destino = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        self.frame_destino.grid(row=4, column=0, sticky="ew", pady=(0, 20))
        self.frame_destino.grid_columnconfigure(0, weight=1)

        self.entrada_destino = ctk.CTkEntry(self.frame_destino, height=35, placeholder_text="Caminho da pasta destino...")
        self.entrada_destino.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self.btn_destino = ctk.CTkButton(self.frame_destino, text="Procurar", width=100, height=35, fg_color="gray50", hover_color="gray40", command=self.selecionar_destino)
        self.btn_destino.grid(row=0, column=1)

        self.lbl_data = ctk.CTkLabel(self.frame_principal, text="Processar arquivos a partir de (DD/MM/AAAA):", font=ctk.CTkFont(size=14))
        self.lbl_data.grid(row=5, column=0, sticky="w", pady=(0, 5))

        self.entrada_data = ctk.CTkEntry(self.frame_principal, width=150, height=35)
        self.entrada_data.insert(0, "01/04/2026")
        self.entrada_data.grid(row=6, column=0, sticky="w", pady=(0, 30))

        # Botões
        self.frame_botoes = ctk.CTkFrame(self.frame_principal, fg_color="transparent")
        self.frame_botoes.grid(row=7, column=0, sticky="ew", pady=(10, 0))
        self.frame_botoes.grid_columnconfigure((0, 1), weight=1)

        self.btn_iniciar = ctk.CTkButton(self.frame_botoes, text="Iniciar Automação", font=ctk.CTkFont(size=16, weight="bold"), height=50, command=self.iniciar)
        self.btn_iniciar.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        self.btn_cancelar = ctk.CTkButton(self.frame_botoes, text="Cancelar", font=ctk.CTkFont(size=16, weight="bold"), height=50, fg_color="#E90C0C", hover_color="#922B21", state="disabled", command=self.cancelar)
        self.btn_cancelar.grid(row=0, column=1, sticky="ew", padx=(5, 0))

        # Checar atualizações de forma segura para a Interface Gráfica (espera 1 segundo)
        self.after(1000, verificar_atualizacao)

    def selecionar_origem(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.entrada_origem.delete(0, ctk.END)
            self.entrada_origem.insert(0, pasta)

    def selecionar_destino(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.entrada_destino.delete(0, ctk.END)
            self.entrada_destino.insert(0, pasta)

    def iniciar(self):
        origem = self.entrada_origem.get()
        destino = self.entrada_destino.get()
        data_corte = self.entrada_data.get()

        if not origem or not destino or not data_corte:
            messagebox.showwarning("Atenção", "Preencha todas as pastas e a data de corte!")
            return

        self.evento_cancelar.clear()
        
        # Cria um "sinal de trânsito" para pausar o robô
        self.evento_login = threading.Event()
        self.evento_login.clear()

        self.btn_iniciar.configure(text="Iniciando Navegador...", state="disabled")
        self.btn_cancelar.configure(text="Cancelar Automação", state="normal", fg_color="#C0392B")
        
        # Inicia a Thread
        threading.Thread(target=self.rodar_background, args=(origem, destino, data_corte), daemon=True).start()

    def mostrar_aviso_login(self):
        messagebox.showinfo("Ação Necessária", "1. O navegador foi aberto.\n2. Faça o login no RAE.\n3. Quando estiver na tela de 'Pesquisa Clientes', clique em OK nesta janela para o robô começar.")
        self.evento_login.set() # Sinal Verde! Libera o robô para trabalhar.
        self.btn_iniciar.configure(text="Robô Trabalhando...")

    def callback_pausa_login(self):
        # Pede para a interface exibir a caixa de mensagem com segurança
        self.after(0, self.mostrar_aviso_login)
        # O robô senta e espera (0% de CPU) até você clicar em OK!
        self.evento_login.wait() 

    def rodar_background(self, origem, destino, data_corte):
        # Passamos a função callback para o orquestrador usar
        resultado = processar_tudo(origem, destino, data_corte, self.evento_cancelar, self.callback_pausa_login)
        
        # Devolve o resultado para a tela final
        self.after(0, self.finalizar_interface, resultado)

    def finalizar_interface(self, resultado):
        self.btn_iniciar.configure(text="Iniciar Automação", state="normal")
        self.btn_cancelar.configure(text="Cancelar", state="disabled", fg_color="gray50")

        if resultado["status"] == "erro":
            messagebox.showerror("Erro", resultado["msg"])
        elif resultado["status"] == "erro_fatal":
            messagebox.showerror("Erro Fatal no Navegador", resultado["msg"])
        elif resultado["status"] == "cancelado":
            messagebox.showinfo("Cancelado", "A operação foi interrompida pelo usuário.")
        elif resultado["status"] == "sucesso":
            erros = resultado["erros"]
            if len(erros) > 0:
                lista_erros = "\n".join(erros)
                messagebox.showwarning("Atenção: Cadastros Pendentes", f"Os seguintes CNPJs tiveram erro:\n\n{lista_erros}")
            else:
                messagebox.showinfo("Sucesso!", f"Processo concluído!\n\n{resultado['arquivos']} arquivo(s) organizado(s) e registrado(s).")
    def cancelar(self):
        self.evento_cancelar.set()
        self.btn_cancelar.configure(text="Cancelando...", state="disabled", fg_color="gray50")

if __name__ == "__main__":
    app = App()
    app.mainloop()