import json
import os
import threading
import time

import customtkinter as ctk
from tkinter import filedialog, messagebox

from core.orquestrador import processar_tudo, verificar_compatibilidade_chrome
from core.automacao_web import MAPA_UNIDADES, carregar_base_dados
from utils.paths import appdata_dir, config_path, resource_path
from utils.updater import verificar_atualizacao
from utils.relatorio_execucao import (
    caminho_historico_execucoes,
    caminho_pasta_relatorios,
    caminho_relatorio_execucao,
    formatar_resumo_para_usuario,
)
from version import VERSAO_ATUAL

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


def carregar_config() -> dict:
    try:
        with open(config_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return {}


def salvar_config(dados: dict):
    try:
        with open(config_path(), "w", encoding="utf-8") as f:
            json.dump(dados, f, ensure_ascii=False, indent=2)
    except Exception as e:
        messagebox.showwarning("Configuração", f"Não foi possível salvar as configurações.\n\nMotivo: {e}")


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

        ctk.CTkLabel(frame_plano, text="Escritório Regional ou Unidade:", anchor="w").pack(fill="x", padx=16)
        self.combo_unidade = ctk.CTkComboBox(frame_plano, values=list(MAPA_UNIDADES.keys()), height=35, command=self._ao_alterar_unidade, state="readonly")
        self.combo_unidade.pack(fill="x", **pad)

        ctk.CTkLabel(frame_plano, text="Projeto:", anchor="w").pack(fill="x", padx=16)
        self.combo_projeto = ctk.CTkComboBox(frame_plano, values=["Selecione a unidade..."], height=35, command=self._ao_alterar_projeto, state="readonly")
        self.combo_projeto.pack(fill="x", **pad)

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
                if config_atual.get("acao"):
                    self.combo_acao.set(config_atual["acao"])

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
            self._ao_alterar_projeto(lista_nomes[0])
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
        if any(not valor for valor in self.resultado.values()):
            messagebox.showwarning("Configuração incompleta", "Preencha todos os campos da unidade antes de continuar.")
            return
        self.destroy()


class DialogSobre(ctk.CTkToplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.title("Sobre o RAE Turbo")
        self.geometry("520x430")
        self.resizable(False, False)
        self.grab_set()
        self.after(50, self._centralizar, parent)
        self._montar_ui()

    def _montar_ui(self):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="both", expand=True, padx=20, pady=20)

        ctk.CTkLabel(
            frame,
            text="RAE Turbo",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(pady=(18, 4))

        ctk.CTkLabel(frame, text=f"Versão: {VERSAO_ATUAL}", text_color="gray75").pack(pady=(0, 16))

        texto = (
            "Automação desktop para organização dos PDFs de atendimentos MEI "
            "e registro dos atendimentos no RAE.\n\n"
            "A ferramenta utiliza o login do próprio usuário autorizado, "
            "gera relatório final de execução e mantém histórico em CSV para acompanhamento de produtividade."
        )
        ctk.CTkLabel(frame, text=texto, justify="left", wraplength=440).pack(anchor="w", padx=20, pady=(0, 14))

        info = (
            f"Pasta de dados do app:\n{appdata_dir()}\n\n"
            f"Último relatório:\n{caminho_relatorio_execucao()}\n\n"
            f"Histórico CSV:\n{caminho_historico_execucoes()}"
        )
        ctk.CTkLabel(frame, text=info, justify="left", text_color="gray80", wraplength=440).pack(anchor="w", padx=20, pady=(0, 14))

        botoes = ctk.CTkFrame(frame, fg_color="transparent")
        botoes.pack(fill="x", padx=20, pady=(4, 18))

        ctk.CTkButton(botoes, text="Abrir pasta de relatórios", command=self._abrir_pasta_relatorios).pack(side="left", expand=True, padx=(0, 6))
        ctk.CTkButton(botoes, text="Fechar", fg_color="gray40", command=self.destroy).pack(side="left", expand=True, padx=(6, 0))

    def _abrir_pasta_relatorios(self):
        caminho = caminho_pasta_relatorios()
        try:
            os.startfile(caminho)
        except Exception as e:
            messagebox.showwarning("Sobre", f"Não foi possível abrir a pasta de relatórios.\n\n{e}")

    def _centralizar(self, parent):
        self.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - self.winfo_height()) // 2
        self.geometry(f"+{x}+{y}")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"RAE Turbo {VERSAO_ATUAL}")
        self.geometry("500x560")
        try:
            self.iconbitmap(resource_path("assets/icone.ico"))
        except Exception:
            pass
        self.evento_cancelar = threading.Event()
        self.evento_login_confirmado = threading.Event()
        self._montar_ui()
        self.after(1000, verificar_atualizacao)

    def _montar_ui(self):
        self.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self, text="RAE Turbo", font=ctk.CTkFont(size=24, weight="bold")).grid(row=0, column=0, pady=(20, 6))
        ctk.CTkLabel(self, text=VERSAO_ATUAL, text_color="gray70").grid(row=1, column=0, pady=(0, 12))
        self._criar_campo_pasta("Origem:", "entrada_origem", 2)
        self._criar_campo_pasta("Destino:", "entrada_destino", 4)

        ctk.CTkLabel(self, text="Processar a partir de (DD/MM/AAAA):").grid(row=6, column=0, sticky="w", padx=30)
        self.entrada_data = ctk.CTkEntry(self, width=150)
        self.entrada_data.insert(0, "01/01/2026")
        self.entrada_data.grid(row=7, column=0, sticky="w", padx=30, pady=(0, 16))

        self.label_status = ctk.CTkLabel(self, text="Pronto para iniciar.", text_color="gray80")
        self.label_status.grid(row=8, column=0, padx=30, pady=(0, 10), sticky="w")

        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Automação", height=50, font=ctk.CTkFont(weight="bold"), command=self.iniciar)
        self.btn_iniciar.grid(row=9, column=0, padx=30, pady=5, sticky="ew")

        self.btn_cancelar = ctk.CTkButton(self, text="Cancelar", height=40, fg_color="#C0392B", hover_color="#A93226", state="disabled", command=self.cancelar)
        self.btn_cancelar.grid(row=10, column=0, padx=30, pady=5, sticky="ew")

        self.btn_sobre = ctk.CTkButton(self, text="Sobre", height=35, fg_color="gray30", hover_color="gray25", command=self.abrir_sobre)
        self.btn_sobre.grid(row=11, column=0, padx=30, pady=(12, 5), sticky="ew")

    def _criar_campo_pasta(self, label, var_name, row):
        ctk.CTkLabel(self, text=label).grid(row=row, column=0, sticky="w", padx=30)
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row + 1, column=0, sticky="ew", padx=30, pady=(0, 15))
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
            messagebox.showwarning("Atenção", "Preencha todos os campos.")
            return
        comp = verificar_compatibilidade_chrome()
        if comp["status"] == "erro":
            messagebox.showerror("Erro", comp["msg"])
            return
        if comp["status"] == "aviso":
            continuar = messagebox.askyesno("Aviso", comp["msg"] + "\n\nDeseja continuar mesmo assim?")
            if not continuar:
                return
        dialogo = DialogConfiguracoes(self, carregar_config())
        self.wait_window(dialogo)
        if dialogo.resultado:
            salvar_config(dialogo.resultado)
            self.evento_cancelar.clear()
            self.evento_login_confirmado.clear()
            self.btn_iniciar.configure(text="Abrindo RAE...", state="disabled", command=self.iniciar)
            self.btn_cancelar.configure(state="normal")
            self.label_status.configure(text="Abrindo navegador e aguardando login...")
            threading.Thread(target=self.rodar_background, args=(origem, destino, data, dialogo.resultado), daemon=True).start()

    def confirmar_login(self):
        self.evento_login_confirmado.set()
        self.btn_iniciar.configure(text="Trabalhando...", state="disabled", command=self.iniciar)
        self.label_status.configure(text="Login confirmado. Processando atendimentos...")

    def rodar_background(self, o, d, dt, cfg):
        def aguardar_login():
            def preparar_botao():
                self.btn_iniciar.configure(text="Continuar após login", state="normal", command=self.confirmar_login)
                self.label_status.configure(text="Faça login no RAE e clique em 'Continuar após login'.")
                messagebox.showinfo(
                    "Login necessário",
                    "O navegador do RAE foi aberto.\n\n"
                    "1. Faça login normalmente.\n"
                    "2. Aguarde chegar à tela inicial/pesquisa.\n"
                    "3. Volte ao RAE Turbo e clique em 'Continuar após login'."
                )
            self.after(0, preparar_botao)
            while not self.evento_login_confirmado.is_set():
                if self.evento_cancelar.is_set():
                    return False
                time.sleep(0.2)
            return True
        res = processar_tudo(o, d, dt, self.evento_cancelar, aguardar_login, cfg)
        self.after(0, lambda: self.finalizar(res))

    def finalizar(self, r):
        self.btn_iniciar.configure(text="Iniciar Automação", state="normal", command=self.iniciar)
        self.btn_cancelar.configure(text="Cancelar", state="disabled")
        self.label_status.configure(text="Pronto para iniciar.")
        status = r.get("status")
        if status == "sucesso":
            mensagem = formatar_resumo_para_usuario(r.get("resumo", {}))
            messagebox.showinfo("Resumo da execução", mensagem)
        elif status in ("erro", "erro_fatal"):
            mensagem_erro = r.get("msg", "Ocorreu um erro durante o processo.")
            if r.get("resumo"):
                mensagem_erro += "\n\n" + formatar_resumo_para_usuario(r.get("resumo", {}))
            messagebox.showerror("Erro", mensagem_erro)
        elif status == "cancelado":
            mensagem = formatar_resumo_para_usuario(r.get("resumo", {})) if r.get("resumo") else "Automação interrompida com sucesso."
            messagebox.showinfo("Cancelado", mensagem)
        else:
            messagebox.showwarning("Atenção", f"Processo finalizado com status inesperado: {status}")

    def abrir_sobre(self):
        DialogSobre(self)

    def cancelar(self):
        self.evento_cancelar.set()
        self.evento_login_confirmado.set()
        self.btn_cancelar.configure(text="Cancelando...")
        self.label_status.configure(text="Cancelando automação...")


if __name__ == "__main__":
    app = App()
    app.mainloop()
