import os
import subprocess
import sys
import threading
from datetime import datetime

import customtkinter as ctk
from tkinter import END, filedialog, messagebox

from core.orquestrador import processar_tudo
from utils.paths import caminho_historico, resource_path
from utils.updater import baixar_atualizacao, instalar_atualizacao, verificar_atualizacao
from version import VERSAO_ATUAL

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"RAE Turbo - v{VERSAO_ATUAL}")
        self.geometry("620x610")
        self.minsize(620, 610)

        try:
            self.iconbitmap(resource_path("assets/icone.ico"))
        except Exception:
            pass

        self.evento_cancelar = threading.Event()
        self.processando = False
        self._montar_ui()
        self.after(1500, self._verificar_atualizacao_em_background)

    def _montar_ui(self):
        self.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(
            self, text="RAE Turbo", font=ctk.CTkFont(size=28, weight="bold")
        ).grid(row=0, column=0, pady=(25, 4))
        ctk.CTkLabel(
            self,
            text="Organizador automatizado de documentos",
            font=ctk.CTkFont(size=14),
        ).grid(row=1, column=0, pady=(0, 20))

        self._campo("Pasta de origem:", "entrada_origem", 2)
        self._campo("Pasta de destino:", "entrada_destino", 4)

        ctk.CTkLabel(
            self,
            text="Processar arquivos a partir de (DD/MM/AAAA):",
            anchor="w",
        ).grid(row=6, column=0, sticky="w", padx=35, pady=5)
        self.entrada_data = ctk.CTkEntry(self, width=170, height=38)
        self.entrada_data.insert(0, datetime.now().strftime("%d/%m/%Y"))
        self.entrada_data.grid(row=7, column=0, sticky="w", padx=35, pady=(0, 18))

        self.progress = ctk.CTkProgressBar(self, height=12)
        self.progress.set(0)
        self.progress.grid(row=8, column=0, sticky="ew", padx=35, pady=(0, 8))

        self.lbl_status = ctk.CTkLabel(
            self, text="Pronto para organizar os PDFs.", anchor="w"
        )
        self.lbl_status.grid(row=9, column=0, sticky="w", padx=35, pady=(0, 10))

        self.btn_iniciar = ctk.CTkButton(
            self,
            text="Organizar PDFs",
            height=48,
            font=ctk.CTkFont(weight="bold"),
            command=self.iniciar,
        )
        self.btn_iniciar.grid(row=10, column=0, sticky="ew", padx=35, pady=5)

        self.btn_cancelar = ctk.CTkButton(
            self,
            text="Cancelar",
            height=40,
            fg_color="#C0392B",
            hover_color="#A93226",
            state="disabled",
            command=self.cancelar,
        )
        self.btn_cancelar.grid(row=11, column=0, sticky="ew", padx=35, pady=5)

        frame_botoes = ctk.CTkFrame(self, fg_color="transparent")
        frame_botoes.grid(row=12, column=0, sticky="ew", padx=35, pady=(8, 20))
        frame_botoes.columnconfigure((0, 1), weight=1)
        ctk.CTkButton(
            frame_botoes, text="Histórico", command=self.abrir_historico
        ).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        ctk.CTkButton(
            frame_botoes, text="Sobre", command=self.abrir_sobre
        ).grid(row=0, column=1, sticky="ew", padx=(5, 0))

    def _campo(self, label, var, row):
        ctk.CTkLabel(self, text=label).grid(
            row=row, column=0, sticky="w", padx=35, pady=5
        )
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=row + 1, column=0, sticky="ew", padx=35, pady=(0, 12))
        frame.columnconfigure(0, weight=1)
        entrada = ctk.CTkEntry(frame, height=38)
        entrada.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        setattr(self, var, entrada)
        ctk.CTkButton(
            frame, text="Abrir", width=85, command=lambda: self._selecionar(entrada)
        ).grid(row=0, column=1)

    def _selecionar(self, entrada):
        pasta = filedialog.askdirectory()
        if pasta:
            entrada.delete(0, END)
            entrada.insert(0, pasta)

    def iniciar(self):
        origem = self.entrada_origem.get().strip()
        destino = self.entrada_destino.get().strip()
        data = self.entrada_data.get().strip()

        if not origem or not destino or not data:
            return messagebox.showwarning("Atenção", "Preencha todos os campos.")

        try:
            datetime.strptime(data, "%d/%m/%Y")
        except ValueError:
            return messagebox.showwarning("Data inválida", "Use DD/MM/AAAA.")

        if not os.path.isdir(origem):
            return messagebox.showerror("Erro", "A pasta de origem não existe.")
        if os.path.abspath(origem) == os.path.abspath(destino):
            return messagebox.showwarning(
                "Atenção", "Origem e destino precisam ser diferentes."
            )

        self.evento_cancelar.clear()
        self.processando = True
        self.btn_iniciar.configure(text="Organizando...", state="disabled")
        self.btn_cancelar.configure(text="Cancelar", state="normal")
        self.progress.set(0)
        threading.Thread(
            target=self.rodar, args=(origem, destino, data), daemon=True
        ).start()

    def rodar(self, origem, destino, data):
        def progresso(i, total, nome, resultado):
            self.after(0, self.atualizar, i, total, nome, resultado)

        resultado = processar_tudo(
            origem, destino, data, self.evento_cancelar, progresso, None
        )
        self.after(0, self.finalizar, resultado)

    def atualizar(self, i, total, nome, resultado):
        self.progress.set(i / total if total else 0)
        self.lbl_status.configure(
            text=(
                f"{i}/{total} — {nome}\n"
                f"Organizados: {resultado.get('arquivos_organizados', 0)} | "
                f"Erros: {resultado.get('arquivos_com_erro', 0)}"
            )
        )

    def finalizar(self, resultado):
        self.processando = False
        self.btn_iniciar.configure(text="Organizar PDFs", state="normal")
        self.btn_cancelar.configure(text="Cancelar", state="disabled")

        status = resultado.get("status")
        resumo = resultado.get("resumo", {})
        if status == "sucesso":
            self.progress.set(1)
            self.lbl_status.configure(text="Processo concluído.")
            messagebox.showinfo(
                "Organização concluída",
                f"PDFs encontrados: {resumo.get('pdfs_encontrados', 0)}\n"
                f"Organizados: {resumo.get('arquivos_organizados', 0)}\n"
                f"Já existentes: {resumo.get('arquivos_ja_existentes', 0)}\n"
                f"Com erro: {resumo.get('arquivos_com_erro', 0)}\n"
                f"Tempo total: {resumo.get('duracao_segundos', 0):.2f}s",
            )
        elif status == "cancelado":
            messagebox.showinfo("Cancelado", "A organização foi interrompida.")
        else:
            messagebox.showerror(
                "Erro", resultado.get("msg", "Ocorreu um erro durante a organização.")
            )

    def cancelar(self):
        self.evento_cancelar.set()
        self.btn_cancelar.configure(text="Cancelando...", state="disabled")

    def abrir_historico(self):
        caminho = caminho_historico()
        if not os.path.exists(caminho):
            return messagebox.showinfo(
                "Histórico", "Ainda não existem execuções registradas."
            )
        try:
            if sys.platform == "win32":
                os.startfile(caminho)
            else:
                subprocess.Popen(["xdg-open", caminho])
        except Exception as erro:
            messagebox.showerror("Histórico", str(erro))

    def abrir_sobre(self):
        janela = ctk.CTkToplevel(self)
        janela.title("Sobre o RAE Turbo")
        janela.geometry("500x380")
        janela.resizable(False, False)
        janela.grab_set()

        ctk.CTkLabel(
            janela, text="RAE Turbo", font=ctk.CTkFont(size=26, weight="bold")
        ).pack(pady=(30, 5))
        ctk.CTkLabel(janela, text=f"Versão {VERSAO_ATUAL}").pack(pady=(0, 20))
        ctk.CTkLabel(
            janela,
            text=(
                "Organizador automatizado de documentos de atendimento.\n\n"
                "Identifica PDFs, extrai os dados necessários e organiza os documentos "
                "por ano, mês, tipo e cliente.\n\n"
                "O processamento documental é realizado localmente e não depende do "
                "sistema RAE."
            ),
            wraplength=420,
            justify="left",
        ).pack(padx=35, pady=10)
        ctk.CTkButton(janela, text="Fechar", command=janela.destroy).pack(pady=20)

    def _verificar_atualizacao_em_background(self):
        threading.Thread(
            target=verificar_atualizacao,
            args=(self._receber_atualizacao,),
            daemon=True,
        ).start()

    def _receber_atualizacao(self, info):
        # O callback da consulta pode ocorrer em thread secundária.
        self.after(0, self._perguntar_atualizacao, info)

    def _perguntar_atualizacao(self, info):
        if not info:
            return
        deseja = messagebox.askyesno(
            "Atualização disponível",
            f"Uma nova versão ({info['versao']}) do RAE Turbo foi encontrada.\n\n"
            "Deseja atualizar agora?",
        )
        if deseja:
            self.btn_iniciar.configure(state="disabled")
            self.lbl_status.configure(text=f"Baixando atualização {info['versao']}...")
            threading.Thread(
                target=self._baixar_atualizacao,
                args=(info,),
                daemon=True,
            ).start()

    def _baixar_atualizacao(self, info):
        try:
            caminho = baixar_atualizacao(info)
            self.after(0, self._instalar_atualizacao, caminho)
        except Exception as erro:
            self.after(0, self._erro_atualizacao, str(erro))

    def _instalar_atualizacao(self, caminho):
        try:
            instalar_atualizacao(caminho)
        except Exception as erro:
            self.btn_iniciar.configure(state="normal")
            self.lbl_status.configure(text="Pronto para organizar os PDFs.")
            messagebox.showerror(
                "Erro na atualização",
                f"Não foi possível atualizar o RAE Turbo.\n\n{erro}",
            )

    def _erro_atualizacao(self, mensagem):
        self.btn_iniciar.configure(state="normal")
        self.lbl_status.configure(text="Pronto para organizar os PDFs.")
        messagebox.showerror("Erro na atualização", mensagem)


if __name__ == "__main__":
    App().mainloop()
