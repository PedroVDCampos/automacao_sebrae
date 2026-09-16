import os, sys, threading, subprocess
from datetime import datetime
import customtkinter as ctk
from tkinter import filedialog, messagebox
from core.orquestrador import processar_tudo
from utils.paths import caminho_historico, resource_path
from utils.updater import verificar_atualizacao
from version import VERSAO_ATUAL

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__(); self.title(f"Sebrae Turbo - v{VERSAO_ATUAL}"); self.geometry("620x610"); self.minsize(620,610)
        try:self.iconbitmap(resource_path("assets/icone.ico"))
        except Exception:pass
        self.evento_cancelar=threading.Event(); self.processando=False; self._montar_ui()
        self.after(1500, lambda: threading.Thread(target=verificar_atualizacao,daemon=True).start())
    def _montar_ui(self):
        self.grid_columnconfigure(0,weight=1)
        ctk.CTkLabel(self,text="Sebrae Turbo",font=ctk.CTkFont(size=28,weight="bold")).grid(row=0,column=0,pady=(25,4))
        ctk.CTkLabel(self,text="Organizador automatizado de documentos",font=ctk.CTkFont(size=14)).grid(row=1,column=0,pady=(0,20))
        self._campo("Pasta de origem:","entrada_origem",2); self._campo("Pasta de destino:","entrada_destino",4)
        ctk.CTkLabel(self,text="Processar arquivos a partir de (DD/MM/AAAA):",anchor="w").grid(row=6,column=0,sticky="w",padx=35,pady=5)
        self.entrada_data=ctk.CTkEntry(self,width=170,height=38); self.entrada_data.insert(0,datetime.now().strftime("%d/%m/%Y")); self.entrada_data.grid(row=7,column=0,sticky="w",padx=35,pady=(0,18))
        self.progress=ctk.CTkProgressBar(self,height=12); self.progress.set(0); self.progress.grid(row=8,column=0,sticky="ew",padx=35,pady=(0,8))
        self.lbl_status=ctk.CTkLabel(self,text="Pronto para organizar os PDFs.",anchor="w"); self.lbl_status.grid(row=9,column=0,sticky="w",padx=35,pady=(0,10))
        self.btn_iniciar=ctk.CTkButton(self,text="Organizar PDFs",height=48,font=ctk.CTkFont(weight="bold"),command=self.iniciar); self.btn_iniciar.grid(row=10,column=0,sticky="ew",padx=35,pady=5)
        self.btn_cancelar=ctk.CTkButton(self,text="Cancelar",height=40,fg_color="#C0392B",hover_color="#A93226",state="disabled",command=self.cancelar); self.btn_cancelar.grid(row=11,column=0,sticky="ew",padx=35,pady=5)
        f=ctk.CTkFrame(self,fg_color="transparent"); f.grid(row=12,column=0,sticky="ew",padx=35,pady=(8,20)); f.columnconfigure((0,1),weight=1)
        ctk.CTkButton(f,text="Histórico",command=self.abrir_historico).grid(row=0,column=0,sticky="ew",padx=(0,5)); ctk.CTkButton(f,text="Sobre",command=self.abrir_sobre).grid(row=0,column=1,sticky="ew",padx=(5,0))
    def _campo(self,label,var,row):
        ctk.CTkLabel(self,text=label).grid(row=row,column=0,sticky="w",padx=35,pady=5); f=ctk.CTkFrame(self,fg_color="transparent"); f.grid(row=row+1,column=0,sticky="ew",padx=35,pady=(0,12)); f.columnconfigure(0,weight=1)
        e=ctk.CTkEntry(f,height=38); e.grid(row=0,column=0,sticky="ew",padx=(0,10)); setattr(self,var,e); ctk.CTkButton(f,text="Abrir",width=85,command=lambda:self._selecionar(e)).grid(row=0,column=1)
    def _selecionar(self,e):
        p=filedialog.askdirectory()
        if p:e.delete(0,ctk.END);e.insert(0,p)
    def iniciar(self):
        o,d,dt=self.entrada_origem.get().strip(),self.entrada_destino.get().strip(),self.entrada_data.get().strip()
        if not o or not d or not dt:return messagebox.showwarning("Atenção","Preencha todos os campos.")
        try:datetime.strptime(dt,"%d/%m/%Y")
        except ValueError:return messagebox.showwarning("Data inválida","Use DD/MM/AAAA.")
        if not os.path.isdir(o):return messagebox.showerror("Erro","A pasta de origem não existe.")
        if os.path.abspath(o)==os.path.abspath(d):return messagebox.showwarning("Atenção","Origem e destino precisam ser diferentes.")
        self.evento_cancelar.clear();self.processando=True;self.btn_iniciar.configure(text="Organizando...",state="disabled");self.btn_cancelar.configure(state="normal");self.progress.set(0)
        threading.Thread(target=self.rodar,args=(o,d,dt),daemon=True).start()
    def rodar(self,o,d,dt):
        def prog(i,total,nome,res):self.after(0,self.atualizar,i,total,nome,res)
        r=processar_tudo(o,d,dt,self.evento_cancelar,prog,None);self.after(0,self.finalizar,r)
    def atualizar(self,i,total,nome,res):
        self.progress.set(i/total if total else 0);self.lbl_status.configure(text=f"{i}/{total} — {nome}\nOrganizados: {res.get('arquivos_organizados',0)} | Erros: {res.get('arquivos_com_erro',0)}")
    def finalizar(self,r):
        self.processando=False;self.btn_iniciar.configure(text="Organizar PDFs",state="normal");self.btn_cancelar.configure(state="disabled")
        s=r.get("status");res=r.get("resumo",{})
        if s=="sucesso":
            self.progress.set(1);self.lbl_status.configure(text="Processo concluído.");messagebox.showinfo("Organização concluída",f"PDFs encontrados: {res.get('pdfs_encontrados',0)}\nOrganizados: {res.get('arquivos_organizados',0)}\nJá existentes: {res.get('arquivos_ja_existentes',0)}\nCom erro: {res.get('arquivos_com_erro',0)}\nTempo total: {res.get('duracao_segundos',0):.2f}s")
        elif s=="cancelado":messagebox.showinfo("Cancelado","A organização foi interrompida.")
        else:messagebox.showerror("Erro",r.get("msg","Ocorreu um erro durante a organização."))
    def cancelar(self):
        self.evento_cancelar.set();self.btn_cancelar.configure(text="Cancelando...",state="disabled")
    def abrir_historico(self):
        p=caminho_historico()
        if not os.path.exists(p):return messagebox.showinfo("Histórico","Ainda não existem execuções registradas.")
        try:os.startfile(p) if sys.platform=="win32" else subprocess.Popen(["xdg-open",p])
        except Exception as e:messagebox.showerror("Histórico",str(e))
    def abrir_sobre(self):
        w=ctk.CTkToplevel(self);w.title("Sobre o Sebrae Turbo");w.geometry("500x380");w.resizable(False,False);w.grab_set()
        ctk.CTkLabel(w,text="Sebrae Turbo",font=ctk.CTkFont(size=26,weight="bold")).pack(pady=(30,5));ctk.CTkLabel(w,text=f"Versão {VERSAO_ATUAL}").pack(pady=(0,20))
        ctk.CTkLabel(w,text="Organizador automatizado de documentos de atendimento.\n\nIdentifica PDFs, extrai os dados necessários e organiza os documentos por ano, mês, tipo e cliente.\n\nO processamento documental é realizado localmente e não depende do sistema RAE.",wraplength=420,justify="left").pack(padx=35,pady=10)
        ctk.CTkButton(w,text="Fechar",command=w.destroy).pack(pady=20)

if __name__=="__main__":App().mainloop()
