from core.services.services import processar_arquivo_xml, criar_planilhas_por_empresa
from core.data import dataBase as db

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
import glob

import sys 

def resource_path(relative_path):
    
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class AppJiraParser(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Processador de Chamados Jira")
        icone_path = resource_path('core/assets/icon_app.ico')
        self.iconbitmap(icone_path)
        self.geometry("700x550") 
        ctk.set_appearance_mode("System")
        

        self.caminho_selecionado = "" 
        self.pasta_destino = ""

        # Layout com grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(4, weight=1) 

     
        self.frame_modo = ctk.CTkFrame(self)
        self.frame_modo.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        self.frame_modo.grid_columnconfigure(0, weight=1)
        self.frame_modo.grid_columnconfigure(4, weight=1)

        
        self.lbl_modo = ctk.CTkLabel(self.frame_modo, text="Modo de Seleção:")
        self.lbl_modo.grid(row=0, column=0, padx=(10, 15), pady=10)
        
        self.modo_selecao = ctk.StringVar(value="pasta") 

        self.radio_pasta = ctk.CTkRadioButton(self.frame_modo, text="Analisar Pasta Inteira", variable=self.modo_selecao, value="pasta" )
        self.radio_pasta.grid(row=0, column=1, padx=15, pady=10) 
        
        self.radio_arquivo = ctk.CTkRadioButton(self.frame_modo, text="Analisar Arquivo Único", variable=self.modo_selecao, value="arquivo")
        self.radio_arquivo.grid(row=0, column=3, padx=15, pady=10)

     
        self.frame_selecao = ctk.CTkFrame(self)
        self.frame_selecao.grid(row=1, column=0, padx=10, pady=5, sticky="ew")
        self.frame_selecao.grid_columnconfigure(0, weight=1)
       
        self.label_origem = ctk.CTkLabel(self.frame_selecao, text="Nenhum arquivo ou pasta de origem selecionado", text_color="gray", anchor="w")
        self.label_origem.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.btn_selecionar = ctk.CTkButton(self.frame_selecao, text="Selecionar Origem...", command=self.abrir_dialogo_selecao)
        self.btn_selecionar.grid(row=0, column=1, padx=10, pady=10)


        self.frame_destino = ctk.CTkFrame(self)
        self.frame_destino.grid(row=2, column=0, padx=10, pady=5, sticky="ew")
        self.frame_destino.grid_columnconfigure(0, weight=1)
   
        self.label_destino = ctk.CTkLabel(self.frame_destino, text="Destino Padrão: pasta 'Relatorios_Jira'", text_color="gray", anchor="w")
        self.label_destino.grid(row=0, column=0, padx=10, pady=10, sticky="ew")
        self.btn_destino = ctk.CTkButton(self.frame_destino, text="Selecionar Destino...", command=self.selecionar_pasta_destino)
        self.btn_destino.grid(row=0, column=1, padx=10, pady=10)



        self.btn_selecionar = ctk.CTkButton(self.frame_selecao, text="Selecionar...", command=self.abrir_dialogo_selecao,)
        self.btn_selecionar.grid(row=0, column=1, padx=10, pady=10)

        # Botão de iniciar
        self.btn_iniciar = ctk.CTkButton(self, text="Iniciar Processamento", command=self.iniciar_processo_thread, height=40)
        self.btn_iniciar.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

        # Painel de Log
        self.log_area = ctk.CTkTextbox(self, state="disabled", wrap="word")
        self.log_area.grid(row=4, column=0, padx=10, pady=10, sticky="nsew")

    def adicionar_log(self, mensagem):
        self.log_area.configure(state="normal")
        self.log_area.insert("end", mensagem + "\n")
        self.log_area.configure(state="disabled")
        self.log_area.see("end")


    def abrir_dialogo_selecao(self):
        modo = self.modo_selecao.get()
        caminho = ""

        if modo == "pasta":
            caminho = filedialog.askdirectory(title="Selecione a pasta com os arquivos XML")
        else:
            caminho = filedialog.askopenfilename(title="Selecione um arquivo XML", filetypes=[("Arquivos XML", "*.xml"), ("Todos os arquivos", "*.*")])
        
        if caminho:
            self.caminho_selecionado = caminho
         
            self.label_origem.configure(text=caminho, text_color="black") 
            self.adicionar_log(f"Origem definida: {caminho}")

    def selecionar_pasta_destino(self):
        caminho = filedialog.askdirectory(title="Selecione a pasta para salvar os relatórios")
        if caminho:
            self.pasta_destino = caminho
            # Apenas atualizamos o texto do Label e sua cor
            self.label_destino.configure(text=caminho, text_color="black") 
            self.adicionar_log(f"Pasta de destino definida: {caminho}")

    def iniciar_processo_thread(self):
        if not self.caminho_selecionado:
            messagebox.showwarning("Aviso", "Por favor, selecione um arquivo ou pasta antes de iniciar.")
            return

        self.btn_selecionar.configure(state="disabled")
        self.radio_pasta.configure(state="disabled")
        self.radio_arquivo.configure(state="disabled")
        self.btn_iniciar.configure(state="disabled", text="Processando...")
        
        self.log_area.configure(state="normal")
        self.log_area.delete("1.0", "end")
        self.log_area.configure(state="disabled")
        
        thread = threading.Thread(target=self.executar_processamento, args=(self.caminho_selecionado,))
        thread.start()

    def executar_processamento(self, caminho):
        try:
            log_callback = self.adicionar_log
            # map_cto_agregado = {cto[0]: [] for cto in db.list_of_Cto}
            map_cto_agregado = dict.fromkeys(db.list_of_Cto[0], [])  


            lista_de_arquivos = []
            if os.path.isdir(caminho):
                log_callback(f"Modo Pasta selecionado. Procurando arquivos .xml em '{os.path.basename(caminho)}'...")
                lista_de_arquivos = glob.glob(os.path.join(caminho, '*.xml'))
            elif os.path.isfile(caminho):
                log_callback(f"Modo Arquivo Único selecionado.")
                lista_de_arquivos = [caminho] # A lista contém apenas um arquivo
            if not lista_de_arquivos:
                log_callback("[AVISO] Nenhum arquivo .xml válido para processar.")
            else:
                log_callback(f"Encontrados {len(lista_de_arquivos)} arquivo(s). Iniciando...")
                for arquivo_xml in lista_de_arquivos:
                    processar_arquivo_xml(arquivo_xml, map_cto_agregado, log_callback)
                
                criar_planilhas_por_empresa(map_cto_agregado, log_callback, self.pasta_destino)
                log_callback("\nProcesso concluído com sucesso!")
                messagebox.showinfo("Sucesso", "O processamento foi concluído e os relatórios foram gerados!")

        except Exception as e:
            log_callback(f"\n[ERRO FATAL] Ocorreu um erro inesperado: {e}")
            messagebox.showerror("Erro", f"Ocorreu um erro: {e}")
        finally:
            self.reabilitar_botoes()

    def reabilitar_botoes(self):
        self.btn_selecionar.configure(state="normal")
        self.radio_pasta.configure(state="normal")
        self.radio_arquivo.configure(state="normal")
        self.btn_iniciar.configure(state="normal", text="Iniciar Processamento")