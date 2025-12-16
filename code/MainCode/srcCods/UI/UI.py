import customtkinter as ctk
from tkinter import filedialog, messagebox
import webbrowser
import sys
import threading
import queue
import os
import pathlib as pl

# === Ajuste de caminhos (funciona dentro do .exe também) ===
# Import correto
import funcs   # AGORA FUNCIONA



# ==========================================================
# DualOutput — captura PRINT mesmo no EXE
# ==========================================================
class DualOutput:
    def __init__(self, textbox):
        self.textbox = textbox
        self.original_stdout = sys.stdout
        self.queue = queue.Queue()
        self.running = True

        sys.stdout = self  # redireciona prints

        threading.Thread(
            target=self._update_textbox, daemon=True
        ).start()

    def write(self, text):
        try:
            self.original_stdout.write(text)
            self.original_stdout.flush()
        except:
            pass

        if text.strip():
            self.queue.put(text)

    def flush(self):
        pass

    def _update_textbox(self):
        while self.running:
            try:
                text = self.queue.get(timeout=0.1)
                self.textbox.after(0, lambda: self._append(text))
            except queue.Empty:
                continue

    def _append(self, text):
        if self.textbox.get("1.0", "end").strip() == "Os registros aparecerão aqui":
            self.textbox.delete("1.0", "end")

        if not text.endswith("\n"):
            text += "\n"

        self.textbox.insert("end", text)
        self.textbox.see("end")


# ==========================================================
#                       APLICATIVO
# ==========================================================
class AppMain(ctk.CTk):  
    def __init__(self): 
        super().__init__()

        largura = self.winfo_screenwidth()
        altura = self.winfo_screenheight()
        self.geometry(f"{largura-200}x{altura-200}+100+50")
        self.title("Tabela Salarial - Generator Tool")

        ctk.set_appearance_mode("dark")
        self.purple = "#780CFB"

        # Variáveis ---------------------#
        try:
            self.token = funcs.extrair_token()
            self.TabelasOpts = funcs.listarJson(self.token)
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao carregar tabelas: {e}")
            self.TabelasOpts = []
        #---------------------------------#
        self.TabelasNomes = [item["names"] for item in self.TabelasOpts]
        self.Tabelasids = [item["ids"] for item in self.TabelasOpts]
        self.mapa_nome_id = {item["names"]: item["ids"] for item in self.TabelasOpts}

        self.path_output = ctk.StringVar()
        self.nome_tabela = ctk.StringVar()
        self.nome_aba = ctk.StringVar()
        self.criar_aba = ctk.BooleanVar(value=False)
        self.tabela_id = None  # evitar crash

        # ---------------- UI -----------------
        titulo = ctk.CTkLabel(self, text="📊 Folhas Excel Generator",
                              font=ctk.CTkFont(size=22, weight="bold"))
        titulo.pack(pady=(10, 0))

        autor = ctk.CTkLabel(self, text="Criado por Yuri Bertola",
                             font=ctk.CTkFont(size=12),
                             cursor="hand2")
        autor.pack()
        autor.bind("<Button-1>", lambda e: webbrowser.open("https://www.linkedin.com/in/yuri-bertola"))

        inputs_frame = ctk.CTkFrame(self)
        inputs_frame.pack(fill="x", padx=20, pady=10)

        # Pasta de saída
        linha1 = ctk.CTkFrame(inputs_frame)
        linha1.pack(fill="x", pady=3)
        ctk.CTkLabel(linha1, text="Pasta de Saída:").pack(side="left", padx=10)
        ctk.CTkEntry(linha1, textvariable=self.path_output).pack(side="left", expand=True, fill="x", padx=10)
        ctk.CTkButton(linha1, text="Selecionar", command=self.select_folder,
                      fg_color=self.purple).pack(side="left")

        # Nome tabela
        linha2 = ctk.CTkFrame(inputs_frame)
        linha2.pack(fill="x", pady=3)
        ctk.CTkLabel(linha2, text="Nome da Tabela:").pack(side="left", padx=10)
        ctk.CTkEntry(linha2, textvariable=self.nome_tabela).pack(side="left", expand=True, fill="x", padx=10)

        # Criar aba
        linha3 = ctk.CTkFrame(inputs_frame)
        linha3.pack(fill="x", pady=3)
        ctk.CTkCheckBox(linha3, text="Criar nova aba", variable=self.criar_aba).pack(side="left", padx=10)
        ctk.CTkEntry(linha3, placeholder_text="Nome da aba", textvariable=self.nome_aba)\
            .pack(side="left", expand=True, fill="x", padx=10)

        # Dropdown Senior
        linha4 = ctk.CTkFrame(inputs_frame)
        linha4.pack(fill="x", pady=3)
        ctk.CTkLabel(linha4, text="Tabela Senior:").pack(side="left", padx=10)
        self.dropdown = ctk.CTkComboBox(linha4, values=self.TabelasNomes, command=self.on_select)
        self.dropdown.pack(side="left", expand=True, fill="x", padx=10)

       # Linha 5 – botões lado a lado
        linha5 = ctk.CTkFrame(inputs_frame)
        linha5.pack(fill="x", pady=8)

        btn1 = ctk.CTkButton(
            linha5,
            text="Gerar Excel 🚀",
            fg_color="#FF4000",
            height=40,
            command=self.executar_geracao
        )
        btn1.pack(side="left", expand=True, fill="x", padx=5)

        btn2 = ctk.CTkButton(
            linha5,
            text="Gerar Todas as Tabelas 🚀",
            fg_color="#6200FF",
            height=40,
            command=self.executarTodas_thread
        )
        btn2.pack(side="left", expand=True, fill="x", padx=5)

        btn3 = ctk.CTkButton(
                linha5,
                text="Gerar Todas as Tabelas em Sequencia 🚀",
                fg_color="#000000",
                height=40,
                command=lambda: self.executarTodas_thread(onf=True)
        )
        btn3.pack(side="left", expand=True, fill="x", padx=5)

        # Log
        self.logbox = ctk.CTkTextbox(self, height=300)
        self.logbox.pack(fill="both", expand=True, padx=20, pady=10)
        self.logbox.insert("1.0", "Os registros aparecerão aqui")

        self.dual = DualOutput(self.logbox)


    # ---------------------- FUNÇÕES ----------------------


    def on_select(self, nome):
        self.tabela_id = self.mapa_nome_id[nome]
        print(f"Selecionado → {nome} | ID = {self.tabela_id}")

    def select_folder(self):
        pasta = filedialog.askdirectory()
        if pasta:
            self.path_output.set(pasta)
            print("📁 Pasta selecionada:", pasta)

    def validar_inputs(self):
        if not self.path_output.get():
            messagebox.showerror("Erro", "Selecione a pasta de saída.")
            return False

        if not self.nome_tabela.get():
            messagebox.showerror("Erro", "Digite o nome da tabela.")
            return False

        if self.criar_aba.get() and not self.nome_aba.get():
            messagebox.showerror("Erro", "Digite o nome da aba.")
            return False

        if self.tabela_id is None:
            messagebox.showerror("Erro", "Selecione uma tabela Senior.")
            return False

        return True

    def executar_geracao(self):
        if not self.validar_inputs():
            return

        try:
            funcs.criartabela(
                nomeAba=self.nome_aba.get(),
                nometabela=self.nome_tabela.get(),
                CaminhoPasta=self.path_output.get(),
                tabela_id=self.tabela_id,
                aba=self.criar_aba.get()
            )
            funcs.LINHAS.clear()

        except PermissionError:
            messagebox.showerror(
                "Arquivo Aberto",
                "Feche o arquivo Excel antes de gerar."
            )
            funcs.LINHAS.clear()

        except Exception as e:
            messagebox.showerror("Erro inesperado", str(e))

    def mostrar_loading(self, texto="Processando..."):
        self.loading = ctk.CTkToplevel(self)
        self.loading.title("Aguarde")
        self.loading.geometry("300x120")
        self.loading.resizable(False, False)
        self.loading.grab_set()       # bloqueia a janela principal

        ctk.CTkLabel(
            self.loading,
            text=texto,
            font=ctk.CTkFont(size=16, weight="bold")
        ).pack(pady=20)

        self.loading_progress = ctk.CTkProgressBar(self.loading)
        self.loading_progress.pack(padx=20, pady=10)
        self.loading_progress.configure(mode="indeterminate")
        self.loading_progress.start()

    def fechar_loading(self):
        try:
            self.loading_progress.stop()
            self.loading.destroy()
        except:
            pass

    def executarTodas_thread(self, onf=False):
        threading.Thread(
            target=lambda: self._executarTodas_wrapper(onf = onf),
            daemon=True
        ).start()
        self.mostrar_loading("Gerando todas as tabelas...")

    def _executarTodas_wrapper(self, onf=False):
        try:
            self.executarTodasAsFolhas(onef=onf)   # NÃO existe parâmetro 'onef'
        except Exception as e:
            messagebox.showerror("Erro", str(e))
        finally:
            self.after(0, self.fechar_loading)


    def executarTodasAsFolhas(self, onef=False):
            KpA = True
            funcs.LINHAS.clear()
            print("GERANDO TODAS AS TABELAS")
            print("----------------------------------")
            if onef == True:
                print("Gerando Tabela Sequencial")
                criar_aba = False
                nomeTabelaAtual = ""
                funcs.criartabela(
                    nomeAba=nomeTabelaAtual,
                    nometabela=self.nome_tabela.get(),
                    CaminhoPasta=self.path_output.get(),
                    tabela_id=self.Tabelasids,
                    aba=criar_aba,
                    keepalive=True
                    )
                funcs.LINHAS.clear()
        
            else:
             for idx,id in enumerate(self.Tabelasids):
                criar_aba = True
                nomeTabelaAtual = self.TabelasNomes[idx]

                try:
                    print("Gerando Tabela Por Abas")
                    funcs.criartabela(
                        nomeAba=nomeTabelaAtual,
                        nometabela=self.nome_tabela.get(),
                        CaminhoPasta=self.path_output.get(),
                        tabela_id=self.Tabelasids,
                        aba=True,
                        keepalive=True
                        )
                    if onef == False or KpA == True:
                        funcs.LINHAS.clear()
                        KpA = False
                        break
                except PermissionError:
                    messagebox.showerror(
                        "Arquivo Aberto",
                        "Feche o arquivo Excel antes de gerar."
                    )
                    if onef == False:
                        funcs.LINHAS.clear()
            funcs.LINHAS.clear()

        

# ==========================================================
# MAIN
# ==========================================================
if __name__ == "__main__":
    app = AppMain()
    app.mainloop()
