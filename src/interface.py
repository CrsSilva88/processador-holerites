import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
from processamento import processar_pdfs
from informe import processar_pdfs_informe
from utils import atualizar_barra_progresso


def iniciar_interface():
    root = tk.Tk()
    root.title("Processamento de PDFs")
    root.geometry("550x450")
    root.minsize(550, 450)

    canvas = tk.Canvas(root)
    scrollbar = ttk.Scrollbar(root, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )

    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    style = ttk.Style()
    style.theme_use('clam')
    style.configure('TButton', font=('Arial', 10))
    style.configure('TLabel', font=('Arial', 10))
    style.configure('TEntry', font=('Arial', 10))

    def selecionar_excel():
        caminho_excel = filedialog.askopenfilename(
            filetypes=[("Arquivos Excel", "*.xlsx")])
        caminho_excel_label.config(text=caminho_excel)

    def selecionar_pasta_pdfs():
        pasta_pdfs = filedialog.askdirectory()
        pasta_pdfs_label.config(text=pasta_pdfs)

    def selecionar_pasta_destino():
        pasta_destino = filedialog.askdirectory()
        pasta_destino_label.config(text=pasta_destino)

    def atualizar_campos_visiveis(*args):
        modo = modo_var.get()
        if modo == "Com Ano e Mês":
            label_ano.grid()
            entry_ano.grid()
            label_mes.grid()
            entry_mes.grid()
        else:
            label_ano.grid_remove()
            entry_ano.grid_remove()
            label_mes.grid_remove()
            entry_mes.grid_remove()

    def executar_processamento():
        caminho_excel = caminho_excel_label.cget("text")
        pasta_pdfs = pasta_pdfs_label.cget("text")
        pasta_destino = pasta_destino_label.cget("text")
        modo_selecionado = modo_var.get()

        if not caminho_excel or not pasta_pdfs or not pasta_destino:
            messagebox.showerror(
                "Erro", "Por favor, selecione todos os arquivos e pastas.")
            return

        atualizar_barra_progresso(barra_progresso, 0, 100)
        status_label.config(text="Processando arquivos...")

        try:
            if modo_selecionado == "Com Ano e Mês":
                ano = entry_ano.get()
                mes = entry_mes.get()

                if not ano or not mes:
                    messagebox.showerror("Erro", "Informe o ano e o mês.")
                    return

                if not mes.isdigit():
                    messagebox.showerror("Erro", "O mês deve ser numérico.")
                    return

                processar_pdfs(pasta_pdfs, caminho_excel, pasta_destino,
                               ano, int(mes), atualizar_barra_progresso, barra_progresso)
            else:
                processar_pdfs_informe(pasta_pdfs, caminho_excel, pasta_destino,
                                       atualizar_barra_progresso, barra_progresso)

            status_label.config(text="Processamento finalizado com sucesso!")
            messagebox.showinfo(
                "Sucesso", "Processamento concluído com sucesso!")
        except Exception as e:
            status_label.config(text="Erro durante o processamento.")
            messagebox.showerror("Erro", f"Ocorreu um erro: {str(e)}")

    def iniciar_processamento_thread():
        thread = threading.Thread(target=executar_processamento)
        thread.start()

    ttk.Label(scrollable_frame, text="Planilha Excel:").grid(
        row=0, column=0, padx=10, pady=5, sticky='w')
    ttk.Button(scrollable_frame, text="Selecionar", command=selecionar_excel).grid(
        row=0, column=1, padx=10, pady=5)
    caminho_excel_label = ttk.Label(scrollable_frame, text="", wraplength=400)
    caminho_excel_label.grid(row=1, column=0, columnspan=2, padx=10, pady=5)

    ttk.Label(scrollable_frame, text="Pasta de PDFs:").grid(
        row=2, column=0, padx=10, pady=5, sticky='w')
    ttk.Button(scrollable_frame, text="Selecionar", command=selecionar_pasta_pdfs).grid(
        row=2, column=1, padx=10, pady=5)
    pasta_pdfs_label = ttk.Label(scrollable_frame, text="", wraplength=400)
    pasta_pdfs_label.grid(row=3, column=0, columnspan=2, padx=10, pady=5)

    ttk.Label(scrollable_frame, text="Pasta de Destino:").grid(
        row=4, column=0, padx=10, pady=5, sticky='w')
    ttk.Button(scrollable_frame, text="Selecionar", command=selecionar_pasta_destino).grid(
        row=4, column=1, padx=10, pady=5)
    pasta_destino_label = ttk.Label(scrollable_frame, text="", wraplength=400)
    pasta_destino_label.grid(row=5, column=0, columnspan=2, padx=10, pady=5)

    ttk.Label(scrollable_frame, text="Modo de Processamento:").grid(
        row=6, column=0, padx=10, pady=5, sticky='w')
    modo_var = tk.StringVar(value="Com Ano e Mês")
    modo_menu = ttk.OptionMenu(
        scrollable_frame, modo_var, "Com Ano e Mês", "Com Ano e Mês", "Somente Código")
    modo_menu.grid(row=6, column=1, padx=10, pady=5)

    label_ano = ttk.Label(scrollable_frame, text="Ano (ex: 2024)")
    label_ano.grid(row=7, column=0, padx=10, pady=5, sticky='w')
    entry_ano = ttk.Entry(scrollable_frame)
    entry_ano.grid(row=7, column=1, padx=10, pady=5)

    label_mes = ttk.Label(scrollable_frame, text="Mês (1-12)")
    label_mes.grid(row=8, column=0, padx=10, pady=5, sticky='w')
    entry_mes = ttk.Entry(scrollable_frame)
    entry_mes.grid(row=8, column=1, padx=10, pady=5)

    modo_var.trace_add('write', atualizar_campos_visiveis)

    ttk.Button(scrollable_frame, text="Iniciar Processamento", command=iniciar_processamento_thread).grid(
        row=9, column=0, columnspan=2, padx=10, pady=15)

    barra_progresso = ttk.Progressbar(
        scrollable_frame, length=400, mode='determinate')
    barra_progresso.grid(row=10, column=0, columnspan=2, padx=10, pady=5)

    status_label = ttk.Label(scrollable_frame, text="")
    status_label.grid(row=11, column=0, columnspan=2, padx=10, pady=5)

    atualizar_campos_visiveis()
    root.mainloop()
