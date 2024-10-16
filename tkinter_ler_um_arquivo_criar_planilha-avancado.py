import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
import pyexcel as pe
import shutil
import os

# Função para selecionar o arquivo de modelo
def selecionar_modelo():
    arquivo = filedialog.askopenfilename(filetypes=[("ODS files", "*.ods")])
    if arquivo:
        entrada_modelo.set(arquivo)

# Função para selecionar a planilha de nomes
def selecionar_composicao():
    arquivo = filedialog.askopenfilename(filetypes=[("ODS files", "*.ods")])
    if arquivo:
        entrada_composicao.set(arquivo)
        # Carrega e exibe os dados da planilha
        try:
            spreadsheet = pe.get_sheet(file_name=arquivo, name_columns_by_row=0)
            tabela.delete(*tabela.get_children())  # Limpa a tabela
            for r in spreadsheet.row_range():
                valores = [spreadsheet.cell_value(r, c) for c in range(spreadsheet.column_range()[-1] + 1)]
                tabela.insert('', 'end', values=valores)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

# Função para iniciar o processo de cópia
def iniciar_copia():
    modelo = entrada_modelo.get()
    composicao = entrada_composicao.get()
    if not modelo or not composicao:
        messagebox.showwarning("Aviso", "Selecione os arquivos de modelo e composição.")
        return
    
    try:
        spreadsheet = pe.get_sheet(file_name=composicao, name_columns_by_row=0)
        total_linhas = spreadsheet.number_of_rows()
        
        for r in spreadsheet.row_range():
            valorLinha = spreadsheet.cell_value(r, 0)
            valorColuna = spreadsheet.cell_value(r, 1)
            
            if not valorLinha or not valorColuna:
                continue

            # Nome customizável
            nome_arquivo = f"{valorLinha}_{valorColuna}.ods"
            novo_arquivo = os.path.join(destino_arquivos.get(), nome_arquivo)
            
            # Copia o modelo e atualiza a barra de progresso
            shutil.copy(modelo, novo_arquivo)
            progresso['value'] = (r / total_linhas) * 100
            janela.update_idletasks()
        
        messagebox.showinfo("Sucesso", "Processo concluído!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))

# Função para selecionar a pasta de destino
def selecionar_pasta_destino():
    pasta = filedialog.askdirectory()
    if pasta:
        destino_arquivos.set(pasta)

# Criação da interface
janela = ttk.Window(themename="cosmo")  # Escolha de tema do ttkbootstrap: "darkly", "flatly", "superhero", "cyborg", "journal".
janela.title("Gerar Planilhas a partir de Modelo - Versão Avançada com ttkbootstrap")
janela.geometry('800x600')

# Variáveis para armazenar os caminhos
entrada_modelo = ttk.StringVar()
entrada_composicao = ttk.StringVar()
destino_arquivos = ttk.StringVar()

# Layout usando ttkbootstrap para estilo moderno
ttk.Label(janela, text="Modelo a ser clonado", bootstyle="info").grid(row=0, column=0, padx=10, pady=10, sticky='w')
ttk.Entry(janela, textvariable=entrada_modelo, width=60, bootstyle="secondary").grid(row=0, column=1, padx=10, pady=10)
ttk.Button(janela, text="Selecionar", bootstyle="success", command=selecionar_modelo).grid(row=0, column=2, padx=10, pady=10)

ttk.Label(janela, text="Composição de nomes", bootstyle="info").grid(row=1, column=0, padx=10, pady=10, sticky='w')
ttk.Entry(janela, textvariable=entrada_composicao, width=60, bootstyle="secondary").grid(row=1, column=1, padx=10, pady=10)
ttk.Button(janela, text="Selecionar", bootstyle="success", command=selecionar_composicao).grid(row=1, column=2, padx=10, pady=10)

ttk.Label(janela, text="Pasta de destino", bootstyle="info").grid(row=2, column=0, padx=10, pady=10, sticky='w')
ttk.Entry(janela, textvariable=destino_arquivos, width=60, bootstyle="secondary").grid(row=2, column=1, padx=10, pady=10)
ttk.Button(janela, text="Selecionar", bootstyle="success", command=selecionar_pasta_destino).grid(row=2, column=2, padx=10, pady=10)

# Exibição dos dados da planilha de composição
ttk.Label(janela, text="Pré-visualização da planilha de composição:", bootstyle="info").grid(row=3, column=0, columnspan=3, padx=10, pady=10)

colunas = ['Linha', 'Coluna', 'Outro Dado']
tabela = ttk.Treeview(janela, columns=colunas, show='headings', bootstyle="light")
for col in colunas:
    tabela.heading(col, text=col)
tabela.grid(row=4, column=0, columnspan=3, padx=10, pady=10)

# Barra de progresso
progresso = ttk.Progressbar(janela, orient="horizontal", length=400, mode='determinate', bootstyle="info")
progresso.grid(row=5, column=0, columnspan=3, pady=20)

# Botão para iniciar o processo
ttk.Button(janela, text="Iniciar", bootstyle="primary", command=iniciar_copia).grid(row=6, column=1, pady=20)

janela.mainloop()
