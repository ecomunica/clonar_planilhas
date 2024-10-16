# clonar_planilhas
---

# **How-To: Instalação e Utilização do Script para Geração de Planilhas com Interface Gráfica em ttkbootstrap**

### Índice
1. **Pré-requisitos**
2. **Instalação**
    - 2.1 Instalando o Python
    - 2.2 Instalando as bibliotecas necessárias
3. **Execução do Script**
4. **Como o Script Funciona**
5. **Dicas Adicionais**
6. **Resolução de Problemas**

---

## 1. **Pré-requisitos**

Para rodar o script, você precisará:
- **Python**: versão 3.7 ou superior.
- **Bibliotecas**: *ttkbootstrap*, *Pillow*, *pyexcel*, e *shutil*.
- Um arquivo de planilha de **modelo** no formato `.ods` (Modelo_a_ser_Clonado.ods).
- Uma planilha de **composição de nomes** no formato `.ods` (composicao.ods), que contém os dados para gerar novas planilhas.

---

## 2. **Instalação**

### 2.1. **Instalando o Python**

Se o Python não estiver instalado na sua máquina:
- **Windows**:
  1. Acesse [python.org/downloads](https://www.python.org/downloads/) e baixe a versão mais recente do Python.
  2. Durante a instalação, marque a opção "Add Python to PATH".
  
- **Linux** (Debian/Ubuntu):
  ```bash
  sudo apt update
  sudo apt install python3 python3-pip
  ```

### 2.2. **Instalando as bibliotecas necessárias**

Abra o terminal ou o prompt de comando e execute o seguinte comando para instalar as bibliotecas necessárias:

```bash
pip install ttkbootstrap Pillow pyexcel pyexcel-ods
```

Se você estiver no Linux e for necessário instalar o *tkinter* (usado pelo *ttkbootstrap*), também rode:

```bash
sudo apt-get install python3-tk
```

---

## 3. **Execução do Script**

### 3.1. **Baixe ou crie o script**.

Crie um arquivo Python, por exemplo `gerar_planilhas.py`, e cole o seguinte código:

```python
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

            nome_arquivo = f"{valorLinha}_{valorColuna}.ods"
            novo_arquivo = os.path.join(destino_arquivos.get(), nome_arquivo)
            
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
janela = ttk.Window(themename="superhero")  # Escolha de tema do ttkbootstrap
janela.title("Gerar Planilhas a partir de Modelo - Versão Avançada com ttkbootstrap")
janela.geometry('800x600')

# Variáveis para armazenar os caminhos
entrada_modelo = ttk.StringVar()
entrada_composicao = ttk.StringVar()
destino_arquivos = ttk.StringVar()

# Layout usando ttkbootstrap
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
```

### 3.2. **Executando o Script**:

No terminal ou prompt de comando, navegue até o diretório onde o script foi salvo e execute o comando:

```bash
python gerar_planilhas.py
```

---

## 4. **Como o Script Funciona**

1. **Seleção de Arquivos**: O usuário seleciona:
   - Um arquivo de modelo (`Modelo_a_ser_Clonado.ods`).
   - Um arquivo de composição de nomes (`composicao.ods`), que contém os dados para gerar as novas planilhas.

2. **Pré-visualização dos Dados**: A planilha de composição é carregada e exibida na interface para revisão.

3. **Seleção do Diretório de Destino**: O usuário escolhe onde salvar os novos arquivos gerados.

4. **Geração dos Arquivos**: Para cada linha na planilha de composição, o script copia o arquivo de modelo, renomeando-o com base nos valores das colunas da planilha.

5. **Barra de Progresso**: O progresso da criação dos arquivos é mostrado em tempo real.

---

## 5. **Dicas Adicionais**

- **Temas**: O `ttkbootstrap` permite mudar o tema da interface. Você pode mudar o tema substituindo o nome `superhero` no comando `ttk.Window

(themename="superhero")` por outro tema, como `flatly`, `darkly`, ou `journal`.

- **Personalização**: Se você quiser customizar o formato dos nomes dos arquivos gerados, edite a linha `nome_arquivo = f"{valorLinha}_{valorColuna}.ods"`.

---

## 6. **Resolução de Problemas**

### Problema: **ImportError: cannot import name 'ImageTk' from 'PIL'**
- Solução: Instale ou atualize o Pillow executando:
  ```bash
  pip install --upgrade Pillow
  ```

### Problema: **Tkinter não está disponível**
- Solução: Certifique-se de que o `tkinter` está instalado. No **Ubuntu/Debian**, execute:
  ```bash
  sudo apt-get install python3-tk
  ```

### Problema: **O progresso não avança ou o script trava**
- Solução: Verifique se as permissões de escrita estão corretas na pasta de destino e se os arquivos de modelo e composição estão corretos.

---

