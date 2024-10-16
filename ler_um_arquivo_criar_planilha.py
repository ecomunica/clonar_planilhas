import pyexcel as pe
import shutil

# Carrega a planilha
spreadsheet = pe.get_sheet(file_name="composicao.ods", name_columns_by_row=0)

# Loop sobre as linhas da planilha
for r in spreadsheet.row_range():
    valorLinha = spreadsheet.cell_value(r, 0)
    valorColuna = spreadsheet.cell_value(r, 1)
    
    # Verifica se linha ou coluna estão vazios
    if not valorLinha or not valorColuna:
        print('Linha ou coluna vazia, pulando...')
        continue
    
    # Gera o nome do arquivo de destino e copia o modelo
    novo_arquivo = f"{valorLinha}_{valorColuna}.ods"
    print(f'Copiando arquivo para: {novo_arquivo}')
    shutil.copy("Modelo_a_ser_Clonado.ods", novo_arquivo)
