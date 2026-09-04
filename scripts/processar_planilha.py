import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

# Caminhos
ARQUIVO_ENTRADA = '../dados/vendas.xlsx'
ARQUIVO_SAIDA = '../dados/vendas_processadas.xlsx'

# 1. Ler a planilha com pandas
df = pd.read_excel(ARQUIVO_ENTRADA)

# 2. Calcular coluna Total
df['Total'] = df['Preco'] * df['Quantidade']

# 3. Salvar sem formatação (opcional)
df.to_excel(ARQUIVO_SAIDA, index=False)

# 4. Formatar com openpyxl
wb = load_workbook(ARQUIVO_SAIDA)
ws = wb.active

# Negrito no cabeçalho
for cell in ws[1]:
    cell.font = Font(bold=True)

# Destacar totais maiores que 10000
for row in ws.iter_rows(min_row=2, min_col=4, max_col=4):
    for cell in row:
        if cell.value and cell.value > 10000:
            cell.fill = PatternFill(start_color='FFFF00', end_color='FFFF00', fill_type='solid')

# Ajustar largura das colunas
for col in ws.columns:
    max_length = 0
    column = col[0].column_letter
    for cell in col:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2)
    ws.column_dimensions[column].width = adjusted_width

# Salvar arquivo final
wb.save(ARQUIVO_SAIDA)

print(f"Planilha processada e formatada salva em {ARQUIVO_SAIDA}")