# Ler dados de qualquer Planilha
# Automizar entrada de dados em planilhas
# Inserir dados de qualquer fonte(Word, Banco de dados, Outros Sistemas) --> Planinhas Excel
# (Bônus) Enviar Informações dados por E-mail, whatsapp ou telegram
from openpyxl import load_workbook

planilha_vendas = load_workbook('./dados/vendas_de_lanches.xlsx')
pagina_vendas = planilha_vendas('Sheet')

