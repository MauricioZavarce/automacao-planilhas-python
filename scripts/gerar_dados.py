import pandas as pd

dados = {
    'Produto': ['Notebook', 'Mouse', 'Teclado', 'Monitor', 'Webcam'],
    'Preco': [3500, 80, 150, 900, 250],
    'Quantidade': [10, 200, 150, 30, 100]
}

df = pd.DataFrame(dados)
df.to_excel('../dados/vendas.xlsx', index=False)
print("Planilha de exemplo criada em dados/vendas.xlsx")