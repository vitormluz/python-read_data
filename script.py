import pandas as pd
import openpyxl

# Importar base de dados
tabela_vendas = pd.read_excel('Vendas.xlsx')

# Visualizar base de dados
pd.set_option('display.max_columns', None)
print(tabela_vendas)

print('')
print('')
print('')
print('')

# Faturamento por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)

print('')
print('')
print('')
print('-' * 50)

# Quantidade de produtos por loja
qtd_produtos = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(qtd_produtos)

print('')
print('')
print('')
print('-' * 50)

# Ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / qtd_produtos['Quantidade']).to_frame()
print(ticket_medio)


print('')
print('')
print('')
print('-' * 50)

# Enviar um email com o relatório