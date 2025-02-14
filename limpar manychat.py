import pandas as pd
import re

# Função para adicionar o código de país 55 ao telefone, caso não tenha
def adicionar_ddd_55(telefone):
    if isinstance(telefone, str):  # Garantir que o telefone seja uma string
        # Se o telefone não começar com +55 ou 55, adicionar 55
        if not telefone.startswith(('55', '+55')):
            telefone = '55' + telefone
    return telefone

input_file = 'dados/[ManyChat] Leads Aniversário de São Paulo 2025.xlsx'

# Ler o arquivo Excel
df = pd.read_excel(input_file)

# Exibir as primeiras linhas para ver como estão os dados
print("Primeiras linhas da planilha:")
print(df.head())

# Garantir que a coluna TELEFONE seja tratada como string
df['TELEFONE'] = df['TELEFONE'].astype(str)

# Remover qualquer ".0" no final do número
df['TELEFONE'] = df['TELEFONE'].apply(lambda x: x.rstrip('.0') if isinstance(x, str) else x)

# Adicionar o código DDD 55 nos números de telefone que não tiverem
df['TELEFONE'] = df['TELEFONE'].apply(adicionar_ddd_55)

# Exibir as primeiras linhas após o processamento para ver como ficaram os dados
print("Primeiras linhas após o processamento:")
print(df.head())

# Salvar o DataFrame limpo sobrescrevendo a planilha original
df.to_excel(input_file, index=False)

print(f'Arquivo limpo e sobrescrito em: {input_file}')
