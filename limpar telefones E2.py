import pandas as pd
import re

file_path = 'output2/todos-telefones.xlsx'

df = pd.read_excel(file_path)

def validar_telefone(telefone):
    padrao = r'^\+55\d{2}\d{8,9}$'
    if re.match(padrao, str(telefone)):
        return True
    return False

df['valido'] = df['Fone'].apply(validar_telefone)

df_validos = df[df['valido'] == True]

df_validos = df_validos.drop(columns=['valido'])

df_validos.to_excel(file_path, index=False)

print("Números inválidos removidos e resultados armazenados na planilha!")
