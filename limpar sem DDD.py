import pandas as pd
import re

def limpar_numero(numero):
    if pd.isna(numero):
        return ""
    return re.sub(r'\D', '', str(numero)).lstrip("0")

def validar_email(email):
    if pd.isna(email):
        return ""
    email = email.strip().lower()
    if re.match(r'^\S+@\S+\.\S+$', email) and ".com" in email:
        return email
    else:
        return ""

def extrair_numeros(fone):  
    fone = limpar_numero(fone)
    numeros = re.findall(r'(?:55)?(\d{10,11})', fone)  

    if not numeros:
        return []
    
    numeros_formatados = []
    for numero in numeros:
        if len(numero) == 10:  
            numero = f"9{numero}"  
        numero_completo = f"55{numero}"
        numeros_formatados.append(numero_completo)
    
    return numeros_formatados

def processar_varias_planilhas(arquivos_entrada, arquivo_saida_telefones, arquivo_saida_emails):
    telefones = []
    emails = []
    
    for arquivo in arquivos_entrada:
        caminho_entrada = f"dados/{arquivo}"
        try:
            df = pd.read_excel(caminho_entrada, engine="openpyxl")
        except Exception as e:
            print(f"Erro ao abrir o arquivo {arquivo}: {e}")
            continue

        if 'ds_fone' not in df.columns or 'ds_email' not in df.columns:
            print(f"Erro: As colunas 'ds_fone' e 'ds_email' não foram encontradas no arquivo {arquivo}.")
            print(f"Colunas disponíveis: {df.columns.tolist()}")
            continue

        print(f"Processando {arquivo}...")
        
        for _, row in df.iterrows():
            fone = row.get('ds_fone', "")  # Linha para trocar (antes 'Fone', agora 'ds_fone')
            email = validar_email(row.get('ds_email', ""))  # Linha para trocar (antes 'E-mail', agora 'ds_email')

            numeros = extrair_numeros(fone)
            telefones.extend(numeros)
            if email:
                emails.append(email)
    
    df_telefones = pd.DataFrame({'Fone': list(set(telefones))})
    df_emails = pd.DataFrame({'E-mail': list(set(emails))})
    
    df_telefones.to_excel(f"output3/{arquivo_saida_telefones}", index=False)
    df_emails.to_excel(f"output3/{arquivo_saida_emails}", index=False)
    
    print("Tarefa concluída: as planilhas foram processadas e unificadas com sucesso.")

arquivos = ["lista carnaval 2024.xlsx"]
processar_varias_planilhas(arquivos, "carnaval-2024-telefones.xlsx", "carnaval-2024-emails.xlsx")
