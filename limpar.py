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
    return email if re.match(r'^\S+@\S+\.\S+$', email) else ""

def extrair_numeros(ddd, fone):
    ddd = limpar_numero(ddd)
    fone = limpar_numero(fone)
    numeros = re.findall(r'\d{8,9}', fone)

    if not numeros:
        return []

    numeros_formatados = []
    for numero in numeros:
        if len(ddd) == 2 and not numero.startswith(ddd):
            numeros_formatados.append(f"55{ddd}{numero}")
        else:
            numeros_formatados.append(f"55{numero}")
    
    return numeros_formatados

def processar_planilha(arquivo_entrada, arquivo_saida):
    caminho = f"dados/{arquivo_entrada}"
    df = pd.read_excel(caminho)

    linhas_expandidas = []
    telefones = []
    emails = []

    for _, row in df.iterrows():
        ddd = row.get('DDD', "")
        fone = row.get('Fone', "")
        email = validar_email(row.get('Ds_email', ""))
        
        numeros = extrair_numeros(ddd, fone)
        for numero in numeros:
            linhas_expandidas.append({'Telefone': numero, 'Email': email})
            telefones.append(numero)

        if not numeros and email:
            linhas_expandidas.append({'Telefone': "", 'Email': email})
            emails.append(email)

    df_saida = pd.DataFrame(linhas_expandidas)
    caminho_saida = f"dados/{arquivo_saida}"
    df_saida.to_excel(caminho_saida, index=False)

    df_telefones = pd.DataFrame({'Telefone': [t for t in telefones if t]})
    df_telefones.to_excel("dados/todos-telefones.xlsx", index=False)

    df_emails = pd.DataFrame({'Email': [e for e in emails if e]})
    df_emails.to_excel("dados/todos-emails.xlsx", index=False)

    print("Tarefa concluída: as planilhas foram processadas com sucesso.")

processar_planilha("MAILING.XLSX", "todos-email-telefone.xlsx")
