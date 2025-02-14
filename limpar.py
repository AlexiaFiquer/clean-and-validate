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

def extrair_numeros(ddd, fone):
    ddd = limpar_numero(ddd)
    fone = limpar_numero(fone)
    numeros = re.findall(r'\d{8,9}', fone)

    if not numeros:
        return []
    
    numeros_formatados = []
    for numero in numeros:
        if len(ddd) == 2 and not numero.startswith(ddd):
            numero_completo = f"55{ddd}{numero}"
        else:
            numero_completo = f"55{numero}"

        if len(numero_completo) == 13 or len(numero_completo) == 12:
            numeros_formatados.append(numero_completo)
    
    return numeros_formatados

def processar_planilha(arquivo_entrada, arquivo_saida):
    caminho_entrada = f"dados/{arquivo_entrada}"
    df = pd.read_excel(caminho_entrada)

    linhas_expandidas = []
    telefones = []
    emails = []

    for _, row in df.iterrows():
        ddd = row.get('DDD', "")
        fone = row.get('Fone', "")
        email = validar_email(row.get('E-mail', ""))

        numeros = extrair_numeros(ddd, fone)
        
        for numero in numeros:
            linhas_expandidas.append({'Fone': numero, 'E-mail': email})
            telefones.append(numero)

        if not numeros and email:
            linhas_expandidas.append({'Fone': "", 'E-mail': email})
            emails.append(email)

    df_saida = pd.DataFrame(linhas_expandidas)
    df_saida = df_saida.drop_duplicates()

    caminho_saida = f"output2/{arquivo_saida}"
    df_saida.to_excel(caminho_saida, index=False)

    df_telefones = pd.DataFrame({'Fone': [t for t in telefones if t]})
    df_telefones = df_telefones.drop_duplicates()

    if not df_telefones.empty:
        df_telefones.to_excel("output3/telefone-carnaval-2024.xlsx", index=False)
    else:
        print("Nenhum telefone válido encontrado para salvar na planilha 'geral-e-televendas-telefones.xlsx'.")

    df_emails = pd.DataFrame({'E-mail': [e for e in emails if e]})
    df_emails = df_emails.drop_duplicates()
    df_emails.to_excel("output3/email-carnaval-2024.xlsx", index=False)

    print("Tarefa concluída: as planilhas foram processadas com sucesso.")

processar_planilha("lista carnaval 2024.xlsx", "geral-e-email-telefone.xlsx")
