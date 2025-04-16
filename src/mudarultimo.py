import pandas as pd

# Função para alterar o número antes do ".pdf"


def alterar_numero_final(codigo, novo_numero):
    # Localizar a posição do ".pdf"
    if ".pdf" in codigo:
        # Divide pelo prefixo antes do número final
        partes = codigo.rsplit("0000000", 1)
        if len(partes) == 2:
            return f"{partes[0]}0000000{novo_numero}.pdf"
    return codigo  # Retorna o código original se não seguir o padrão esperado


# Caminho do arquivo Excel
# Substitua pelo caminho completo
arquivo = r"C:\Users\Clayton\Desktop\projeto_holerites_final\funcionario_arquivo_grm.xlsx"
# Substitua pelos nomes das abas que deseja processar
abas = ["PORT", "GRM"]

# Carregar o Excel com múltiplas abas
# Lê todas as abas em um dicionário
planilhas = pd.read_excel(arquivo, sheet_name=None)

# Solicitar o novo número final ao usuário
novo_numero = input("Digite o novo número antes de '.pdf': ")

# Processar as abas selecionadas
for aba in abas:
    if aba in planilhas:
        df = planilhas[aba]
        # Atualizar a coluna "CODIGO" apenas se ela existir na aba
        if "CODIGO" in df.columns:
            df["CODIGO"] = df["CODIGO"].apply(
                lambda codigo: alterar_numero_final(codigo, novo_numero))
            planilhas[aba] = df
        else:
            print(f"A aba '{aba}' não contém a coluna 'CODIGO'.")
    else:
        print(f"A aba '{aba}' não foi encontrada na planilha.")

# Salvar o arquivo atualizado
arquivo_saida = r"C:\Users\Clayton\Desktop\projeto_holerites_final\tabela_nova22.xlsx"
with pd.ExcelWriter(arquivo_saida, engine="openpyxl") as writer:
    for aba, df in planilhas.items():
        df.to_excel(writer, index=False, sheet_name=aba)

print(f"Tabela atualizada salva como: {arquivo_saida}")
