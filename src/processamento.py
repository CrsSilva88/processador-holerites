import os
import shutil
import pandas as pd
from PyPDF2 import PdfReader
import re

def processar_pdfs(pasta_pdfs, arquivo_excel, pasta_destino, ano, mes, atualizar_barra_progresso, barra_progresso):
    # Carregar apenas a planilha "CODIGO"
    try:
        df_planilha = pd.read_excel(arquivo_excel, sheet_name="CODIGO")
    except Exception as e:
        raise Exception(f"Erro ao carregar a aba 'CODIGO': {str(e)}")

    # Iterar pelos arquivos PDF na pasta
    arquivos = [f for f in os.listdir(pasta_pdfs) if f.endswith(".pdf")]
    total_arquivos = len(arquivos)

    for index, pdf_file in enumerate(arquivos):
        pdf_path = os.path.join(pasta_pdfs, pdf_file)

        # Extrair texto do PDF
        reader = PdfReader(pdf_path)
        texto_pdf = "".join(page.extract_text() or "" for page in reader.pages)

        # Extrair o CPF do texto
        cpf_match = re.search(r"(\d{3}\.\d{3}\.\d{3}-\d{2})", texto_pdf)
        if cpf_match:
            cpf_extraido = cpf_match.group(1)
            cpf_limpo = re.sub(r"\D", "", cpf_extraido)

            for codigo in df_planilha["CODIGO"]:
                if cpf_limpo in codigo:
                    prefixo_ano_mes = f"{ano}{mes:02d}"
                    codigo_sem_prefixo = codigo[6:]  # Remove os 6 primeiros caracteres
                    novo_nome = f"{prefixo_ano_mes}{codigo_sem_prefixo}.pdf"
                    novo_caminho = os.path.join(pasta_destino, novo_nome)

                    shutil.move(pdf_path, novo_caminho)
                    break

        atualizar_barra_progresso(barra_progresso, int((index + 1) / total_arquivos * 100), total_arquivos)
