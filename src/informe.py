import os
import shutil
import pandas as pd
import pdfplumber
import re
from pdf2image import convert_from_path
import pytesseract

pytesseract.pytesseract.tesseract_cmd = r"C:\\Program Files\\Tesseract-OCR\\tesseract.exe"


def processar_pdfs_informe(pasta_pdfs, arquivo_excel, pasta_destino, atualizar_barra_progresso, barra_progresso):
    try:
        df_planilha = pd.read_excel(arquivo_excel, sheet_name="INFORME")
    except Exception as e:
        raise Exception(f"Erro ao carregar a aba 'INFORME': {str(e)}")

    arquivos = [f for f in os.listdir(pasta_pdfs) if f.endswith(".pdf")]
    total_arquivos = len(arquivos)

    codigos_dict = {}
    for codigo in df_planilha["CODIGO"]:
        cpf_somente_numeros = re.sub(r"\D", "", str(codigo))
        if len(cpf_somente_numeros) >= 11:
            cpf_chave = cpf_somente_numeros[-11:]
            codigos_dict[cpf_chave] = codigo

    arquivos_sem_cpf = []
    arquivos_sem_codigo = []
    arquivos_processados = []

    for index, pdf_file in enumerate(arquivos):
        print(f"Processando {index + 1}/{total_arquivos}: {pdf_file}")
        pdf_path = os.path.join(pasta_pdfs, pdf_file)

        with pdfplumber.open(pdf_path) as pdf:
            texto_pdf = "".join(page.extract_text()
                                or "" for page in pdf.pages)

        if not texto_pdf.strip():
            imagens = convert_from_path(pdf_path)
            texto_ocr = ""
            for img in imagens:
                texto_ocr += pytesseract.image_to_string(img, lang='por')
            texto_pdf = texto_ocr

        print(f"Texto extraído de {pdf_file}:")
        print(texto_pdf)
        print("-" * 40)

        cpf_match = re.search(r"(\d{3}\.\d{3}\.\d{3}-\d{2})", texto_pdf)
        if not cpf_match:
            arquivos_sem_cpf.append(pdf_file)
            continue

        cpf_extraido = cpf_match.group(1)
        cpf_limpo = re.sub(r"\D", "", cpf_extraido)

        if cpf_limpo in codigos_dict:
            codigo = codigos_dict[cpf_limpo]
            novo_nome = f"{codigo}.pdf"
            novo_caminho = os.path.join(pasta_destino, novo_nome)

            contador = 1
            while os.path.exists(novo_caminho):
                novo_nome = f"{codigo}_{contador}.pdf"
                novo_caminho = os.path.join(pasta_destino, novo_nome)
                contador += 1

            shutil.move(pdf_path, novo_caminho)
            arquivos_processados.append(pdf_file)
        else:
            arquivos_sem_codigo.append(pdf_file)

        atualizar_barra_progresso(barra_progresso, int(
            (index + 1) / total_arquivos * 100), total_arquivos)

    print("\nResumo do processamento:")
    print(f"Arquivos processados: {len(arquivos_processados)}")
    print(f"Sem CPF encontrado: {len(arquivos_sem_cpf)}")
    print(
        f"CPF encontrado mas sem código correspondente: {len(arquivos_sem_codigo)}")

    if arquivos_sem_cpf:
        print("\nArquivos sem CPF:")
        for f in arquivos_sem_cpf:
            print(" -", f)

    if arquivos_sem_codigo:
        print("\nArquivos com CPF não encontrado na planilha:")
        for f in arquivos_sem_codigo:
            print(" -", f)
