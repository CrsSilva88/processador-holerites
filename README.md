# processador-holerites

Aplicativo em Python com interface gráfica para automatizar o processamento e renomeação de holerites e informes de rendimentos em PDF, com suporte a extração de texto via OCR (Tesseract), leitura de planilhas Excel e relatórios detalhados. Ideal para processar grandes volumes de arquivos de forma rápida, segura e precisa.

## 🎯 Motivação

Na empresa onde atuo, enfrentávamos um problema recorrente: o recebimento de centenas de holerites em PDF a cada fechamento de mês, retirados do sistema de fechamneto de folha, que precisavam ser renomeados manualmente com base em códigos de cada funcionario de um outro sistema de gestão. Esse processo era repetitivo, sujeito a erros e tomava tempo da equipe.

Criei este projeto para automatizar totalmente essa tarefa. Agora, com um clique, conseguimos processar milhares de documentos com segurança, integrando OCR, leitura de planilhas e renomeação automatizada de arquivos, reduzindo o tempo de trabalho de horas para minutos.

---

## Funcionalidades

- Leitura de arquivos PDF com extração de CPF
- Leitura por OCR (Reconhecimento Óptico de Caracteres) para PDFs escaneados
- Integração com planilhas Excel (.xlsx) com duas abas distintas:
- **CODIGO** → usada no modo "Com ano e mês"
- **INFORME** → usada no modo "Somente código"
- Renomeia automaticamente os PDFs de acordo com os dados da planilha
- Interface gráfica desenvolvida com Tkinter
- Barra de progresso e status informativo
- Geração de relatórios no terminal (arquivos sem CPF, sem código, processados)

---

## Tecnologias utilizadas

- Python 3.11+
- Tkinter
- pandas
- pdfplumber
- pdf2image
- pytesseract
- openpyxl

---

## Instalação

### 1. Clone o repositório:

```
git clone https://github.com/CrsSilva88/processador-holerites.git

```

### 2. Instale as dependências:

```
pip install -r requirements.txt
```

### 4. Instale os pré-requisitos do sistema:

- **Tesseract OCR:** [Download aqui](https://github.com/UB-Mannheim/tesseract/wiki)
- **Poppler for Windows (necessário para pdf2image):** [Baixe aqui](https://github.com/oschwartz10612/poppler-windows/releases)

---

## Como usar

1. Execute o script principal:

```
python app.py
```

2. Na interface:
   - Selecione a planilha Excel (.xlsx)
   - Escolha a pasta com os arquivos PDF
   - Escolha a pasta de destino dos arquivos processados
   - Escolha o modo de processamento: **Com Ano e Mês** ou **Somente Código**
   - Clique em "Iniciar Processamento"

---

## Estrutura esperada da planilha Excel

- A planilha deve conter duas abas: `CODIGO` e `INFORME`
- Ambas devem conter uma coluna chamada `CODIGO`
- O código deve conter o CPF (com ou sem formatação)

---

## 📄 Licença

Este projeto está licenciado sob os termos da [MIT License](LICENSE).
