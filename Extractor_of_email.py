import PyPDF2
import re
from openpyxl import Workbook
import os

# Função para verificar se o arquivo está vazio
def is_file_empty(file_path):
    return os.path.getsize(file_path) == 0

# Função para extrair texto de um PDF
def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ""
        for page in reader.pages:
            text += page.extract_text()
    return text

# Função para extrair e-mails do texto usando regex
def extract_emails(text):
    email_pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    emails = re.findall(email_pattern, text)
    return emails

# Função para salvar e-mails em uma planilha Excel
def save_emails_to_excel(emails, output_path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"
    
    for index, email in enumerate(emails, start=1):
        ws.cell(row=index, column=1, value=email)
    
    wb.save(output_path)

# Função para processar múltiplos arquivos PDF em uma pasta
def process_pdfs_in_folder(pdf_folder, output_path):
    all_emails = []
    
    for root, _, files in os.walk(pdf_folder):
        for filename in files:
            if filename.endswith('.pdf'):
                pdf_path = os.path.join(root, filename)
                
                # Verifica se o arquivo está vazio
                if is_file_empty(pdf_path):
                    print(f"Arquivo vazio ignorado: {pdf_path}")
                    continue
                
                print(f"Processando {pdf_path}...")
                
                try:
                    text = extract_text_from_pdf(pdf_path)
                    emails = extract_emails(text)
                    all_emails.extend(emails)
                except PyPDF2.errors.PdfReadError as e:
                    print(f"Erro ao processar {pdf_path}: {e}")
    
    # Remover duplicados, se necessário
    all_emails = list(set(all_emails))
    
    save_emails_to_excel(all_emails, output_path)
    print(f"Extração concluída! {len(all_emails)} e-mails encontrados e salvos em '{output_path}'.")

# Caminho da pasta com os PDFs e do arquivo Excel de saída
pdf_folder = 'ori'
output_path = 'emails_extraidos.xlsx'

# Processar os PDFs na pasta
process_pdfs_in_folder(pdf_folder, output_path)
