import os
import re
import pandas as pd
from pdfminer.high_level import extract_text

def convert_pdf_to_txt(pdf_path, txt_path):
    text = extract_text(pdf_path)
    with open(txt_path, 'w', encoding='utf-8') as txt_file:
        txt_file.write(text)
    print(f"PDF convertido para TXT e salvo em: {txt_path}")

def extract_data_from_txt(file_path):
    extracted_data = []
    current_cpf = None
    current_nome = None
    current_data_nascimento = None
    current_convenio = None

    # Padrões de regex
    cpf_pattern = re.compile(r'\d{9}-\d{2}')
    nome_pattern = re.compile(r'\d{7}\s\d{6}\s+([A-Z\s]+)')  # Nome está após um número de matrícula seguido de outro número
    data_nascimento_pattern = re.compile(r'\d{2}/\d{2}/\d{4}')
    convenio_pattern = re.compile(r'(SITUACAO SERVIDOR:\s+([A-Z\s]+))')  # Capturar convênio

    rubrica_pattern = re.compile(r'AMORT.*CARTAO.*(CREDITO|BENEFICIO).*', re.IGNORECASE)

    with open(file_path, 'r', encoding='utf-8') as file:
        lines = file.readlines()

    for idx, line in enumerate(lines):
        # Verificar se a linha contém CPF e resetar os dados
        cpf_match = cpf_pattern.search(line)
        if cpf_match:
            # Ao encontrar um novo CPF, reiniciamos os dados anteriores
            current_cpf = cpf_match.group()
            current_nome = None
            current_data_nascimento = None
            current_convenio = None

        # Verificar se a linha contém o nome do servidor
        nome_match = nome_pattern.search(line)
        if nome_match:
            current_nome = nome_match.group(1).strip()

        # Verificar se a linha contém a data de nascimento
        data_nascimento_match = data_nascimento_pattern.search(line)
        if data_nascimento_match:
            current_data_nascimento = data_nascimento_match.group()

        # Verificar se a linha contém o convênio
        convenio_match = convenio_pattern.search(line)
        if convenio_match:
            current_convenio = convenio_match.group(2).strip()

        # Verificar se a linha contém a rubrica de interesse
        if rubrica_pattern.search(line) and current_cpf:
            # Extrair o nome completo da rubrica
            rubrica_nome_match = re.search(r'AMORT.*CARTAO.*(?:CREDITO|BENEFICIO).*', line, re.IGNORECASE)
            rubrica_nome = rubrica_nome_match.group() if rubrica_nome_match else ''

            # Extrair o valor (assumindo que é o último número na linha)
            valor_match = re.findall(r'[\d,.]+', line)
            valor = valor_match[-1] if valor_match else '0'

            # Armazenar os dados
            extracted_data.append({
                'CPF': current_cpf,
                'Nome': current_nome,
                'Data de Nascimento': current_data_nascimento,
                'Convenio': current_convenio,
                'Rubrica': rubrica_nome.strip(),
                'Valor': valor.replace(',', '.')
            })
            print(f"Rubrica encontrada: {rubrica_nome.strip()} - Valor: {valor.replace(',', '.')}")

    return extracted_data

def save_incremental_data(df, output_excel):
    # Se o arquivo Excel já existir, append os dados
    if os.path.exists(output_excel):
        existing_df = pd.read_excel(output_excel)
        df = pd.concat([existing_df, df], ignore_index=True)

    # Salvar os dados incrementais no Excel
    df.to_excel(output_excel, index=False)
    print(f"Dados salvos em {output_excel}")

def process_pdfs_in_folder(folder_path):
    output_excel = os.path.join(folder_path, 'dados_extracao_amortizacao_todos_pdfs.xlsx')

    # Percorrer todos os arquivos na pasta
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            txt_path = os.path.join(folder_path, f"{os.path.splitext(filename)[0]}.txt")

            # Converter o PDF em TXT
            convert_pdf_to_txt(pdf_path, txt_path)

            # Extrair os dados do arquivo TXT
            data = extract_data_from_txt(txt_path)

            # Verificar se existem dados extraídos
            if data:
                # Converter para DataFrame
                df = pd.DataFrame(data)

                # Certificar-se de que a coluna 'Rubrica' existe antes de aplicar as transformações
                if 'Rubrica' in df.columns:
                    # Remover números e vírgulas da coluna "Rubrica"
                    df['Rubrica'] = df['Rubrica'].apply(lambda x: re.sub(r'\d', '', x))
                    df['Rubrica'] = df['Rubrica'].apply(lambda x: re.sub(r',', '', x))

                # Salvar os dados de forma incremental no Excel
                save_incremental_data(df, output_excel)
            else:
                print(f"Nenhum dado extraído de {txt_path}")

# Caminho da pasta onde os PDFs estão localizados
folder_path = os.getcwd()  # Usando a pasta atual do projeto

# Processar todos os PDFs na pasta do projeto
process_pdfs_in_folder(folder_path)
