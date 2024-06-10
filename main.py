import streamlit as st
import xlrd
from openpyxl import Workbook
import pandas as pd


def convert_to_xlsx(file_path):
    # Ler a planilha XLS
    book = xlrd.open_workbook(file_path)
    sheet = book.sheet_by_index(0)

    # Criar um novo arquivo XLSX e copiar os dados
    new_book = Workbook()
    new_sheet = new_book.active

    for row_index in range(sheet.nrows):
        for col_index in range(sheet.ncols):
            new_sheet.cell(row=row_index + 1, column=col_index + 1).value = sheet.cell_value(row_index, col_index)

    # Salvar o arquivo XLSX
    new_file_path = file_path.replace('.xls', '.xlsx')
    new_book.save(new_file_path)

    return new_file_path


def process_excel(file_path):
    # Carregar a planilha Excel
    planilha = pd.read_excel(file_path)

    # Iterar sobre cada linha da planilha
    for index, row in planilha.iterrows():
        if pd.notna(row['DOCUMENTO']):  # Verificar se a coluna DOCUMENTO não está vazia
            if row['DOCUMENTO'] == 'Pix':  # Se o documento for 'Pix'
                # Verificar se a próxima linha está vazia
                if pd.isna(planilha.at[index + 2, 'VALOR']):
                    # Recortar o valor na linha atual e colar duas linhas abaixo
                    planilha.at[index + 2, 'VALOR'] = row['VALOR']
                    # Limpar o valor na linha atual
                    planilha.at[index, 'VALOR'] = None
            else:
                # Verificar se a próxima linha está vazia
                if pd.isna(planilha.at[index + 1, 'VALOR']):
                    # Recortar o valor na linha atual e colar uma linha abaixo
                    planilha.at[index + 1, 'VALOR'] = row['VALOR']
                    # Limpar o valor na linha atual
                    planilha.at[index, 'VALOR'] = None

    # Salvar a planilha com as alterações
    corrected_file_path = file_path.replace('.xlsx', '_corrigida.xlsx')
    planilha.to_excel(corrected_file_path, index=False)

    return corrected_file_path


# Interface do Streamlit
st.title('Conversão e Processamento de Planilhas Excel')

uploaded_file = st.file_uploader('Faça o upload de um arquivo XLS', type=['xls'])

if uploaded_file is not None:
    # Salvar o arquivo enviado
    xls_path = f"temp_{uploaded_file.name}"
    with open(xls_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())

    # Converter para XLSX
    xlsx_path = convert_to_xlsx(xls_path)
    st.write(f'Arquivo convertido: {xlsx_path}')

    # Processar a planilha
    corrected_file_path = process_excel(xlsx_path)
    st.write(f'Arquivo processado: {corrected_file_path}')

    # Permitir download do arquivo corrigido
    with open(corrected_file_path, 'rb') as f:
        st.download_button(
            label='Baixar planilha corrigida',
            data=f,
            file_name=corrected_file_path,
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )