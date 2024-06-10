import streamlit as st
import pandas as pd
import xlrd
from openpyxl import Workbook

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

def rearrange_values(planilha):
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
    
    return planilha

def main():
    st.title('Correção de Planilha')

    uploaded_file = st.file_uploader("Escolha uma planilha XLS ou XLSX", type=["xls", "xlsx"])

    if uploaded_file is not None:
        st.write("Planilha incorreta carregada!")
        st.write("Executando a correção...")

        try:
            # Salvar o arquivo carregado temporariamente
            file_path = f"temp_{uploaded_file.name}"
            with open(file_path, "wb") as f:
                f.write(uploaded_file.getbuffer())

            # Converter para XLSX, se necessário
            if file_path.endswith('.xls'):
                file_path = convert_to_xlsx(file_path)
            
            # Ler a planilha XLSX
            planilha = pd.read_excel(file_path)

            # Corrigir a planilha
            planilha_corrigida = rearrange_values(planilha.copy())

            # Salvar a planilha corrigida em um novo arquivo
            planilha_corrigida_file_path = "planilha_corrigida.xlsx"
            planilha_corrigida.to_excel(planilha_corrigida_file_path, index=False)

            st.write("Planilha corrigida pronta para download!")
            with open(planilha_corrigida_file_path, "rb") as file:
                st.download_button(
                    label="Baixar planilha corrigida",
                    data=file,
                    file_name=planilha_corrigida_file_path,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Ocorreu um erro durante a correção da planilha: {e}")

if __name__ == "__main__":
    main()
