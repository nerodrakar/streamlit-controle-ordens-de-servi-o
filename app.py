import streamlit as st
import pandas as pd
import tempfile

df = pd.read_excel('Controle Ordens de Serviço ADM.xlsx', sheet_name='Ordens de Serviço - 2024', usecols='A:O')
planilha = df.copy()

st.write(planilha)

temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
temp_file.close()

planilha.to_excel(temp_file.name, index=False)

with open(temp_file.name, 'rb') as f:
    data = f.read()

st.download_button(
    label="Download da Planilha",
    data=data,
    file_name='planilha.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)


selecao_coluns = st.selectbox("Selecione uma das colunas", planilha.columns)
if st.button("Imprimir a coluna"):
    st.write(df[selecao_coluns])
