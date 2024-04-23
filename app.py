import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile
from io import BytesIO

planilha = pd.read_excel('Controle Ordens de Serviço ADM.xlsx', sheet_name='Ordens de Serviço - 2024', usecols='A:O')
planilha = planilha.copy()
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

lista_depto = sorted(planilha['DEPTO'].astype(str).unique().tolist())
lista_depto = ['GERAIS'] + lista_depto
selecao_depto = st.selectbox("Selecione um dos departamentos", lista_depto)
if st.button("Imprimir o gráfico"):
    if selecao_depto == 'GERAIS':
        dados = planilha
    else:
        dados = planilha[planilha['DEPTO'].isin(lista_depto)]

    fig, ax = plt.subplots(figsize=(14, 7))
    ax.barh(dados['TIPO DE SERVIÇO'].astype(str), dados.index)
    ax.set_xlabel('Índice')
    ax.set_ylabel('Tipo de Serviços')
    ax.set_title(f'Gráfico de "Tipo de Serviços" para o departamento {selecao_depto}')
    plt.subplots_adjust(left=0.3)
    buf = BytesIO()
    plt.savefig(buf, format='png')
    buf.seek(0)
    st.image(buf)
