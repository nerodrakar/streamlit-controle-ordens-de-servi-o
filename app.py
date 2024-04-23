import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile

planilha = pd.read_excel('Controle Ordens de Serviço ADM.xlsx', sheet_name='Ordens de Serviço - 2024', usecols='A:O')
planilha = planilha.copy()
st.write(planilha)

temp_file_name = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
planilha.to_excel(temp_file_name, index=False)

st.download_button(
    label="Download da Planilha",
    data=temp_file_name,
    file_name='planilha.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)

lista_depto = sorted(planilha['DEPTO'].astype(str).unique().tolist())
lista_depto = ['GERAIS'] + lista_depto
selecao_depto = st.selectbox("Selecione um dos departamentos", lista_depto)
if st.button("Imprimir o gráfico"):
    if selecao_depto == 'GERAIS':
        dados = planilha
        texto = f'Gráfico de "Tipo de Serviços" para {selecao_depto}'
    else:
        dados = planilha[planilha['DEPTO'].isin(lista_depto)]
        texto = f'Gráfico de "Tipo de Serviços" para o departamento {selecao_depto}'

    temp_image = tempfile.NamedTemporaryFile(delete=False, suffix='.png').name
    fig, ax = plt.subplots(figsize=(14, 7))
    ax.barh(dados['TIPO DE SERVIÇO'].astype(str), dados.index)
    ax.set_xlabel('Índice')
    ax.set_ylabel('Tipo de Serviços')
    ax.set_title(texto)
    plt.subplots_adjust(left=0.3)
    plt.savefig(temp_image)
    st.image(temp_image)
