import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import tempfile

planilha = pd.read_excel('Controle Ordens de Serviço ADM.xlsx', sheet_name='Ordens de Serviço - 2024', usecols='A:O', parse_dates=['DATA E HORA SOLICITAÇÃO', 'DATA E HORA DA EXECUÇÃO'], dtype={'DATA E HORA SOLICITAÇÃO': str, 'DATA E HORA DA EXECUÇÃO': str})
planilha = planilha.copy()
st.write(planilha)

def download_planilha():
    temp_file_name = 'planilha.xlsx'
    planilha.to_excel(temp_file_name, index=False)
    with open(temp_file_name, 'rb') as f:
        bytes_data = f.read()
    st.download_button(
        label="Download da Planilha",
        data=bytes_data,
        file_name=temp_file_name,
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )

download_planilha()

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
