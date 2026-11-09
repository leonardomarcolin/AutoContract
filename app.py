import os
import pandas as pd
import streamlit as st
from docx import Document
from time import sleep
import shutil
import win32com.client as win32
import pythoncom


pasta_contratos = r'C:\Users\discp\Desktop\Leo\Projetos - Python\AutoContract'
arquivo_zip = r'C:\Users\discp\Desktop\Leo\Projetos - Python\AutoContract\Contratos PDF.zip'

def word_para_pdf(word_caminho, pdf_caminho):
    try:
        pythoncom.CoInitialize()
        word_doc = win32.DispatchEx('Word.Application')
        doc = word_doc.Documents.Open(word_caminho)
        doc.SaveAs(pdf_caminho, FileFormat=17)  # 17 = PDF
        doc.Close()
        word_doc.Quit()
        return True
    except Exception as e:
        st.error(f"Erro na conversão: {e}")
        return False
    finally:
        pythoncom.CoUninitialize()

st.title('Contratos')
st.header('Contratos e banco de dados')
st.write('')
st.write('')
with st.container():
    st.write('Contrato padrão')
    word_base = st.file_uploader(label='📄Modelo de contrato.docx',key='upload_word')
    with open('modelo contrato.docx', 'rb') as file:
        file_bytes = file.read()
    st.download_button(
        label='📄Modelo de contrato.docx',
        data=file_bytes,
        file_name='Modelo de contrato.docx',
        mime='application / vnd.openxmlformats - officedocument.wordprocessingml.document',
        key='modelo contrato'
    )
st.write('')
st.write('')
with st.container():
    st.write('Banco de dados')
    csv = st.file_uploader(label='📒Banco de dados.csv',key='upload_csv')
    with open('dados_contrato.csv', 'rb') as file:
        file_bytes = file.read()
    st.download_button(
        label='📒Banco de dados.csv',
        data=file_bytes,
        file_name='Banco de dados.csv',
        mime='text/csv',
        key='banco de dados'
    )
st.write('')
st.write('')
col1, col2, col3 = st.columns(3)
with col1:
    if st.button('✒️Gerar contratos', help='gerar contratos'):
        if word_base is None:
            st.error('Nenhum modelo de contrato foi carregado. Por gentileza, anexe um.')
        if csv is None:
            st.error('Nenhum banco de dados foi carregado. Por gentileza, anexe um.')

        os.system('taskkill /f /im winword.exe')


        word = Document(word_base)
        data = pd.read_csv(csv,delimiter=',')
        nomes_lista = data['nome'].tolist()

        if os.path.exists(pasta_contratos):
            for item in os.listdir(pasta_contratos):
                os.remove(f'{pasta_contratos}/{item}')
        if os.path.exists(arquivo_zip):
            os.remove(arquivo_zip)

        if not os.path.exists(pasta_contratos):
                os.makedirs(pasta_contratos)

        if len(word.paragraphs) <= 0:
            st.error('Modelo de contrato vazio. Troque seu anexo.')

        for nomes in nomes_lista:
            nomes = nomes.upper()
            caminho_word = rf'C:\Users\discp\OneDrive - SENAC-SC\PycharmProjects\djabo\Contrato-{nomes}.docx'
            caminho_pdf = rf'C:\Users\discp\OneDrive - SENAC-SC\PycharmProjects\djabo\Contrato-{nomes}.pdf'
            word.paragraphs[0].text = f'{nomes},'
            word.save(caminho_word)

            if word_para_pdf(caminho_word,caminho_pdf):
                shutil.move(caminho_pdf, os.path.join(pasta_contratos, f'Contrato-{nomes}.pdf'))
            os.remove(caminho_word)

        shutil.make_archive('Contratos PDF', 'zip', pasta_contratos)

        st.success('Arquivo ZIP gerado com sucesso.')
with col2:
    if os.path.exists(arquivo_zip) is not None:
        try:
            with open(arquivo_zip, 'rb') as file:
                file_bytes = file.read()
            st.download_button(
                label='Contratos PDF',
                data=file_bytes,
                file_name='Contratos PDF.zip',
                mime='application/zip',
                key='arquivo zip'
            )
        except Exception as e:
            st.write(f'Nenhum arquivo gerado. Pressione o botão de gerar. {e}')
with col3:
    if st.button('Limpar arquivo ZIP'):
        if os.path.exists(arquivo_zip):
            os.remove(arquivo_zip)
            st.success('Arquivo ZIP excluído')
            sleep(2)
        else:
            st.error('Arquivo ZIP já foi excluído ou não existe.')
            sleep(2)
        st.rerun()