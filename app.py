import re
import shutil
import tempfile
from pathlib import Path
from time import sleep

import pandas as pd
import pythoncom
import streamlit as st
import win32com.client as win32
from docx import Document

BASE = Path(__file__).parent
pasta_contratos = BASE / "contratos_gerados"
arquivo_zip = BASE / "contratosPDF.zip"
pasta_contratos.mkdir(exist_ok=True)

def word_para_pdf(word_caminho, pdf_caminho):
    pythoncom.CoInitialize()
    word_app = None
    try:
        word_app = win32.DispatchEx('Word.Application')
        word_app.Visible = False
        doc = word_app.Documents.Open(str(word_caminho))
        doc.SaveAs(str(pdf_caminho), FileFormat=17)  # 17 = PDF
        doc.Close(False)
        return True
    except Exception as e:
        st.error(f"Erro na conversão: {e}")
        return False
    finally:
        if word_app is not None:
            word_app.Quit()
        pythoncom.CoUninitialize()

def trocar_primeiro_paragrafo(doc, texto):
    p = doc.paragraphs[0]
    if p.runs:
        p.runs[0].text = texto #mantem formatacao do primeiro trecho
        for run in p.runs[1:]:
            run.text = ''
    else:
        p.text = texto

st.title('Contratos')
st.header('Contratos e banco de dados')
st.write('')
st.write('')

with st.container():
    st.write('Contrato padrão')
    word_base = st.file_uploader(label='📄Modelo de contrato.docx',key='upload_word')
    st.download_button(
        label='📄Modelo de contrato.docx',
        data=(BASE / 'modelo contrato.docx').read_bytes(),
        file_name='Modelo contrato.docx',
        mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        key='modelo contrato'
    )

st.write('')
st.write('')

with st.container():
    st.write('Banco de dados')
    csv = st.file_uploader(label='📒Banco de dados.csv',key='upload_csv')
    st.download_button(
        label='📒Banco de dados.csv',
        data=(BASE / 'dados_contrato.csv').read_bytes(),
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
            st.stop()
        if csv is None:
            st.error('Nenhum banco de dados foi carregado. Por gentileza, anexe um.')
            st.stop()

        word = Document(word_base)
        if len(word.paragraphs) == 0:
            st.error('Modelo de contrato vazio. Troque seu anexo.')
            st.stop()

        data = pd.read_csv(csv,delimiter=',')
        if 'nome' not in data.columns:
            st.error("O CSV precisa ter uma coluna chamada 'nome'.")
            st.stop()
        nomes_lista = data['nome'].dropna().tolist()

        for f in pasta_contratos.glob('*.pdf'):
            f.unlink()
        arquivo_zip.unlink(missing_ok=True)

        with tempfile.TemporaryDirectory() as tmp:
            for nome in nomes_lista:
                nome = str(nome).upper()
                nome_arquivo = re.sub(r'[\\/:*?"<>|]','-', nome)
                caminho_word = Path(tmp) / f'Contrato-{nome_arquivo}.docx'
                caminho_pdf = pasta_contratos / f'Contrato-{nome_arquivo}.pdf'

                word_base.seek(0)
                word = Document(word_base)
                trocar_primeiro_paragrafo(word, f'{nome},')
                word.save(caminho_word)
                word_para_pdf(caminho_word, caminho_pdf)

        shutil.make_archive(str(arquivo_zip.with_suffix('')), 'zip', pasta_contratos)
        st.success('Arquivo ZIP gerado com sucesso.')

with col2:
    if arquivo_zip.exists():
        st.download_button(
            label='Contratos PDF',
            data=arquivo_zip.read_bytes(),
            file_name='Contratos PDF.zip',
            mime='application/zip',
            key='arquivo zip'
        )

with col3:
    if st.button('Limpar arquivo ZIP'):
        if arquivo_zip.exists():
            arquivo_zip.unlink()
            st.sucess('Arquivo ZIP excluído!')
        else:
            st.error('Arquivo ZIP já foi excluído ou não existe.')
        sleep(2)
        st.rerun()