import os
import streamlit as st
import json
import traceback
import tempfile
from dotenv import load_dotenv
import openai
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor
from pathlib import Path
import base64
import re
from cvformater import *

cvformatador = cvFormatter()

os.chdir(os.path.abspath(os.curdir))

def main():
    st.set_page_config(page_title="Conversor de CV PDF para DOCX **", page_icon="📄", layout="centered")

    #add_bg_from_local("bg.png")
    cvformatador.add_logo_from_local("Logo2.png")

    st.markdown("<h1 style='text-align: center;'>Conversor de Currículo</h1>", unsafe_allow_html=True)

    with st.form(key="upload_form"):
        uploaded_file = st.file_uploader("Envie seu currículo em PDF", type="pdf")
        submit_button = st.form_submit_button("Processar Currículo")

    if submit_button and uploaded_file:
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.getvalue())
                temp_pdf_path = temp_pdf.name

            status_text.text("Etapa 1: Extraindo texto do PDF...")
            progress_bar.progress(20)
            pdf_text = cvformatador.extract_text_from_pdf(temp_pdf_path)

            if not pdf_text.strip():
                st.error("Não foi possível extrair texto do PDF.")
                return

            #st.write("Texto extraído do PDF:", pdf_text)

            status_text.text("Etapa 2: Processando o texto do currículo...")
            progress_bar.progress(50)
            json_data = cvformatador.process_text_curriculo(pdf_text)

            # IMPRIMINDO NA TELA O TEXTO EXTRAIDO
            #st.write(json_data)

            # print(f'ESTE É O JSON_DATA{json_data}')

            if not json_data:
                st.error("Erro ao gerar JSON do currículo.")
                return

            st.session_state['cv_json_data'] = json_data
            st.session_state['cv_processado'] = True

            status_text.text("Processo concluído")
            progress_bar.progress(100)
            st.success("Currículo processado! Preencha as informações abaixo (opcional) e clique em Gerar Currículo.")

        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
            st.error(traceback.format_exc())

    if st.session_state.get('cv_processado'):

        @st.fragment
        def campo_perfil_profissional():
            # Aplica texto melhorado pendente ANTES de instanciar a caixa
            # (o Streamlit não permite alterar o session_state de uma caixa
            # depois que ela já foi desenhada na mesma execução).
            if 'pending_perfil_profissional' in st.session_state:
                st.session_state['perfil_profissional'] = st.session_state.pop('pending_perfil_profissional')

            st.text_area(
                "Perfil Profissional",
                key="perfil_profissional",
                placeholder="Perfil profissional do candidato (opcional)",
                height=140,
            )
            _, col_botao = st.columns([3, 1])
            with col_botao:
                if st.button("Melhorar Texto", key="melhorar_perfil_profissional", use_container_width=True):
                    texto_atual = st.session_state.get('perfil_profissional', '').strip()
                    if texto_atual:
                        with st.spinner("Melhorando texto..."):
                            st.session_state['pending_perfil_profissional'] = cvformatador.melhorar_texto(texto_atual)
                        st.rerun(scope="fragment")
                    else:
                        st.warning("Escreva algo no campo antes de melhorar o texto.")

        @st.fragment
        def campo_perfil_comportamental():
            if 'pending_perfil_comportamental' in st.session_state:
                st.session_state['perfil_comportamental'] = st.session_state.pop('pending_perfil_comportamental')

            st.text_area(
                "Perfil Comportamental",
                key="perfil_comportamental",
                placeholder="Perfil comportamental do candidato (opcional)",
                height=140,
            )
            _, col_botao = st.columns([3, 1])
            with col_botao:
                if st.button("Melhorar Texto", key="melhorar_perfil_comportamental", use_container_width=True):
                    texto_atual = st.session_state.get('perfil_comportamental', '').strip()
                    if texto_atual:
                        with st.spinner("Melhorando texto..."):
                            st.session_state['pending_perfil_comportamental'] = cvformatador.melhorar_texto(texto_atual)
                        st.rerun(scope="fragment")
                    else:
                        st.warning("Escreva algo no campo antes de melhorar o texto.")

        campo_perfil_profissional()
        campo_perfil_comportamental()

        gerar_button = st.button("Gerar Currículo")

        if gerar_button:
            try:
                json_data = st.session_state['cv_json_data']
                json_data['perfil_profissional'] = st.session_state.get('perfil_profissional', '').strip() or "Não foram acrescentadas informações"
                json_data['perfil_comportamental'] = st.session_state.get('perfil_comportamental', '').strip() or "Não foram acrescentadas informações"

                with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode='w', encoding='utf-8') as temp_json:
                    json.dump(json_data, temp_json, indent=2)
                    temp_json_path = temp_json.name

                with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
                    cvformatador.create_docx_curriculo(temp_json_path, temp_docx.name)
                    temp_docx_path = temp_docx.name

                st.success("Conversão concluída com sucesso! Baixe seu currículo abaixo.")
                with open(temp_docx_path, "rb") as file:
                    st.download_button(
                        label="Baixar currículo em DOCX",
                        data=file.read(),
                        file_name=f"Curriculo_{json_data['informacoes_pessoais']['nome']}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
                st.error(traceback.format_exc())

if __name__ == "__main__":
    main()
