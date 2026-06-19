import os
import streamlit as st
import streamlit.components.v1 as components
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

    st.markdown("<h1 style='text-align: center;'>Gerador de Parecer</h1>", unsafe_allow_html=True)

    with st.form(key="upload_form"):
        uploaded_file = st.file_uploader("Envie seu currículo em PDF", type="pdf")

        st.text_input(
            "Responsável",
            value=st.session_state.get('responsavel', ''),
            key='responsavel',
            placeholder="Nome do responsável pelo parecer"
        )
        st.text_input(
            "Disponibilidade",
            value=st.session_state.get('disponibilidade', ''),
            key='disponibilidade',
            placeholder="Disponibilidade para início das atividades"
        )
        st.text_input(
            "Modalidade",
            value=st.session_state.get('modalidade', ''),
            key='modalidade',
            placeholder="Modalidade de trabalho"
        )
        st.text_input(
            "Dados Pessoais",
            value=st.session_state.get('dados_pessoais', ''),
            key='dados_pessoais',
            placeholder="Idade, estado civil, residência..."
        )
        st.text_area(
            "Perfil Profissional",
            key="perfil_profissional",
            placeholder="Perfil profissional do candidato",
            value=st.session_state.get("perfil_profissional", ""),
            height=140,
            max_chars=None,
        )
        st.text_area(
            "Perfil Comportamental",
            key="perfil_comportamental",
            placeholder="Perfil comportamental do candidato",
            value=st.session_state.get("perfil_comportamental", ""),
            height=140,
            max_chars=None,
        )

        submit_button = st.form_submit_button("Gerar Parecer")

    campos_vazios = []

    if submit_button:
        campos_obrigatorios = {
            'responsavel': 'Responsável',
            'disponibilidade': 'Disponibilidade',
            'modalidade': 'Modalidade',
            'dados_pessoais': 'Dados Pessoais',
            'perfil_profissional': 'Perfil Profissional',
            'perfil_comportamental': 'Perfil Comportamental',
        }
        campos_vazios = [
            label for key, label in campos_obrigatorios.items()
            if not st.session_state.get(key, '').strip()
        ]

        alertas_html = ""

        estilo_btn = "float:right;background:none;border:none;font-size:20px;cursor:pointer;line-height:1;color:inherit;opacity:0.5;margin-left:8px;"

        if not uploaded_file:
            alertas_html += f"""
            <div class="alerta-item" style="background:#fff3cd;border-left:5px solid #f0a500;padding:14px 18px;border-radius:8px;margin:0 0 10px 0;font-size:15px;box-shadow:0 2px 8px rgba(0,0,0,0.12);">
                <button class="close-btn" style="{estilo_btn}">×</button>
                <strong>⚠️ Currículo não anexado</strong><br>
                Por favor, envie um currículo em PDF antes de gerar o parecer.
            </div>"""

        if campos_vazios:
            itens = "".join(f"<li>{c}</li>" for c in campos_vazios)
            alertas_html += f"""
            <div class="alerta-item" style="background:#fdecea;border-left:5px solid #e74c3c;padding:14px 18px;border-radius:8px;margin:0 0 10px 0;font-size:15px;box-shadow:0 2px 8px rgba(0,0,0,0.12);">
                <button class="close-btn" style="{estilo_btn}">×</button>
                <strong>🚫 Campos obrigatórios não preenchidos — reveja:</strong>
                <ul style="margin:6px 0 0 0;padding-left:20px;">{itens}</ul>
            </div>"""
            st.markdown("""
            <style>
            div[data-testid="stForm"] input:placeholder-shown,
            div[data-testid="stForm"] textarea:placeholder-shown {
                border: 2px solid #e74c3c !important;
                border-radius: 6px !important;
            }
            </style>
            """, unsafe_allow_html=True)

        alertas_html_js = alertas_html.replace('`', '\\`').replace('${', '\\${')

        if alertas_html:
            inject_alerts = f"""
            var el = parent.document.createElement('div');
            el.id = 'st-alert-overlay';
            el.style.cssText = 'position:fixed;bottom:24px;right:24px;z-index:9999;width:340px;max-width:90vw;';
            el.innerHTML = `{alertas_html_js}`;
            parent.document.body.appendChild(el);
            el.querySelectorAll('.close-btn').forEach(function(btn) {{
                btn.addEventListener('click', function() {{
                    var item = this.closest('.alerta-item');
                    if (item) item.remove();
                    if (el.querySelectorAll('.alerta-item').length === 0) el.remove();
                }});
            }});
            """
        else:
            inject_alerts = ""

        import time
        components.html(f"""
        <script>
            // ts:{int(time.time() * 1000)}
            var existing = parent.document.getElementById('st-alert-overlay');
            if (existing) existing.remove();
            {inject_alerts}
        </script>
        """, height=0)

    if submit_button and uploaded_file and not campos_vazios:
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
            progress_bar.progress(40)
            json_data = cvformatador.process_text_parecer(pdf_text)

            status_text.text("Processando dados adicionais...")
            progress_bar.progress(60)
            json_data.update({
                'disponibilidade': st.session_state.disponibilidade,
                'modalidade': st.session_state.modalidade,
                'dados_pessoais': st.session_state.dados_pessoais,
                'perfil_profissional': st.session_state.perfil_profissional,
                'perfil_comportamental': st.session_state.perfil_comportamental,
                'responsavel': st.session_state.responsavel
            })

            # print(f'ESTE É O JSON_DATA!!!!!!!{json_data}')

            if not json_data:
                st.error("Erro ao gerar JSON do currículo.")
                return

            with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode='w', encoding='utf-8') as temp_json:
                json.dump(json_data, temp_json, indent=2)
                temp_json_path = temp_json.name

            status_text.text("Etapa 3: Convertendo texto para formato PowerPoint...")
            progress_bar.progress(80)
            # with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_pptx:
            #     # cvformatador.create_parecer_pptx(temp_json_path, temp_docx.name)
            #     cvformatador.create_parecer_pptx(temp_json_path, temp_pptx.name)
            #     temp_docx_path = temp_pptx.name

            # status_text.text("Processo concluído")
            # progress_bar.progress(100)
            # st.success("Conversão concluída com sucesso! Baixe seu currículo abaixo.")
            # with open(temp_docx_path, "rb") as file:
            #     st.download_button(
            #         label="Baixar currículo em DOCX",
            #         data=file.read(),
            #         file_name='parecer.docx',
            #         mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            #     )
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pptx") as temp_pptx:
                # Gera o PPTX a partir do JSON
                cvformatador.create_parecer_pptx(
                    arquivo_json=temp_json_path,
                    arquivo_saida=temp_pptx.name,
                    responsavel=json_data['responsavel']
                )
                temp_pptx_path = temp_pptx.name

            status_text.text("Processo concluído")
            progress_bar.progress(100)
            st.success("Conversão concluída com sucesso! Baixe o parecer abaixo.")

            with open(temp_pptx_path, "rb") as file:
                st.download_button(
                    label="Baixar parecer em PPTX",
                    data=file.read(),
                    file_name="parecer.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
            st.error(traceback.format_exc())

if __name__ == "__main__":
    main()
