import os
import streamlit as st
import json
import traceback
import tempfile
from engine import extract_text_from_pdf, process_text, create_docx_from_json


def main():
    st.set_page_config(page_title="Conversor de CV PDF para DOCX", page_icon="📄", layout="centered")

    st.markdown("<h1 style='text-align: center;'>Conversor de Currículo</h1>", unsafe_allow_html=True)

    # Upload de arquivo
    uploaded_file = st.file_uploader("Envie seu currículo em PDF", type="pdf")

    # Inicializar variáveis temporárias como None
    temp_pdf_path = None
    temp_json_path = None
    temp_docx_path = None

    if uploaded_file:
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            # Salvar PDF temporariamente no sistema de arquivos temporário
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.getvalue())
                temp_pdf_path = temp_pdf.name

            # Etapa 1: Extração de texto
            status_text.text("Etapa 1: Extraindo texto do PDF...")
            progress_bar.progress(20)
            pdf_text = extract_text_from_pdf(temp_pdf_path)

            if not pdf_text.strip():
                st.error("Não foi possível extrair texto do PDF.")
                return

            # Etapa 2: Processamento do texto
            status_text.text("Etapa 2: Processando o texto do currículo...")
            progress_bar.progress(50)
            try:
                json_data = process_text(pdf_text)
            except Exception as api_error:
                st.error(f"Erro ao processar o texto com a API OpenAI: {api_error}")
                return

            st.write("Texto extraído do PDF:", pdf_text)
            st.write("JSON gerado pela API:", json_data)

            if not json_data.get("informacoes_pessoais", {}).get("nome"):
                st.error("O JSON gerado está vazio ou incompleto.")
                return

            # Salvar JSON temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode='w', encoding='utf-8') as temp_json:
                json.dump(json_data, temp_json, indent=2)
                temp_json_path = temp_json.name

            # Etapa 3: Criando documento Word
            status_text.text("Etapa 3: Convertendo texto para formato Word...")
            progress_bar.progress(80)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
                create_docx_from_json(temp_json_path, temp_docx.name)
                temp_docx_path = temp_docx.name

            # Etapa 4: Conclusão
            progress_bar.progress(100)
            st.success("Conversão concluída com sucesso! Baixe seu currículo abaixo.")
            with open(temp_docx_path, "rb") as file:
                st.download_button(
                    label="Baixar currículo em DOCX",
                    data=file.read(),
                    file_name="Curriculo_Formatado.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
            st.error(traceback.format_exc())

        finally:
            # Limpar arquivos temporários, se existirem
            for temp_file in [temp_pdf_path, temp_json_path, temp_docx_path]:
                try:
                    if temp_file and os.path.exists(temp_file):
                        os.remove(temp_file)
                except Exception as cleanup_error:
                    print(f"Erro ao limpar arquivo temporário: {cleanup_error}")


if __name__ == "__main__":
    main()
