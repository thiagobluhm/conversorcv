import os
import streamlit as st
import json
import traceback
import tempfile
from engine import extract_text_from_pdf, process_text, create_docx_from_json


def main():
    st.set_page_config(page_title="Conversor de CV PDF para DOCX", page_icon="📄", layout="centered")

    # Título e descrição
    st.markdown("<h1 style='text-align: center;'>Conversor de Currículo</h1>", unsafe_allow_html=True)

    chave_api = os.getenv("OPENAI_API_KEY")
    if chave_api:
        st.write(f"Chave da API carregada com sucesso. {chave_api[-4:]}")
    else:
        st.write("Chave da API não encontrada. Verifique as configurações do Streamlit Secrets.")

    # Formulário para upload do arquivo
    with st.form(key="upload_form"):
        uploaded_file = st.file_uploader("Envie seu currículo em PDF", type="pdf")
        submit_button = st.form_submit_button("Converter currículo")

    # Inicializar variáveis temporárias
    temp_pdf_path = temp_json_path = temp_docx_path = None

    if submit_button and uploaded_file:
        progress_bar = st.progress(0)
        status_text = st.empty()

        try:
            # Etapa 1: Salvar PDF temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_pdf.write(uploaded_file.getvalue())
                temp_pdf_path = temp_pdf.name

            status_text.text("Etapa 1: Extraindo texto do PDF...")
            progress_bar.progress(20)
            pdf_text = extract_text_from_pdf(temp_pdf_path)

            if not pdf_text.strip():
                st.error("Não foi possível extrair texto do PDF.")
                return

            st.write("Texto extraído do PDF:", pdf_text)

            # Etapa 2: Processar texto extraído
            status_text.text("Etapa 2: Processando o texto do currículo...")
            progress_bar.progress(50)
            try:
                json_data = process_text(pdf_text)
                st.write("JSON gerado pela API:", json_data)
            except Exception as e:
                st.error(f"Erro ao processar texto com a API OpenAI: {e}")
                return

            if not json_data.get("informacoes_pessoais", {}).get("nome"):
                st.error("O JSON gerado está vazio ou incompleto.")
                return

            # Etapa 3: Salvar JSON temporariamente
            with tempfile.NamedTemporaryFile(delete=False, suffix=".json", mode='w', encoding='utf-8') as temp_json:
                json.dump(json_data, temp_json, indent=2)
                temp_json_path = temp_json.name

            # Etapa 4: Criar documento Word
            status_text.text("Etapa 3: Convertendo texto para formato Word...")
            progress_bar.progress(80)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
                create_docx_from_json(temp_json_path, temp_docx.name)
                temp_docx_path = temp_docx.name

            # Conclusão
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
            # Remover arquivos temporários
            for temp_file in [temp_pdf_path, temp_json_path, temp_docx_path]:
                try:
                    if temp_file and os.path.exists(temp_file):
                        os.remove(temp_file)
                except Exception as cleanup_error:
                    print(f"Erro ao limpar arquivo temporário: {cleanup_error}")


if __name__ == "__main__":
    main()
