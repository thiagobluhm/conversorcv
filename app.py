import os
import streamlit as st
import json
import traceback
import base64
import tempfile
from pathlib import Path
from engine import extract_text_from_pdf, process_text, create_docx_from_json
from dotenv import load_dotenv
   
# Função para adicionar uma imagem de fundo a partir de um arquivo local
def add_bg_from_local(image_file):
    with Path(image_file).open("rb") as file:
        encoded_string = base64.b64encode(file.read()).decode()
    st.markdown(
        f"""
        <style>
        .stApp {{
            background-image: url(data:image/png;base64,{encoded_string});
            background-size: cover;
            background-position: center;
            background-repeat: no-repeat;
        }}
        </style>
        """,
        unsafe_allow_html=True
    )

# Função para exibir uma imagem de logo no topo a partir de um arquivo local
def add_logo_from_local(logo_file):
    with Path(logo_file).open("rb") as file:
        encoded_string = base64.b64encode(file.read()).decode()
    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] > .main {{
            padding-top: 0px;
        }}
        .logo-container {{
            display: flex;
            justify-content: center;
            align-items: center;
            padding: 1rem 0;
        }}
        .logo-container img {{
            max-width: 200px;
            height: auto;
        }}
        </style>
        <div class="logo-container">
            <img src="data:image/png;base64,{encoded_string}" alt="Logo">
        </div>
        """,
        unsafe_allow_html=True
    )

# Função principal do aplicativo
def main():
    st.set_page_config(page_title="Conversor de CV PDF para DOCX", page_icon="📄", layout="centered")

    # Aplicar imagens de fundo e logo
    add_bg_from_local("bg.png")
    add_logo_from_local("logo.png")

    load_dotenv()
    chave_api = os.environ.get('OPENAI_API_KEY')
        
    if not chave_api:
        st.write("Chave de API OpenAI não encontrada. Defina OPENAI_API_KEY no arquivo .env.")
    else:
        st.write("Conexao com a OPENAI ok!")
    
    st.markdown(
        "<h1 style='text-align: center; color: #4A9;'>Conversor de Currículo</h1>", 
        unsafe_allow_html=True
    )
    st.markdown(
        "<p style='text-align: center; font-size: 16px; color: #5A5A5A;'>"
        "O jeito mais fácil de formatar seu currículo em um documento Word!</p>",
        unsafe_allow_html=True
    )

    with st.form(key="upload_form"):
        uploaded_file = st.file_uploader(
            label="Envie seu currículo em PDF",
            type="pdf",
            help="Apenas arquivos em PDF são suportados. Verifique se o arquivo está legível e abaixo de 5MB."
        )
        submit_button = st.form_submit_button(label="Converter currículo")

    if submit_button and uploaded_file:
        st.markdown("---")
        progress_bar = st.progress(0)
        status_text = st.empty()

        with st.spinner("Convertendo o seu PDF..."):
            try:
                # Salvar PDF temporariamente usando tempfile
                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                    temp_pdf.write(uploaded_file.getvalue())
                    temp_pdf_path = temp_pdf.name
                
                # Etapa 1: Extração de texto
                status_text.text("Etapa 1: Extraindo texto do PDF...")
                progress_bar.progress(20)
                pdf_text = extract_text_from_pdf(temp_pdf_path)
                st.write("Texto extraído do PDF:", pdf_text)
                
                if not pdf_text.strip():
                    st.error("Não foi possível extrair texto do PDF. Verifique o arquivo e tente novamente.")
                    return
                
                # Etapa 2: Processamento do texto
                status_text.text("Etapa 2: Processando o texto do currículo...")
                progress_bar.progress(50)
                json_data = process_text(pdf_text)
                st.write("JSON gerado pela API:", json_data)

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
                st.markdown("### 📥 Baixe o currículo formatado")
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
                # Limpar arquivos temporários
                for temp_file in [temp_pdf_path, temp_json_path, temp_docx_path]:
                    if os.path.exists(temp_file):
                        os.remove(temp_file)

if __name__ == "__main__":
    main()
