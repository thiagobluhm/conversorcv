# Imagem base com Python
FROM python:3.12-slim

# Define diretório de trabalho
WORKDIR /app

# Copia os arquivos da aplicação
COPY . .

# Instala dependências
RUN pip install --upgrade pip && \
    pip install -r requirements.txt

# Expõe a porta que o Azure injetar
EXPOSE 8501

# Roda o Streamlit escutando na porta certa
CMD streamlit run Converter_Curriculo.py --server.port=8501 --server.address=0.0.0.0 --server.enableCORS false