import os
import re
import json
import traceback
from dotenv import load_dotenv
from PyPDF2 import PdfReader
from langchain_openai import ChatOpenAI
from langchain_core.prompts import PromptTemplate
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor
import streamlit as st

def validate_json(dados, estrutura_padrao):
    """Valida e completa o JSON com estrutura padrão."""
    for chave in estrutura_padrao:
        if chave not in dados:
            dados[chave] = estrutura_padrao[chave]
    return dados

def create_docx_from_json(arquivo_json, arquivo_saida='curriculo.docx'):
    """Cria um documento Word formatado a partir de dados de um currículo em JSON."""
    try:
        # Carregar dados JSON
        with open(arquivo_json, 'r', encoding='utf-8') as f:
            dados = json.load(f)

        # Estrutura padrão para validação
        estrutura_padrao = {
            "informacoes_pessoais": {"nome": "", "cidade": "", "email": "", "telefone": "", "cargo": ""},
            "resumo_qualificacoes": [],
            "experiencia_profissional": [],
            "educacao": [],
            "certificacoes": []
        }
        dados = validate_json(dados, estrutura_padrao)

        # Criar um novo Documento
        doc = Document()

        # Definir fonte padrão
        estilo = doc.styles['Normal']
        estilo.font.name = 'Calibri'
        estilo.font.size = Pt(11)
        estilo.font.color.rgb = RGBColor(0, 0, 0)  # Define a cor preta

        # Função para adicionar espaço entre seções
        def adicionar_espaco():
            doc.add_paragraph().paragraph_format.space_after = Pt(12)

        # Nome (Cabeçalho Centralizado)
        informacoes_pessoais = dados.get('informacoes_pessoais', {})
        nome = informacoes_pessoais.get('nome', 'Nome Não Encontrado')
        paragrafo_nome = doc.add_paragraph(nome)
        if paragrafo_nome.runs:
            nome_run = paragrafo_nome.runs[0]
            nome_run.bold = True
            nome_run.font.size = Pt(16)
        else:
            print("Aviso: Nome vazio ou inválido no JSON.")
        paragrafo_nome.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

        # Informações de Contato (Alinhado à esquerda)
        adicionar_espaco()
        contato = f"""Cidade: {informacoes_pessoais.get('cidade', 'N/A')}
                    Bairro: {informacoes_pessoais.get('bairro', 'N/A')}
                    Email: {informacoes_pessoais.get('email', 'N/A')} 
                    Telefone: {informacoes_pessoais.get('telefone', 'N/A')}
                    Posição: {informacoes_pessoais.get('cargo', 'N/A')}"""
        paragrafo_contato = doc.add_paragraph(contato)
        paragrafo_contato.alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

        # Salvar o documento
        doc.save(arquivo_saida)
        print(f"Currículo salvo em {arquivo_saida}")
    except Exception as e:
        print(f"Erro ao criar documento Word: {e}")
        print(traceback.format_exc())


def process_text(texto):
    """Processa o texto e retorna JSON estruturado com tratamento de erros aprimorado."""
    # Carregar variáveis de ambiente do arquivo .env
    load_dotenv()
    chave_api = os.getenv('OPENAI_API_KEY')
    if not chave_api:
        raise ValueError("Chave da API OpenAI não encontrada. Certifique-se de que a variável está configurada corretamente.")

    # Prompt atualizado para maior clareza e contexto
    modelo_prompt = f"""
    Você é um especialista em extração de informações de currículos. Analise o texto abaixo e produza um JSON estruturado com:
    - Informações pessoais (nome, cidade, email, telefone, cargo desejado).
    - Resumo de qualificações.
    - Experiência profissional (empresa, cargo, período, atividades, projetos).
    - Formação acadêmica (instituição, grau, ano de conclusão).
    - Certificações.

    TEXTO DO CURRÍCULO:
    {texto}

    Formato esperado:
    {{
        "informacoes_pessoais": {{
            "nome": "",
            "cidade": "",
            "email": "",
            "telefone": "",
            "cargo": ""
        }},
        "resumo_qualificacoes": [],
        "experiencia_profissional": [],
        "educacao": [],
        "certificacoes": []
    }}
    """

    try:
        # Log do texto enviado
        print(f"Texto enviado para API (primeiros 500 caracteres): {texto[:500]}")

        # Comunicação com a API
        llm = ChatOpenAI(api_key=chave_api, temperature=0, model="gpt-4")
        resultado = llm.invoke(modelo_prompt)

        st.write(resultado.content)

        # Log da resposta da API
        print("Resposta da API OpenAI:", resultado.content)

        try:
            # Tenta converter a resposta em JSON
            dados_json = json.loads(resultado.content)
            return dados_json
        except json.JSONDecodeError as e:
            print(f"Erro ao decodificar JSON: {e}")
            print("Resposta recebida (não JSON):", resultado.content)
            # Retorna uma estrutura padrão caso o JSON seja inválido
            return {
                "informacoes_pessoais": {"nome": "", "cidade": "", "email": "", "telefone": "", "cargo": ""},
                "resumo_qualificacoes": [],
                "experiencia_profissional": [],
                "educacao": [],
                "certificacoes": []
            }

    except Exception as e:
        print(f"Erro ao processar texto com a API OpenAI: {e}")
        # Retorna uma estrutura padrão em caso de erro geral
        return {
            "informacoes_pessoais": {"nome": "", "cidade": "", "email": "", "telefone": "", "cargo": ""},
            "resumo_qualificacoes": [],
            "experiencia_profissional": [],
            "educacao": [],
            "certificacoes": []
        }



def extract_text_from_pdf(caminho_pdf):
    """Extrai o texto de um arquivo PDF."""
    try:
        leitor = PdfReader(caminho_pdf)
        texto = ""
        for pagina in leitor.pages:
            texto += pagina.extract_text() + "\n"
        texto_limpo = clear_text(texto)
        print(f"Texto extraído do PDF: {texto_limpo[:500]}...")  # Mostra os primeiros 500 caracteres
        return texto_limpo
    except Exception as e:
        print(f"Erro ao extrair texto do PDF: {e}")
        return ""

def clear_text(texto):
    """Limpa e normaliza o texto extraído."""
    texto = re.sub(r'\s+', ' ', texto)
    texto = re.sub(r'\n*Página \d+ de \d+\n*', '', texto)
    texto = re.sub(r'\n{3,}', '\n\n', texto)
    texto = texto.strip()
    return texto
