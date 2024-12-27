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

# Carregar variáveis de ambiente do arquivo .env
load_dotenv()

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
    chave_api = os.environ.get('OPENAI_API_KEY')
    
    if not chave_api:
        raise ValueError("Chave de API OpenAI não encontrada. Defina OPENAI_API_KEY no arquivo .env.")

    # Prompt ajustado para clareza e formato
    modelo_prompt = f"""
    Você é um especialista em extração de informações estruturadas de currículos.
    Analise o texto do currículo abaixo e gere um JSON estruturado com as seguintes informações:
    
    - Informações pessoais: nome, cidade, email, telefone, cargo.
    - Resumo de qualificações.
    - Experiência profissional (empresa, cargo, período, atividades, projetos).
    - Formação acadêmica (instituição, grau, ano de formatura).
    - Certificações.
    
    TEXTO DO CURRÍCULO:
    {texto}
    
    O JSON gerado deve seguir esta estrutura:
    {{
        "informacoes_pessoais": {{
            "nome": "Nome Completo",
            "cidade": "Cidade, Estado/País",
            "email": "email@exemplo.com",
            "telefone": "Número de telefone",
            "cargo": "Cargo Atual ou Desejado"
        }},
        "resumo_qualificacoes": ["Resumo das principais habilidades e conquistas"],
        "experiencia_profissional": [
            {{
                "empresa": "Nome da Empresa",
                "cargo": "Título do Cargo",
                "periodo": "Data de Início - Data de Término",
                "atividades": ["Descrição das atividades desempenhadas"],
                "projetos": ["Descrição dos projetos relevantes"]
            }}
        ],
        "educacao": [
            {{
                "instituicao": "Nome da Instituição",
                "grau": "Grau Obtido",
                "ano_formatura": "Ano de Formatura"
            }}
        ],
        "certificacoes": ["Nome das Certificações"]
    }}
    """

    try:
        print(f"Texto enviado à API:\n{texto[:500]}...\n")  # Log do texto enviado       
        llm = ChatOpenAI(api_key=chave_api, temperature=0, model="gpt-4")
        resultado = llm.invoke(modelo_prompt)
        
        print("Resposta da API OpenAI:")
        print(resultado.content)  # Exibe a resposta completa da API

        try:
            dados_json = json.loads(resultado.content)  # Tenta converter a resposta para JSON
        except json.JSONDecodeError as e:
            print(f"Erro ao decodificar JSON: {e}. Conteúdo da resposta:\n{resultado.content}")
            return {
                "informacoes_pessoais": {"nome": "", "cidade": "", "email": "", "telefone": "", "cargo": ""},
                "resumo_qualificacoes": [],
                "experiencia_profissional": [],
                "educacao": [],
                "certificacoes": []
            }

        # Verificação final e fallback para estrutura padrão
        estrutura_padrao = {
            "informacoes_pessoais": {"nome": "", "cidade": "", "email": "", "telefone": "", "cargo": ""},
            "resumo_qualificacoes": [],
            "experiencia_profissional": [],
            "educacao": [],
            "certificacoes": []
        }
        return validate_json(dados_json, estrutura_padrao)

    except Exception as e:
        print("Erro ao processar texto com a API OpenAI:", e)
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
