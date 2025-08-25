import json
import traceback
from dotenv import load_dotenv
import openai
from PyPDF2 import PdfReader
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import RGBColor as docxRGBColor
from pathlib import Path
import base64
import re
import os 
import streamlit as st
os.chdir(os.path.abspath(os.curdir))
from docx.shared import Inches, Cm
import datetime
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor as pptxRGBColor
from datetime import date
import json, re, unicodedata


class cvFormatter():
    def __init__(self):
        pass
    
    def validate_json(self, dados, estrutura_padrao):
        """Valida e completa o JSON com estrutura padrão."""
        for chave in estrutura_padrao:
            if chave not in dados:
                dados[chave] = estrutura_padrao[chave]
        return dados

    def create_docx_curriculo(self, arquivo_json, arquivo_saida='curriculo.docx', logo_path='Logo2.png'):
        """Cria um documento Word formatado a partir de dados de um currículo em JSON e adiciona um logo."""
        try:
            with open(arquivo_json, 'r', encoding='utf-8') as f:
                dados = json.load(f)

            estrutura_padrao = {
                "informacoes_pessoais": {"nome": "", "cidade": "", "email": "", "telefone": "", "cargo": ""},
                "resumo_qualificacoes": [],
                "experiencia_profissional": [],
                "educacao": [],
                "certificacoes": []
            }
            dados = self.validate_json(dados, estrutura_padrao)

            doc = Document()
            estilo = doc.styles['Normal']
            estilo.font.name = 'Calibri'
            estilo.font.size = Pt(11)
            estilo.font.color.rgb = docxRGBColor(0, 0, 0)

            def adicionar_espaco():
                """Adiciona um parágrafo vazio para espaçamento."""
                doc.add_paragraph().paragraph_format.space_after = Pt(12)

            if logo_path:
                section = doc.sections[0]
                section.header_distance = Cm(0.6)

                header = section.header
                header_paragraph = header.paragraphs[0]
                run = header_paragraph.add_run()
                run.add_picture(logo_path, width=Inches(0.8))  # Ajusta o tamanho do logo
                header_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT  # Alinha à direita

            # Informações pessoais
            informacoes_pessoais = dados.get('informacoes_pessoais', {})
            nome = informacoes_pessoais.get('nome', 'Nome Não Encontrado')
            paragrafo_nome = doc.add_paragraph(nome)
            if paragrafo_nome.runs:
                nome_run = paragrafo_nome.runs[0]
                nome_run.bold = True
                nome_run.font.size = Pt(16)
            paragrafo_nome.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            adicionar_espaco()
            contato = f"Cidade: {informacoes_pessoais.get('cidade', 'N/A')}\nEmail: {informacoes_pessoais.get('email', 'N/A')}\nTelefone: {informacoes_pessoais.get('telefone', 'N/A')}\nPosição: {informacoes_pessoais.get('cargo', 'N/A')}"
            doc.add_paragraph(contato)

            adicionar_espaco()

            # Resumo de qualificações
            doc.add_heading('Resumo de Qualificações', level=2)
            for qualificacao in dados.get('resumo_qualificacoes', []):
                doc.add_paragraph(f"- {qualificacao}")

            adicionar_espaco()

            # Experiência profissional
            doc.add_heading('Experiência Profissional', level=2)
            for experiencia in dados.get('experiencia_profissional', []):
                empresa = experiencia.get('empresa', 'Empresa Não Informada')
                cargo = experiencia.get('cargo', 'Cargo Não Informado')
                periodo = experiencia.get('periodo', 'Período Não Informado')
                local = experiencia.get('local', 'Local Não Informado')
                atividades = experiencia.get('atividades_exercidas', [])

                doc.add_paragraph(f"{empresa} ({local})", style='Heading 3')
                doc.add_paragraph(f"{cargo} - {periodo}", style='Normal')
                
                if atividades:
                    doc.add_paragraph("Atividades exercidas:", style='Normal')
                    for atividade in atividades:
                        doc.add_paragraph(f"{atividade}", style='List Bullet')

                ferramentas = experiencia.get('ferramentas', [])
                if ferramentas:
                    doc.add_paragraph("Ferramentas utilizadas:", style='Normal')
                    for ferramenta in ferramentas:
                        doc.add_paragraph(f"{ferramenta}", style='List Bullet')

            adicionar_espaco()

            # Educação
            doc.add_heading('Educação', level=2)
            for educacao in dados.get('educacao', []):
                instituicao = educacao.get('instituicao', 'Instituição Não Informada')
                curso = educacao.get('curso', 'Curso Não Informado')
                periodo = educacao.get('periodo', 'Período Não Informado')

                doc.add_paragraph(f"{instituicao}", style='Heading 3')
                doc.add_paragraph(f"{curso} - {periodo}", style='Normal')

            adicionar_espaco()

            # Certificações
            doc.add_heading('Certificações', level=2)
            for certificacao in dados.get('certificacoes', []):
                doc.add_paragraph(f"- {certificacao}", style='Normal')

            # Salvar o documento Word
            doc.save(arquivo_saida)
            print(f"Currículo salvo em {arquivo_saida}")

        except Exception as e:
            print(f"Erro ao criar documento Word: {e}")
            print(traceback.format_exc())


    # def create_docx_parecer(self, arquivo_json, arquivo_saida: str = "parecer.docx", responsavel: str = "Responsável"):

    #     try:
    #         with open(arquivo_json, 'r', encoding='utf-8') as f:
    #             dados = json.load(f)

    #         """
    #         Cria um parecer de candidato em DOCX com estrutura básica:
    #         - Cabeçalho: título fixo + responsável + data
    #         - Nome do candidato (se existir)
    #         - Blocos: Formação, Perfil Profissional, Perfil Comportamental (em branco)

    #         O JSON pode conter:
    #         {
    #             "nome": "Nome do Candidato",
    #             "formacao": [...],
    #             "perfil_profissional": [...],
    #             "perfil_comportamental": "...",  # opcional
    #             ...
    #         }
    #         """

    #         estrutura_padrao = {
    #             "nome": '',
    #             "formacao": [...],
    #             "competencias": [...], 
    #             "perfil_profissional": [...],
    #             "perfil_comportamental": "..."
    #         }

    #         dados = self.validate_json(dados, estrutura_padrao)

    #         doc = Document()

    #         # ---------- estilo base ----------
    #         estilo = doc.styles["Normal"]
    #         estilo.font.name = "Calibri"
    #         estilo.font.size = Pt(11)
    #         estilo.font.color.rgb = RGBColor(0, 0, 0)

    #         # ---------- título ----------
    #         p_titulo = doc.add_paragraph("PARECER DE CANDIDATO")
    #         p_titulo.runs[0].bold = True
    #         p_titulo.runs[0].font.size = Pt(14)
    #         p_titulo.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    #         # ---------- responsável + data ----------
    #         hoje = datetime.date.today().strftime("%d/%m/%Y")
    #         info = doc.add_paragraph(f"Responsável: {responsavel} | Data: {hoje}")
    #         info.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

    #         doc.add_paragraph()  # linha em branco

    #         # ---------- nome do candidato ----------
    #         if dados.get("nome"):
    #             p_nome = doc.add_paragraph(dados.get("nome", ""))
    #             p_nome.runs[0].bold = True
    #             p_nome.runs[0].font.size = Pt(12)
    #             p_nome.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    #             doc.add_paragraph()

    #         # ---------- formação ----------
    #         doc.add_heading("Formação", level=2)
    #         for f in dados.get("formacao", []):
    #             linha = f"{f.get('grau', '')} em {f.get('curso', '')} - {f.get('instituicao', '')} ({f.get('conclusao', '')})"
    #             doc.add_paragraph(linha)
    #         if not dados.get("formacao"):
    #             doc.add_paragraph("N/A")

    #         doc.add_paragraph()

    #         # ---------- competências ----------
    #         doc.add_heading("Competências", level=2)
    #         for c in dados.get("competencias", []):
    #             doc.add_paragraph(c)
    #         if not dados.get("competencias"):
    #             doc.add_paragraph("N/A")

    #         doc.add_paragraph()

    #         # ---------- perfil profissional ----------
    #         doc.add_heading("Perfil Profissional", level=2)
    #         for item in dados.get("perfil_profissional", []):
    #             # doc.add_paragraph(item, style="List Bullet")
    #             doc.add_paragraph(item)
    #         if not dados.get("perfil_profissional"):
    #             doc.add_paragraph("N/A")

    #         doc.add_paragraph()

    #         # ---------- perfil comportamental ----------
    #         doc.add_heading("Perfil Comportamental", level=2)
    #         perfil_comp = dados.get("perfil_comportamental", "")
    #         doc.add_paragraph(perfil_comp or "—")

    #         doc.add_paragraph()

    #         # ---------- salvar ----------
    #         try:
    #             doc.save(arquivo_saida)
    #             print(f"Parecer salvo em: {arquivo_saida}")
    #         except Exception as e:
    #             print(f"Erro ao salvar DOCX: {e}")
    #             print(traceback.format_exc())

    #     except Exception as e:
    #         print(f"Erro ao criar documento Word: {e}")
    #         print(traceback.format_exc())


    # def create_parecer_pptx(
    #     self,
    #     arquivo_json: str,                 # ← 1º arg = JSON (caminho)
    #     arquivo_saida: str,                # ← 2º arg = PPTX de saída
    #     responsavel: str = "Responsável",
    #     template_path: str | None = None,  # ← template opcional (tem default)
    #     title_hex: str = "#2E578C",
    #     font_size: int = 12,
    # ):
    #     """
    #     Converte o JSON do parecer em um PPTX usando o template.
    #     - Títulos em azul (#2E578C) e bold.
    #     - Corpo em 12pt.
    #     """
    #     # 0) Template default ao lado deste arquivo
    #     if template_path is None:
    #         template_path = str(Path(__file__).with_name("PARECER - Perfil Conforme e Estável.pptx"))
    #     if not Path(template_path).exists():
    #         raise FileNotFoundError(f"Template PPTX não encontrado em: {template_path}")

    #     # 1) Carrega o JSON
    #     with open(arquivo_json, "r", encoding="utf-8") as f:
    #         dados = json.load(f)

    #     # 2) Normaliza chaves de topo
    #     d = {(k.lower() if isinstance(k, str) else k): v for k, v in (dados or {}).items()}

    #     # 3) Helpers locais
    #     def _hex_to_rgbcolor(h: str) -> pptxRGBColor:
    #         h = h.strip().lstrip("#")
    #         return pptxRGBColor(int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16))

    #     def _find(substrings):
    #         if isinstance(substrings, str):
    #             subs = [substrings.lower()]
    #         else:
    #             subs = [s.lower() for s in substrings]
    #         for shp in slide.shapes:
    #             if getattr(shp, "has_text_frame", False):
    #                 low = (shp.text or "").lower()
    #                 if all(s in low for s in subs):
    #                     return shp
    #         return None

    #     def _write_simple(shape, text):
    #         if not shape or not hasattr(shape, "text_frame"):
    #             return
    #         tf = shape.text_frame
    #         tf.clear()
    #         p = tf.paragraphs[0]
    #         r = p.add_run()
    #         r.text = "" if text is None else str(text)
    #         r.font.size = Pt(font_size)

    #     def _write_title_body(shape, title, body=None, title_color=None):
    #         if not shape or not hasattr(shape, "text_frame"):
    #             return
    #         tf = shape.text_frame
    #         tf.clear()
    #         # Título
    #         p0 = tf.paragraphs[0]
    #         r0 = p0.add_run()
    #         r0.text = title
    #         r0.font.bold = True
    #         r0.font.size = Pt(font_size)
    #         if title_color:
    #             r0.font.color.rgb = title_color
    #         # Corpo
    #         if body:
    #             if isinstance(body, str):
    #                 p = tf.add_paragraph()
    #                 p.text = body
    #                 for run in p.runs:
    #                     run.font.size = Pt(font_size)
    #             else:
    #                 for line in body:
    #                     p = tf.add_paragraph()
    #                     p.text = str(line)
    #                     for run in p.runs:
    #                         run.font.size = Pt(font_size)

    #     azul = _hex_to_rgbcolor(title_hex)

    #     # 4) Carrega o template e preenche
    #     prs = Presentation(template_path)
    #     slide = prs.slides[0]

    #     # Cabeçalho
    #     shp = _find("parecer de candidato")
    #     if shp: _write_title_body(shp, "PARECER DE CANDIDATO", None, title_color=azul)

    #     shp = _find(["responsável", "data do parecer"])
    #     if shp: _write_simple(shp, f"Responsável: {responsavel} |  Data do parecer: {date.today():%d/%m/%Y}")

    #     # Nome
    #     shp = _find("nome")
    #     if shp: _write_simple(shp, d.get("nome") or d.get("nome".lower()) or "")

    #     # Competências
    #     shp = _find(["competências", "técnicas"])
    #     if shp:
    #         comps = d.get("competencias") or []
    #         _write_simple(shp, " | ".join(comps) if comps else "Competências | Técnicas |")

    #     # Formação
    #     shp = _find("formação")
    #     if shp:
    #         linhas = []
    #         for f in (d.get("formacao") or []):
    #             linhas.append(f"- {f.get('grau','')} em {f.get('curso','')} — {f.get('instituicao','')} ({f.get('conclusao','')})")
    #         _write_title_body(shp, "FORMAÇÃO:", linhas or None, title_color=azul)

    #     # Coluna esquerda
    #     shp = _find("disponibilidade")
    #     if shp: _write_title_body(shp, "DISPONIBILIDADE:", d.get("disponibilidade") or None, title_color=azul)

    #     shp = _find("modalidade")
    #     if shp: _write_title_body(shp, "MODALIDADE:", d.get("modalidade") or None, title_color=azul)

    #     shp = _find("dados pessoais")
    #     if shp: _write_title_body(shp, "DADOS PESSOAIS:", d.get("dados_pessoais") or None, title_color=azul)

    #     # Perfil profissional
    #     shp = _find("perfil profissional")
    #     if shp:
    #         corpo = "\n\n".join(d.get("perfil_profissional") or []) or None
    #         _write_title_body(shp, "PERFIL PROFISSIONAL:", corpo, title_color=azul)

    #     # Perfil comportamental
    #     shp = _find("perfil comportamental")
    #     if shp:
    #         pc = d.get("perfil_comportamental")
    #         _write_title_body(shp, "PERFIL COMPORTAMENTAL:", (pc if pc else None), title_color=azul)

    #     # 5) Salva
    #     Path(arquivo_saida).parent.mkdir(parents=True, exist_ok=True)
    #     prs.save(arquivo_saida)
    #     return arquivo_saida






    def create_parecer_pptx(
        self,
        arquivo_json: str,
        arquivo_saida: str,
        template_path: str | None = None,
        responsavel: str | None = None,
        
    ):
        """
        Preenche um PPTX a partir de um dicionário/JSON, preservando a formatação do template.
        Estratégia:
        1) Se existir shape com NOME igual à chave (ex.: 'nome', 'modalidade'), escreve nela.
        2) Caso contrário, substitui placeholder {{chave}} no texto de shapes/células.
        Não altera estilos (tamanhos, bullets, espaçamentos): tudo vem do template.
        """

        # ---------- utilitários ----------
        import re, json, unicodedata
        from pathlib import Path
        from datetime import date
        from pptx import Presentation

        def _normalize(s: str) -> str:
            if s is None: return ""
            s = unicodedata.normalize("NFD", s)
            s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")  # remove acentos
            s = re.sub(r"\s+", " ", s.strip().lower())
            return s

        def _iter_shapes(container):
            # percorre shapes e subshapes (grupos)
            for shp in getattr(container, "shapes", []):
                yield shp
                if hasattr(shp, "shapes"):   # grupo
                    for s in _iter_shapes(shp):
                        yield s

        def _clone_paragraph_style(src_paragraph, dst_paragraph):
            """
            Clona o pPr (Paragraph Properties) do parágrafo fonte para o destino,
            garantindo espaçamento/bullets/entrelinhas idênticos.
            """
            try:
                from copy import deepcopy
                src_pPr = getattr(src_paragraph._p, "pPr", None)
                if src_pPr is not None:
                    # garante que o destino tenha pPr e substitui pelo do fonte
                    dst_pPr = dst_paragraph._p.get_or_add_pPr()
                    dst_paragraph._p.remove(dst_pPr)
                    dst_paragraph._p.append(deepcopy(src_pPr))
            except Exception:
                # se a API/estruturas variarem, seguimos com o padrão herdado
                pass

        def _set_text_preservando_estilo(tf, text: str):
            """
            Preserva exatamente o estilo do template:
            - Reutiliza o 1º run do 1º parágrafo (mantém rPr: fonte/tamanho/estilo).
            - Para múltiplas linhas, clona o parágrafo inteiro (p) do p0 para cada linha,
            mantendo pPr e rPr. Assim, 2º parágrafo fica 100% igual ao 1º.
            """
            from copy import deepcopy

            # Garante ao menos 1 parágrafo
            if not tf.paragraphs:
                p0 = tf.add_paragraph()
            else:
                p0 = tf.paragraphs[0]

            # Normaliza linhas (um parágrafo por linha)
            linhas = (text or "").split("\n")
            if not linhas:
                linhas = [""]

            # --- prepara p0 com a 1ª linha, SEM destruir seus runs ---
            # se não existir run, cria um vazio para “pegar” a rPr do parágrafo
            if not p0.runs:
                r = p0.add_run()
                r.text = ""  # recebe rPr default do parágrafo/shape

            # mantém somente o 1º run de p0 (preserva rPr dele)
            first_run = p0.runs[0]
            for r in list(p0.runs)[1:]:
                try:
                    r._r.getparent().remove(r._r)
                except Exception:
                    pass
            first_run.text = linhas[0]

            # --- remove parágrafos excedentes (se já existirem) ---
            while len(tf.paragraphs) > 1:
                p_last = tf.paragraphs[-1]
                tf._element.remove(p_last._p)

            # --- gera parágrafos restantes clonando p0 integralmente ---
            for linha in linhas[1:]:
                # clona o parágrafo inteiro (p0._p) => preserva pPr e rPr
                clone_p = deepcopy(p0._p)
                tf._element.append(clone_p)
                # obtém o objeto paragraph recém-adicionado
                p = tf.paragraphs[-1]

                # garante um run e substitui somente o texto do 1º run
                if not p.runs:
                    rr = p.add_run(); rr.text = ""
                # remove runs além do primeiro (mantendo rPr do primeiro)
                for r in list(p.runs)[1:]:
                    try:
                        r._r.getparent().remove(r._r)
                    except Exception:
                        pass
                p.runs[0].text = linha


        def _replace_placeholders_textlike(s: str, token_map, counts: dict) -> str:
            out = s or ""
            for k, v in token_map.items():
                # {{  chave  }} tolerante a espaços e case-insensitive
                pat = re.compile(r"\{\{\s*" + re.escape(k) + r"\s*\}\}", re.IGNORECASE)
                new_out, n = pat.subn(v, out)
                if n:
                    counts[k] = counts.get(k, 0) + n
                out = new_out
            return out

                # ---------- template ----------
        if template_path is None:
            template_path = str(Path(__file__).with_name("PARECER_MODELO2.pptx"))
        if not Path(template_path).exists():
            raise FileNotFoundError(f"Template PPTX não encontrado em: {template_path}")

        # ---------- carrega JSON ----------
        with open(arquivo_json, "r", encoding="utf-8") as f:
            dados = json.load(f) or {}
        d = {(k.lower() if isinstance(k, str) else k): v for k, v in dados.items()}

        # ---------- formatadores (simples; template manda no estilo) ----------
        def _as_str(v): 
            return "" if v is None else str(v)

        def _as_pipe(v):
            if isinstance(v, list):
                return " | ".join(str(x) for x in v if str(x).strip())
            return _as_str(v)

        def _as_paragraphs(v):
            # lista -> um parágrafo por item (sem linha em branco extra)
            if isinstance(v, list):
                return "\n".join(str(x) for x in v if str(x).strip())
            return _as_str(v)

        def _as_formacao(v):
            # linhas simples; o template decide bullets/estilo
            linhas = []
            for f in (v or []):
                linhas.append(
                    f"{f.get('grau','')} em {f.get('curso','')} — "
                    f"{f.get('instituicao','')} ({f.get('conclusao','')})"
                )
            return "\n".join([ln for ln in linhas if ln.strip()])

        # ---------- responsável + data (prioriza argumento da função) ----------
        resp_arg  = (responsavel or "").strip()
        resp_json = _as_str(d.get("responsavel")).strip()
        resp_final = resp_arg or resp_json or "Responsável"

        data_parecer_str = date.today().strftime("%d/%m/%Y")

        # ---------- valores a preencher ----------
        values = {
            "nome": _as_str(d.get("nome")),
            "disponibilidade": _as_str(d.get("disponibilidade")),
            "modalidade": _as_str(d.get("modalidade")),
            "dados_pessoais": _as_str(d.get("dados_pessoais")),
            "perfil_profissional": _as_paragraphs(d.get("perfil_profissional")),
            "perfil_comportamental": _as_paragraphs(d.get("perfil_comportamental")),
            "competencias": _as_pipe(d.get("competencias")),
            "formacao": _as_formacao(d.get("formacao")),

            # campos individuais
            "responsavel": resp_final,
            "data_parecer": data_parecer_str,

            # linha combinada (use shape 'responsavel_data' ou placeholder correspondente)
            "responsavel_data": f"Responsável: {resp_final}  Data do parecer: {data_parecer_str}",
        }


        # ---------- abre modelo ----------
        prs = Presentation(template_path)

        # Índice de shapes por NOME normalizado (slides + layouts + masters)
        shapes_by_name = {}
        def _indexar(container):
            for shp in _iter_shapes(container):
                nm = _normalize(getattr(shp, "name", ""))
                if nm:
                    shapes_by_name.setdefault(nm, []).append(shp)

        for slide in prs.slides: _indexar(slide)
        for master in prs.slide_masters:
            _indexar(master)
            for layout in master.slide_layouts: _indexar(layout)
        for layout in prs.slide_layouts: _indexar(layout)

        # ---------- 1ª passada: preencher por NOME de shape ----------
        filled_by_name = set()
        for key, val in values.items():
            nm = _normalize(key)
            if nm in shapes_by_name:
                for shp in shapes_by_name[nm]:
                    if getattr(shp, "has_text_frame", False):
                        _set_text_preservando_estilo(shp.text_frame, val)
                    elif getattr(shp, "has_table", False):
                        try:
                            cell = shp.table.cell(0,0)
                            _set_text_preservando_estilo(cell.text_frame, val)
                        except Exception:
                            pass
                filled_by_name.add(key)

        # ---------- 2ª passada: substituir placeholders {{chave}} onde sobrou ----------
        remaining = {k: v for k, v in values.items() if k not in filled_by_name}
        placeholder_hits = {}
        if remaining:
            # slides
            for slide in prs.slides:
                for shp in _iter_shapes(slide):
                    if getattr(shp, "has_text_frame", False):
                        updated = _replace_placeholders_textlike(shp.text_frame.text, remaining, placeholder_hits)
                        if updated != shp.text_frame.text:
                            _set_text_preservando_estilo(shp.text_frame, updated)
                    if getattr(shp, "has_table", False):
                        for row in shp.table.rows:
                            for cell in row.cells:
                                upd = _replace_placeholders_textlike(cell.text, remaining, placeholder_hits)
                                if upd != cell.text:
                                    _set_text_preservando_estilo(cell.text_frame, upd)
            # masters e layouts
            for master in prs.slide_masters:
                for shp in _iter_shapes(master):
                    if getattr(shp, "has_text_frame", False):
                        updated = _replace_placeholders_textlike(shp.text_frame.text, remaining, placeholder_hits)
                        if updated != shp.text_frame.text:
                            _set_text_preservando_estilo(shp.text_frame, updated)
                    if getattr(shp, "has_table", False):
                        for row in shp.table.rows:
                            for cell in row.cells:
                                upd = _replace_placeholders_textlike(cell.text, remaining, placeholder_hits)
                                if upd != cell.text:
                                    _set_text_preservando_estilo(cell.text_frame, upd)
            for layout in prs.slide_layouts:
                for shp in _iter_shapes(layout):
                    if getattr(shp, "has_text_frame", False):
                        updated = _replace_placeholders_textlike(shp.text_frame.text, remaining, placeholder_hits)
                        if updated != shp.text_frame.text:
                            _set_text_preservando_estilo(shp.text_frame, updated)
                    if getattr(shp, "has_table", False):
                        for row in shp.table.rows:
                            for cell in row.cells:
                                upd = _replace_placeholders_textlike(cell.text, remaining, placeholder_hits)
                                if upd != cell.text:
                                    _set_text_preservando_estilo(cell.text_frame, upd)

        # ---------- salva (sem tocar no modelo) ----------
        Path(arquivo_saida).parent.mkdir(parents=True, exist_ok=True)
        prs.save(arquivo_saida)

        # ---------- logs (opcional) ----------
        try:
            missing = [k for k in values.keys() if k not in filled_by_name and k not in placeholder_hits]
            print(f"[PPTX] Por nome: {sorted(list(filled_by_name))}")
            print(f"[PPTX] Por placeholder: {sorted([k for k in placeholder_hits.keys()])}")
            if missing:
                print("[PPTX] Sem destino no template:", ", ".join(missing))
        except Exception:
            pass

        return arquivo_saida
















    def process_text_curriculo(self, texto):
        """Processa o texto e retorna JSON estruturado."""
        load_dotenv()
        chave_api = os.getenv('OPENAI_API_KEY')
        openai.api_key = chave_api

        if not chave_api:
            st.error("Chave da API OpenAI não encontrada.")
            return {}

        modelo_prompt = f"""
                            TEXTO DO CURRÍCULO:
                            {texto}

                            Campos esperados e explicações:
                            1. **informacoes_pessoais**: 
                                Contém as informações pessoais do candidato, incluindo:
                                - "nome": Nome completo do candidato.
                                - "cidade": Cidade e estado de residência.
                                - "email": Endereço de e-mail de contato.
                                - "telefone": Número de telefone para contato.
                                - "cargo": Cargo atual ou pretendido.

                            2. **resumo_qualificacoes**:
                                Lista com as principais habilidades, competências ou realizações do candidato, como:
                                - Conhecimentos técnicos (ex.: Power BI, Python, SQL).
                                - Soft skills (ex.: liderança, trabalho em equipe).
                                - Principais realizações (ex.: "Aumentou a eficiência em X% ao implementar [projeto]").

                            3. **experiencia_profissional**:
                                Lista de experiências profissionais relevantes. Cada entrada deve conter:
                                - "empresa": Nome da empresa.
                                - "cargo": Cargo exercido.
                                - "periodo": Período de atuação (ex.: Janeiro de 2020 - Dezembro de 2022).
                                - "local": Local onde o trabalho foi realizado (ex.: Remoto ou Cidade/Estado).
                                - "atividades_exercidas": Lista de atividades e responsabilidades no cargo. Detalhe as principais contribuições e tarefas realizadas.
                                - "ferramentas": Lista das ferramentas, softwares ou tecnologias utilizadas no cargo (ex.: Power BI, Python, SQL, Tableau).

                            4. **educacao**:
                                Lista de formações acadêmicas do candidato. Cada entrada deve conter:
                                - "instituicao": Nome da instituição de ensino.
                                - "curso": Curso ou programa concluído.
                                - "periodo": Período de realização (ex.: Janeiro de 2016 - Dezembro de 2020).

                            5. **certificacoes**:
                                Lista de certificações relevantes obtidas pelo candidato. Cada entrada deve conter:
                                - Nome da certificação (ex.: "Microsoft Certified: Data Analyst Associate").
                                - Instituição que emitiu a certificação (ex.: Microsoft, AWS, etc.).

                            Formato esperado do JSON de saída:
                            {{
                                "informacoes_pessoais": {{
                                    "nome": "",
                                    "cidade": "",
                                    "email": "",
                                    "telefone": "",
                                    "cargo": ""
                                }},
                                "resumo_qualificacoes": [
                                    "Resumo 1",
                                    "Resumo 2"
                                ],
                                "experiencia_profissional": [
                                    {{
                                        "empresa": "Empresa X",
                                        "cargo": "Cargo Y",
                                        "periodo": "Janeiro de 2020 - Dezembro de 2022",
                                        "local": "Cidade/Estado",
                                        "atividades_exercidas": [
                                            "Atividade 1",
                                            "Atividade 2"
                                        ],
                                        "ferramentas": [
                                            "Ferramenta 1",
                                            "Ferramenta 2"
                                        ]
                                    }}
                                ],
                                "educacao": [
                                    {{
                                        "instituicao": "Instituição A",
                                        "curso": "Curso B",
                                        "periodo": "Janeiro de 2016 - Dezembro de 2020"
                                    }}
                                ],
                                "certificacoes": [
                                    "Certificação 1",
                                    "Certificação 2"
                                ]
                            }}
                            """

        try:
            response = openai.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": """Você é um especialista em análise de currículos e extração de informações.
                                                     Colete todas as informações possíveis, não deixe nada passar. 
                                                     Dê sua resposta APENAS com o json solicitado e nada mais. NÃO ESCREVA ```json na resposta!
                    """},
                    {"role": "user", "content": modelo_prompt}
                ],
                temperature=0.2,
                max_tokens=4096
            )
            
            conteudo = response.choices[0].message.content.replace("```json", "").strip()
            # st.write(f"CONTEUDO: {conteudo}")

            try:
                return json.loads(conteudo)
            except json.JSONDecodeError:
                print("Erro ao converter resposta da API para JSON.")
                return {}
        except Exception as e:
            print(f"Erro ao processar texto com a API OpenAI: {e}")
            return {}

    def process_text_parecer(self, texto):
        """Processa o texto e retorna JSON estruturado."""
        load_dotenv()
        chave_api = os.getenv('OPENAI_API_KEY')
        openai.api_key = chave_api

        if not chave_api:
            st.error("Chave da API OpenAI não encontrada.")
            return {}

        modelo_prompt_parecer = f"""
            TEXTO DO CURRÍCULO ORIGINAL:
            {texto}

            ### INSTRUÇÕES
            - Extraia só o que estiver presente no currículo; não invente dados.
            - Preencha **todos** os campos abaixo sempre que encontrar a informação.
            - **Formato de saída**: JSON **sem** crases, sem ```json, sem comentários.

            ### CAMPOS E PADRÕES ESPERADOS

            1. Nome
            • Nome do candidato 

            2. formacao (lista de objetos)
            • grau        → "Tecnólogo", "Bacharel", "MBA", etc.
            • curso       → Nome do curso
            • instituicao → Onde cursou
            • conclusao   → "2018", "cursando", etc.

            3. Competencias (lista de competencias)
                                Lista com as principais (no máximo 5)competências do candidato, como:
                                - Competências(ex.: Power BI, Python, SQL).

            4. perfil_profissional (listagem de 2 parágrafos, nesta ordem)
            • Parágrafo 1 – trajetória (empresas, cargos, período, volume de entregas).  
            • Parágrafo 2 – competências + projetos relevantes iniciados por verbo no infinitivo/gerúndio.

            ### EXEMPLO DE SAÍDA ESPERADA
            "Nome": "João da Silva",
            "formacao": [
                {{
                "grau": "Tecnólogo"}}]
            {{
            "formacao": [
                {{
                "grau": "Tecnólogo",
                "curso": "Análise de Sistemas",
                "instituicao": "Faculdade X",
                "conclusao": "2012"
                }},
                {{
                "grau": "MBA",
                "curso": "Gestão de Projetos",
                "instituicao": "Universidade Y",
                "conclusao": "2019"
                }}
            ],
            "competencias": [
                "Conhecimentos técnicos: Power BI, Python, SQL.",
                "Soft skills: liderança, trabalho em equipe."
            ],
            "perfil_profissional": [
                "Camila atua desde fevereiro de 2021 na empresa Sankhya como Consultora de Implantação de ERP Sênior – módulo HCM, participando de 15 projetos de implantação e conduzindo treinamentos para clientes em vários estados. Antes disso, trabalhou na Solar Coca-Cola, YDUQS e Adtalem com foco em SAP HCM, somando experiência prévia de seis anos em rotinas de departamento pessoal.",
                "Domina metodologias ágeis e Waterfall, conduz migrações de dados de sistemas legados, parametriza folha, ponto e avaliação de desempenho e implanta soluções de ERP. Implantou dois novos Centros de Distribuição e uma loja, integrou plataformas Totvs e Fortes e automatizou rotinas de importação de pedidos, entregando ganhos de produtividade em até 5 meses."
            ]
            }}
            """
        try:
            response = openai.chat.completions.create(
                model="gpt-4o",
                messages=[
                    {"role": "system", "content": """Você é um especialista em análise de currículos e extração de informações.
                                                     Colete todas as informações possíveis, não deixe nada passar. 
                                                     Dê sua resposta APENAS com o json solicitado e nada mais. NÃO ESCREVA ```json na resposta!
                    """},
                    {"role": "user", "content": modelo_prompt_parecer}
                ],
                temperature=0.2,
                max_tokens=4096
            )
            
            conteudo = response.choices[0].message.content.replace("```json", "").strip()
            print(f"CONTEUDO: {conteudo}")

            try:
                return json.loads(conteudo)
            except json.JSONDecodeError:
                print("Erro ao converter resposta da API para JSON.")
                return {}
        except Exception as e:
            print(f"Erro ao processar texto com a API OpenAI: {e}")
            return {}



    def extract_text_from_pdf(self, caminho_pdf):
        """Extrai o texto de um arquivo PDF."""
        try:
            leitor = PdfReader(caminho_pdf)
            texto = "".join(pagina.extract_text() for pagina in leitor.pages)
            return self.clear_text(texto)
        except Exception as e:
            print(f"Erro ao extrair texto do PDF: {e}")
            return ""

    def clear_text(self, texto):
        """Limpa e normaliza o texto extraído."""
        texto = re.sub(r'\s+', ' ', texto)
        texto = re.sub(r'\n*Página \d+ de \d+\n*', '', texto)
        return texto.strip()

    # Função para adicionar uma imagem de fundo a partir de um arquivo local
    def add_bg_from_local(self, image_file):
        with Path(image_file).open("rb") as file:
            encoded_string = base64.b64encode(file.read()).decode()
        st.markdown(
            f"""
            <style>
            .stApp {{
                background-color: rgba(247,247,247,0.75);
                background-size: contain;
                background-position: center;
                background-repeat: no-repeat;
                border-color: rgba(31,216,135,1) ;
    
            }}
            </style>
            """,
            unsafe_allow_html=True
        )

    # Função para exibir uma imagem de logo no topo a partir de um arquivo local
    # def add_logo_from_local(self, logo_file):
    #     with Path(logo_file).open("rb") as file:
    #         encoded_string = base64.b64encode(file.read()).decode()
    #     st.markdown(
    #         f"""
    #         <style>
    #         [data-testid="stAppViewContainer"] > .main {{
    #             padding-top: 0px;
    #         }}
    #         .logo-container {{
    #             display: flex;
    #             justify-content: center;
    #             align-items: center;
    #             padding: 1rem 0;
    #         }}
    #         .logo-container img {{
    #             max-width: 200px;
    #             height: auto;
    #         }}
    #         </style>
    #         <div class="logo-container">
    #             <img src="data:image/png;base64,{encoded_string}" alt="Logo">
    #         </div>
    #         """,
    #         unsafe_allow_html=True
    #     )
    def add_logo_from_local(self, logo_file):
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
                padding: 1vh 0; /* margem baseada na altura da tela */
            }}
            .logo-container img {{
                max-height: 20vh; /* altura máxima baseada na tela */
                max-width: 80vw;  /* largura máxima baseada na tela */
                height: auto;
                width: auto;
            }}

            /* Ajuste fino para telas pequenas (MacBook e similares) */
            @media only screen and (max-width: 1440px) {{
                .logo-container {{
                    padding: 3vh 0;
                }}
                .logo-container img {{
                    max-height: 10vh;
                }}
            }}
            </style>
            <div class="logo-container">
                <img src="data:image/png;base64,{encoded_string}" alt="Logo">
            </div>
            """,
            unsafe_allow_html=True
        )


