import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN
from pptx.util import Pt
from pptx.dml.color import RGBColor
from collections import defaultdict
from copy import deepcopy
from io import BytesIO
from lxml import etree
import re

# ---------- CONFIGURAÇÃO INICIAL ----------
st.set_page_config(layout="wide")
st.image("logo_jornada.png", use_container_width=True)
st.title("🚀 Gerador Automático de Slides")
st.info("")

# -------------------- FUNÇÕES AUXILIARES --------------------
def formatar_texto(texto, maiusculo_estado=False):
    texto = ' '.join(texto.strip().split())
    return texto.upper() if maiusculo_estado else ' '.join(w.capitalize() for w in texto.split())

def extrair_dados(uploaded_file):
    doc = Document(uploaded_file)
    registros = []
    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            if i == 0:
                continue
            celulas = [c.text.strip() for c in linha.cells]
            if len(celulas) >= 8:
                _, valido, equipe, funcao, escola, cidade, estado, nome = celulas[:8]
                registros.append({
                    "Valido": valido,
                    "Equipe": equipe,
                    "Funcao": funcao.lower(),
                    "Escola": escola,
                    "Cidade": cidade,
                    "Estado": estado,
                    "Nome": nome
                })

    equipes = defaultdict(list)
    for r in registros:
        equipes[r["Equipe"]].append(r)

    def chave_ord(membros):
        try:
            return float(membros[0]["Valido"].replace(",", "."))
        except:
            return float("inf")

    equipes_ordenadas = sorted(equipes.items(), key=lambda x: chave_ord(x[1]))

    dados_finais = []
    for equipe_nome, membros in equipes_ordenadas:
        lider = [m for m in membros if "líder" in m["Funcao"] or "lider" in m["Funcao"]]
        acompanhante = [m for m in membros if "acompanhante" in m["Funcao"]]
        alunos = sorted([m for m in membros if "aluno" in m["Funcao"]], key=lambda m: formatar_texto(m["Nome"]))

        nomes_lider = formatar_texto(lider[0]["Nome"]) if lider else ""
        nomes_acompanhante = formatar_texto(acompanhante[0]["Nome"]) if acompanhante else ""

        linhas_nomes = []
        if nomes_lider:
            linhas_nomes.append(nomes_lider)
        if nomes_acompanhante:
            linhas_nomes.append(nomes_acompanhante)
        linhas_nomes += [formatar_texto(a["Nome"]) for a in alunos]

        nomes_formatados = "\n".join(linhas_nomes)

        info = membros[0]
        dados_finais.append({
            "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info['Valido']} m",
            "{{NOME_EQUIPE}}": f"Equipe: {re.sub(r'[^0-9]+', '', equipe_nome)}" if re.search(r'\d', equipe_nome) else f"Equipe: {equipe_nome.strip()}",
            "{{NOME_ESCOLA}}": formatar_texto(info["Escola"]),
            "{{CIDADE_UF}}": f"{formatar_texto(info['Cidade'])} / {formatar_texto(info['Estado'], True)}",
            "{{NOMES_ALUNOS}}": nomes_formatados
        })
    return dados_finais

def duplicate_slide_with_media(prs, source_slide):
    layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                img_blob = shape.image.blob
            except Exception:
                img_blob = None
            if img_blob:
                image_part, new_rId = new_slide.part.get_or_add_image_part(BytesIO(img_blob))
                new_el_xml = etree.fromstring(new_el.xml)
                blips = new_el_xml.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                for blip in blips:
                    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', new_rId)
                from pptx.oxml import parse_xml
                new_el = parse_xml(etree.tostring(new_el_xml, encoding='utf-8'))
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    return new_slide

def replace_placeholders_in_shape(shape, team_data):
    """Substitui placeholders dentro de cada forma do slide."""
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame

    for paragraph in list(text_frame.paragraphs):
        full_text = "".join(run.text for run in paragraph.runs)
        selected_key = None

        # Identifica o placeholder
        for key in team_data.keys():
            if key in full_text:
                selected_key = key
                break

        if not selected_key:
            continue

        # Substitui o texto do placeholder pelo valor correspondente
        new_text = full_text.replace(selected_key, team_data[selected_key])

        # Limpa os runs anteriores
        for _ in range(len(paragraph.runs)):
            paragraph._p.remove(paragraph.runs[0]._r)

        # --- Regras específicas de formatação ---
        if selected_key == "{{LANCAMENTOS_VALIDOS}}":
            match = re.match(r"(ALCANCE:\s*)([\d,.]+ m)", new_text, re.IGNORECASE)
            if match:
                prefix, valor = match.groups()

                run1 = paragraph.add_run()
                run1.text = prefix
                run1.font.name = "Lexend"
                run1.font.bold = False
                run1.font.size = Pt(28)
                run1.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)

                run2 = paragraph.add_run()
                run2.text = valor
                run2.font.name = "Lexend"
                run2.font.bold = True
                run2.font.underline = True
                run2.font.size = Pt(35)
                run2.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)

        elif selected_key == "{{NOMES_ALUNOS}}":
            text_frame.clear()
            lines = new_text.split("\n")
            for i, nome in enumerate(lines):
                p = text_frame.add_paragraph() if i > 0 else text_frame.paragraphs[0]
                run = p.add_run()
                run.text = nome
                run.font.name = "Lexend"
                run.font.bold = True
                run.font.size = Pt(26.5)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                p.alignment = PP_ALIGN.CENTER

        elif selected_key == "{{NOME_EQUIPE}}":
            run = paragraph.add_run()
            run.text = new_text  # 🔥 Aqui ele de fato escreve o texto
            run.font.name = "Lexend"
            run.font.bold = True
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            paragraph.alignment = PP_ALIGN.CENTER

        elif selected_key in ("{{NOME_ESCOLA}}", "{{CIDADE_UF}}"):
            run = paragraph.add_run()
            run.text = new_text
            run.font.name = "Lexend"
            run.font.bold = True
            run.font.size = Pt(22)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            paragraph.alignment = PP_ALIGN.CENTER

        else:
            run = paragraph.add_run()
            run.text = new_text
            run.font.name = "Lexend"
            run.font.bold = True
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            paragraph.alignment = PP_ALIGN.CENTER


def gerar_apresentacao(dados, template_stream):
    prs = Presentation(template_stream)
    if not dados or not prs.slides:
        return prs

    modelo = prs.slides[0]
    slides_gerados = [modelo]
    for _ in range(len(dados) - 1):
        novo_slide = duplicate_slide_with_media(prs, modelo)
        slides_gerados.append(novo_slide)

    for slide, team in zip(slides_gerados, dados):
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, team)

    return prs

# -------------------- INTERFACE STREAMLIT --------------------
docx_file = st.file_uploader("📄 Arquivo DOCX", type=["docx"])
pptx_file = st.file_uploader("📊 Arquivo PPTX modelo", type=["pptx"])

if st.button("✨ Gerar Apresentação"):
    if not docx_file or not pptx_file:
        st.warning("Envie ambos os arquivos.")
    else:
        try:
            dados = extrair_dados(docx_file)
            if not dados:
                st.warning("Nenhum dado encontrado.")
            else:
                prs_final = gerar_apresentacao(dados, pptx_file)
                buf = BytesIO()
                prs_final.save(buf)
                buf.seek(0)
                st.success(f"Slides gerados: {len(dados)}")
                st.image("tiapamela.gif", caption="Apresentação pronta! 🚀", use_container_width=True)
                st.download_button(
                    "📥 Baixar Apresentação Final",
                    data=buf,
                    file_name="Apresentacao_Final_Equipes.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Erro ao gerar apresentação: {e}")



