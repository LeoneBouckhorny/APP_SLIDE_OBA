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

# Exibe a logo no topo (descomente a linha abaixo se você tiver o arquivo 'logo_jornada.png')
# st.image("logo_jornada.png", use_container_width=True)

st.title("🚀 Gerador Automático de Slides")
st.info("Faça o upload do arquivo .docx com os dados e do arquivo .pptx modelo para gerar a apresentação.")

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
        alunos = sorted([m for m in membros if "aluno" in m["Funcao"]],
                        key=lambda m: formatar_texto(m["Nome"]))

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
            "{{NOME_EQUIPE}}": f"Equipe: {equipe_nome.split()[-1]}",
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
    if not shape.has_text_frame:
        return

    # Trata o caso especial de {{NOMES_ALUNOS}} primeiro, pois ele reestrutura toda a caixa de texto
    if "{{NOMES_ALUNOS}}" in shape.text:
        tf = shape.text_frame
        tf.clear()  # Limpa todo o conteúdo existente

        nomes = team_data.get("{{NOMES_ALUNOS}}", "").split("\n")

        # Adiciona cada nome como um novo parágrafo formatado
        for i, nome in enumerate(nomes):
            p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
            p.text = nome
            p.font.name = "Lexend"
            p.font.bold = True
            p.font.size = Pt(26.5)
            p.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p.alignment = PP_ALIGN.CENTER
        return # Termina a função para esta shape, pois já foi completamente tratada

    # Para os outros placeholders, itera sobre os parágrafos e 'runs'
    for paragraph in shape.text_frame.paragraphs:
        # É preciso iterar sobre uma cópia da lista de runs, pois vamos modificá-la
        for run in list(paragraph.runs):
            full_text = run.text
            for key, value in team_data.items():
                if key in full_text:
                    # Substitui o texto
                    run.text = full_text.replace(key, value)
                    
                    # Aplica formatação específica para cada chave
                    if key == "{{LANCAMENTOS_VALIDOS}}":
                        paragraph.clear() # Limpa o parágrafo para recriar com 2 cores
                        
                        match = re.match(r"(ALCANCE:\s*)(.*)", value, re.IGNORECASE)
                        if match:
                            prefixo, valor_numerico = match.groups()
                            
                            # Run para "ALCANCE: "
                            run1 = paragraph.add_run()
                            run1.text = prefixo
                            run1.font.name = "Lexend"
                            run1.font.size = Pt(28)
                            run1.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)
                            
                            # Run para o número
                            run2 = paragraph.add_run()
                            run2.text = valor_numerico
                            run2.font.name = "Lexend"
                            run2.font.bold = True
                            run2.font.underline = True
                            run2.font.size = Pt(35)
                            run2.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)
                    
                    elif key == "{{NOME_EQUIPE}}":
                        run.font.name = "Lexend"
                        run.font.bold = True
                        run.font.size = Pt(20)
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    
                    elif key in ("{{NOME_ESCOLA}}", "{{CIDADE_UF}}"):
                        run.font.name = "Lexend"
                        run.font.bold = True
                        run.font.size = Pt(22)
                        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                    
                    # Garante alinhamento central para placeholders simples
                    if key in ("{{NOME_EQUIPE}}", "{{NOME_ESCOLA}}", "{{CIDADE_UF}}"):
                         paragraph.alignment = PP_ALIGN.CENTER

def gerar_apresentacao(dados, template_stream):
    """
    LÓGICA CORRETA:
    1. Carrega a apresentação modelo (com 1 slide).
    2. Duplica o slide modelo (limpo) N-1 vezes, onde N é o número total de equipes.
       Agora temos N slides idênticos ao modelo.
    3. Itera por todos os N slides e todos os N dados, preenchendo cada slide
       com os dados da equipe correspondente.
    """
    prs = Presentation(template_stream)
    if not dados or not prs.slides:
        return prs

    # Pega uma referência ao slide modelo original
    slide_modelo = prs.slides[0]

    # Cria N-1 cópias do slide modelo original e LIMPO
    # onde N é o número total de equipes.
    for _ in range(len(dados) - 1):
        duplicate_slide_with_media(prs, slide_modelo)

    # Agora que temos o número certo de slides, iteramos por eles
    # e pelos dados para preencher cada um.
    # `zip` combina cada slide com os dados de uma equipe.
    for slide, team_data in zip(prs.slides, dados):
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, team_data)

    return prs

# -------------------- INTERFACE STREAMLIT --------------------
docx_file = st.file_uploader("📄 Arquivo DOCX com os dados das equipes", type=["docx"])
pptx_file = st.file_uploader("📊 Arquivo PPT
