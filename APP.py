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

# Exibe a logo no topo
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

    full_text_shape = "".join(run.text for p in shape.text_frame.paragraphs for run in p.runs)

    for key, value in team_data.items():
        if key not in full_text_shape:
            # Tenta encontrar placeholder parcialmente (ex: {{NOME_EQUIPE}} quebrado)
            if any(k.replace("{", "").replace("}", "") in full_text_shape for k in team_data.keys()):
                pass
            else:
                continue

        tf = shape.text_frame
        tf.clear()

        if key == "{{NOMES_ALUNOS}}":
            linhas = value.split("\n")
            for i, nome in enumerate(linhas):
                p = tf.add_paragraph() if i > 0 else tf.paragraphs[0]
                run = p.add_run()
                run.text = nome
                run.font.name = "Lexend"
                run.font.bold = True
                run.font.size = Pt(26.5)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
                p.alignment = PP_ALIGN.CENTER
                p.line_spacing = None

        elif key == "{{LANCAMENTOS_VALIDOS}}":
            p = tf.paragraphs[0]
            match = re.match(r"(ALCANCE:\s*)([\d,.]+ m)", value, re.IGNORECASE)
            if match:
                prefix, numero = match.groups()
                run1 = p.add_run()
                run1.text = prefix
                run1.font.name = "Lexend"
                run1.font.bold = False
                run1.font.size = Pt(28)
                run1.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)

                run2 = p.add_run()
                run2.text = numero
                run2.font.name = "Lexend"
                run2.font.bold = True
                run2.font.underline = True
                run2.font.size = Pt(35)
                run2.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)
            else:
                run = p.add_run()
                run.text = value
                run.font.name = "Lexend"
                run.font.bold = True
                run.font.size = Pt(35)
                run.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)

        else:
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = value
            run.font.name = "Lexend"
            run.font.bold = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p.alignment = PP_ALIGN.CENTER
            p.line_spacing = None

            if key == "{{NOME_EQUIPE}}":
                run.font.size = Pt(20)
            elif key in ("{{NOME_ESCOLA}}", "{{CIDADE_UF}}"):
                run.font.size = Pt(22)
            else:
                run.font.size = Pt(18)


def gerar_apresentacao(dados_equipes, arquivo_pptx_modelo):
    """
    Gera uma apresentação PowerPoint a partir de um modelo.
    Estratégia:
    1. Modifica o primeiro slide do modelo com os dados da primeira equipe.
    2. Usa o primeiro slide (agora modificado) como base para duplicar e criar os slides das equipes restantes.
    Isso evita a necessidade de deletar o slide modelo, prevenindo o erro "Element is not a child of this node".
    """
    prs = Presentation(arquivo_pptx_modelo)

    # Verifica se há slides no modelo e dados para processar
    if not prs.slides:
        raise ValueError("A apresentação modelo está vazia.")
    if not dados_equipes:
        # Se não houver dados, retorna a apresentação original sem modificações
        return prs

    # --- Passo 1: Processar a primeira equipe ---
    # Pega o slide modelo (o primeiro e único)
    slide_a_ser_usado_como_modelo = prs.slides[0]
    
    # Pega os dados da primeira equipe
    primeira_equipe = dados_equipes[0]
    
    # Preenche o primeiro slide com os dados da primeira equipe
    for shape in slide_a_ser_usado_como_modelo.shapes:
        replace_placeholders_in_shape(shape, primeira_equipe)

    # --- Passo 2: Processar as equipes restantes ---
    equipes_restantes = dados_equipes[1:]
    
    for dados_equipe in equipes_restantes:
        # Duplica o primeiro slide (que agora já é o slide da primeira equipe)
        novo_slide = duplicate_slide_with_media(prs, slide_a_ser_usado_como_modelo)
        
        # Preenche as informações no novo slide duplicado
        for shape in novo_slide.shapes:
            replace_placeholders_in_shape(shape, dados_equipe)

    return prs

# -------------------- INTERFACE STREAMLIT --------------------
docx_file = st.file_uploader("📄 Arquivo DOCX", type=["docx", "DOCX"])
pptx_file = st.file_uploader("📊 Arquivo PPTX modelo", type=["pptx", "PPTX"])

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





