import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from collections import defaultdict
from copy import deepcopy
from io import BytesIO
from lxml import etree

# ---------------- utilitários ----------------
def formatar_texto(texto, maiusculo_estado=False):
    texto = ' '.join(texto.strip().split())
    return texto.upper() if maiusculo_estado else ' '.join(w.capitalize() for w in texto.split())

def extrair_dados(uploaded_file):
    """Lê o DOCX e retorna lista de dicionários prontos para substituição."""
    doc = Document(uploaded_file)
    registros = []
    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            if i == 0: continue
            celulas = [c.text.strip() for c in linha.cells]
            if len(celulas) >= 8:
                # ignoramos a coluna Medalha (celulas[0])
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

    # ordenar por 'Valido' crescente
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

        blocos = []
        if lider: blocos.append(formatar_texto(lider[0]["Nome"]))
        if acompanhante: blocos.append(formatar_texto(acompanhante[0]["Nome"]))
        blocos += [formatar_texto(a["Nome"]) for a in alunos]

        info = membros[0]
        # chaves conforme seu modelo: {{LANCAMENTOS_VALIDOS}} etc.
        dados_finais.append({
            "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info['Valido']} m",
            "{{NOME_EQUIPE}}": f"Equipe: {equipe_nome.split()[-1]}",
            "{{NOME_ESCOLA}}": formatar_texto(info["Escola"]),
            "{{CIDADE_UF}}": f"{formatar_texto(info['Cidade'])} / {formatar_texto(info['Estado'], True)}",
            "{{NOME_LIDER}}": formatar_texto(lider[0]["Nome"]) if lider else "",
            "{{NOME_ACOMPANHANTE}}": formatar_texto(acompanhante[0]["Nome"]) if acompanhante else "",
            "{{NOMES_ALUNOS}}": "\n".join(blocos[(2 if (lider or acompanhante) else 0):])
        })
    return dados_finais

# --------- cópia robusta do slide (preserva imagens) ----------
def duplicate_slide_with_media(prs, source_slide):
    """
    Cria um novo slide no prs usando o mesmo layout do source_slide,
    copia shapes e images (copiando as partes de mídia e atualizando r:embed).
    Retorna o novo slide.
    """
    layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)

    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        # se for imagem, precisamos copiar a imagem para new_slide.part e trocar rId no XML
        if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
            try:
                img_blob = shape.image.blob
            except Exception:
                img_blob = None
            if img_blob:
                # adiciona a imagem ao pacote do new_slide (reusa se já existir)
                image_part, new_rId = new_slide.part.get_or_add_image_part(BytesIO(img_blob))
                # atualiza o r:embed no XML do novo elemento
                new_el_xml = etree.fromstring(new_el.xml)
                blips = new_el_xml.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip')
                for blip in blips:
                    blip.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed', new_rId)
                from pptx.oxml import parse_xml
                new_el = parse_xml(etree.tostring(new_el_xml, encoding='utf-8'))
        # insere o elemento copiado no novo slide
        new_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')

    return new_slide

# --------- substituição de texto (robusta) ----------
def replace_placeholders_in_shape(shape, team_data):
    if not shape.has_text_frame:
        return
    for paragraph in shape.text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        new_text = full_text
        for k, v in team_data.items():
            if k in new_text:
                new_text = new_text.replace(k, v)
        if new_text != full_text:
            # remove runs e insere novo texto simples
            while paragraph.runs:
                paragraph._p.remove(paragraph.runs[0]._r)
            paragraph.add_run().text = new_text

# --------- geração final ----------
def gerar_apresentacao(dados, template_stream):
    prs = Presentation(template_stream)
    if not dados or not prs.slides:
        return prs

    # Substitui o primeiro slide com a primeira equipe
    first_team = dados[0]
    for shape in prs.slides[0].shapes:
        replace_placeholders_in_shape(shape, first_team)

    # Para cada equipe restante, duplica o primeiro slide (com mídia) e substitui
    for team in dados[1:]:
        new_slide = duplicate_slide_with_media(prs, prs.slides[0])
        for shape in new_slide.shapes:
            replace_placeholders_in_shape(shape, team)

    return prs

# ---------------- Streamlit UI ----------------
st.set_page_config(layout="wide")
st.title("🚀 Gerador Automático de Slides (robusto)")
st.info("Envie o DOCX com as equipes e o PPTX modelo (apenas 1 slide modelo)")

docx_file = st.file_uploader("📄 Arquivo DOCX (dados)", type=["docx"])
pptx_file = st.file_uploader("📊 Arquivo PPTX (modelo com placeholders)", type=["pptx"])

if st.button("✨ Gerar Apresentação"):
    if docx_file is None or pptx_file is None:
        st.warning("Envie ambos os arquivos (DOCX e PPTX).")
    else:
        try:
            teams = extrair_dados(docx_file)
            if not teams:
                st.warning("Nenhum dado válido encontrado no DOCX.")
            else:
                prs_final = gerar_apresentacao(teams, pptx_file)
                buf = BytesIO()
                prs_final.save(buf)
                buf.seek(0)
                st.success(f"Gerados {len(teams)} slides.")
                st.download_button(
                    label="📥 Baixar Apresentação Final",
                    data=buf,
                    file_name="Apresentacao_Final_Equipes.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Ocorreu um erro: {e}")
