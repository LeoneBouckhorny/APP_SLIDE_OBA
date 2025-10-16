import streamlit as st
from docx import Document
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from collections import defaultdict
from copy import deepcopy
from io import BytesIO
from lxml import etree
import re

st.set_page_config(layout="wide")
st.title("🔍 Gerador de Slides — Modo Diagnóstico")
st.info("Carregue os arquivos e ative o diagnóstico para ver onde os placeholders estão (use apenas para debug).")

# ---------- Helpers e lógica (mesma base que você já tem) ----------
def formatar_texto(texto, maiusculo_estado=False):
    if texto is None:
        return ""
    texto = ' '.join(str(texto).strip().split())
    return texto.upper() if maiusculo_estado else ' '.join(w.capitalize() for w in texto.split())

def extrair_dados(uploaded_file):
    doc = Document(uploaded_file)
    registros = []
    if not doc.tables:
        st.error("Nenhuma tabela encontrada no DOCX.")
        return []

    tabela = doc.tables[0]
    # tenta mapear por cabeçalho; fallback para ordem esperada
    header = [c.text.strip().lower() for c in tabela.rows[0].cells]
    # leitura por posição (assumindo estrutura conhecida)
    for i, row in enumerate(tabela.rows):
        if i == 0:
            continue
        cells = [c.text.strip() for c in row.cells]
        if len(cells) < 8:
            continue
        # ignora medalha
        _, valido, equipe, funcao, escola, cidade, estado, nome = cells[:8]
        registros.append({
            "Valido": valido,
            "Equipe": equipe,
            "Funcao": funcao.lower(),
            "Escola": escola,
            "Cidade": cidade,
            "Estado": estado,
            "Nome": nome
        })
    # agrupa por equipe
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
        # monta dicionário de placeholders
        dados_finais.append({
            "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info.get('Valido','')} m",
            "{{NOME_EQUIPE}}": f"Equipe: {re.sub(r'[^0-9]+','', equipe_nome) if re.search(r'\\d', equipe_nome) else equipe_nome}",
            "{{NOME_ESCOLA}}": f"{formatar_texto(info.get('Escola',''))}\n{formatar_texto(info.get('Cidade',''))} / {formatar_texto(info.get('Estado',''), True)}",
            "{{CIDADE_UF}}": f"{formatar_texto(info.get('Cidade',''))} / {formatar_texto(info.get('Estado',''), True)}",
            "{{NOMES_ALUNOS}}": nomes_formatados
        })
    return dados_finais

# --- PPT helpers (duplicação + substituição) ---
def duplicate_slide_with_media(prs, source_slide):
    layout = source_slide.slide_layout
    new_slide = prs.slides.add_slide(layout)
    for shape in source_slide.shapes:
        new_el = deepcopy(shape.element)
        if shape.shape_type == 13:  # MSO_SHAPE_TYPE.PICTURE (literal number to avoid extra import)
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

def replace_placeholders_in_shape(shape, team_data, debug_collect):
    """Substitui placeholders e registra infos em debug_collect (lista)."""
    if not shape.has_text_frame:
        return

    tf = shape.text_frame
    # pega texto completo atual
    text_before = ""
    for p in tf.paragraphs:
        for r in p.runs:
            text_before += r.text
    found_keys = [k for k in team_data.keys() if k in text_before]
    debug_collect.append({"shape_before": text_before, "found_keys": found_keys})

    if not found_keys:
        return

    # substitui cada key encontrada no texto completo e escreve parágrafos resultantes
    new_text = text_before
    for k in found_keys:
        new_text = new_text.replace(k, team_data[k])

    # escreve new_text dividindo por linhas
    tf.clear()
    lines = new_text.split("\n")
    for i, line in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.CENTER  # 🔧 Corrigido — garante centralização sem erro
        run = p.add_run()
        run.text = line
        # formatação simples (ajustável)
        run.font.name = "Lexend"
        run.font.bold = True

        # --- Tamanho condicional (ajuste fino) ---
        if "{{LANCAMENTOS_VALIDOS}}" in text_before or "ALCANCE:" in line:
            run.font.size = Pt(28)
            run.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)
        elif "{{NOME_EQUIPE}}" in text_before or "Equipe:" in line:
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        elif "{{NOME_ESCOLA}}" in text_before or "{{CIDADE_UF}}" in text_before:
            run.font.size = Pt(22)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        else:
            run.font.size = Pt(26.5)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # pega texto depois pra debug
    text_after = ""
    for p in tf.paragraphs:
        for r in p.runs:
            text_after += r.text + " "
    debug_collect[-1]["shape_after"] = text_after.strip()


def gerar_apresentacao_debug(dados, template_stream, max_preview=3):
    prs = Presentation(template_stream)
    if not dados or not prs.slides:
        st.error("Template PPTX inválido.")
        return None, []

    modelo = prs.slides[0]
    # duplicar o necessário
    slides_gerados = [modelo]
    for _ in range(len(dados) - 1):
        slides_gerados.append(duplicate_slide_with_media(prs, modelo))

    # coleta debug global
    debug_global = []
    # preencher e registrar info (somente para as primeiras max_preview equipes mostramos)
    for idx, (slide, team) in enumerate(zip(slides_gerados, dados)):
        debug_for_slide = {"team_index": idx, "team_data": team, "shapes": []}
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, team, debug_for_slide["shapes"])
        debug_global.append(debug_for_slide)
        if idx >= max_preview - 1:
            # ainda substitui para todos, mas só coleta debug das primeiras N
            pass
    return prs, debug_global

# -------------------- INTERFACE --------------------
st.header("1. Faça upload dos arquivos")
docx_file = st.file_uploader("📄 DOCX (tabela)", type=["docx", "DOCX"])
pptx_file = st.file_uploader("📊 PPTX (modelo)", type=["pptx", "PPTX"])

st.markdown("---")
st.header("2. Gerar com diagnóstico")
if st.button("🔧 Gerar Apresentação (Diagnóstico)"):
    if not docx_file or not pptx_file:
        st.warning("Envie os dois arquivos.")
    else:
        dados = extrair_dados(docx_file)
        st.subheader("Amostra dos dados extraídos (primeiras 5 equipes):")
        for i, d in enumerate(dados[:5]):
            st.write(f"Equipe #{i+1}:", d)

        st.subheader("Textos do slide modelo (cada shape):")
        prs_template = Presentation(pptx_file)
        for si, shape in enumerate(prs_template.slides[0].shapes):
            if shape.has_text_frame:
                text = "".join(run.text for p in shape.text_frame.paragraphs for run in p.runs)
                st.write(f"Shape {si} text:", repr(text))
            else:
                st.write(f"Shape {si} (no text) type:", getattr(shape, "shape_type", "unknown"))

        st.info("Executando substituições e coletando debug (mostrando as 3 primeiras equipes)...")
        prs_final, debug = gerar_apresentacao_debug(dados, pptx_file, max_preview=3)

        st.subheader("DEBUG por slide (primeiras 3 equipes):")
        for slide_debug in debug:
            st.write("=== TEAM INDEX:", slide_debug["team_index"], " ===")
            st.write("team_data keys:", list(slide_debug["team_data"].keys()))
            st.write("team_data preview:", slide_debug["team_data"])
            for s_idx, s in enumerate(slide_debug["shapes"]):
                st.write(f" shape {s_idx}:")
                st.write("   before:", repr(s.get("shape_before","")))
                st.write("   found_keys:", s.get("found_keys",[]))
                st.write("   after:", repr(s.get("shape_after","")))
        st.success("Diagnóstico concluído. Se algo ainda não aparecer, copie/cole os textos acima aqui.")
        # disponibiliza download do ppt gerado (opcional)
        buf = BytesIO()
        prs_final.save(buf)
        buf.seek(0)
        st.download_button("📥 Baixar PPT (diagnóstico)", data=buf, file_name="diagnostico_output.pptx", mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

