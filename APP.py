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

# -------------------- HELPERS --------------------
def formatar_texto(texto, maiusculo_estado=False):
    texto = ' '.join(texto.strip().split())
    return texto.upper() if maiusculo_estado else ' '.join(w.capitalize() for w in texto.split())

# Map de nomes esperados das colunas (versões possíveis)
EXPECTED_HEADERS = {
    "medalha": ["medalha"],
    "valido": ["valido", "validação", "lançamentos válidos", "lançamentos_validos"],
    "equipe": ["equipe"],
    "funcao": ["função", "funcao"],
    "escola": ["nome da escola", "escola", "nome_escola"],
    "cidade": ["cidade da escola", "cidade", "cidade_escola"],
    "estado": ["estado da escola", "estado", "uf", "estado_escola"],
    "nome_participante": ["nome_participante", "nome", "participante"]
}

def find_header_indexes(header_row_cells):
    """
    Recebe uma lista de strings (texto do cabeçalho) e retorna dicionário com índices para campos esperados.
    Se não conseguir mapear todos, retorna None para os que faltarem.
    """
    headers = [h.strip().lower() for h in header_row_cells]
    idx_map = {}
    for key, variants in EXPECTED_HEADERS.items():
        found = None
        for i, h in enumerate(headers):
            for v in variants:
                if h == v:
                    found = i
                    break
            if found is not None:
                break
        # tentativa mais frouxa: contains
        if found is None:
            for i, h in enumerate(headers):
                for v in variants:
                    if v in h:
                        found = i
                        break
                if found is not None:
                    break
        idx_map[key] = found
    return idx_map

# -------------------- EXTRAIR DADOS (robusto) --------------------
def extrair_dados(uploaded_file):
    doc = Document(uploaded_file)
    registros = []

    # tenta encontrar a primeira tabela com dados
    if not doc.tables:
        st.error("Nenhuma tabela encontrada no DOCX.")
        return []

    tabela = doc.tables[0]
    # pega cabeçalho (primeira linha) para mapear colunas
    header_cells = [c.text.strip() for c in tabela.rows[0].cells]
    idx_map = find_header_indexes(header_cells)

    # se houver índices None, vamos usar fallback posicional (antiga ordem),
    # mas avisamos o usuário.
    missing = [k for k, v in idx_map.items() if v is None and k != "medalha"]
    if missing:
        st.warning(
            "Aviso: o cabeçalho da tabela não corresponde exatamente ao esperado. "
            "Vou tentar ler pela ordem padrão; se os dados saírem errados, reorganize as colunas conforme o modelo."
        )
        # fallback: assume ordem antiga [MEDALHA, VALIDO, EQUIPE, FUNÇÃO, Nome da Escola, Cidade da Escola, Estado da Escola, Nome_Participante]
        for i, row in enumerate(tabela.rows):
            if i == 0:
                continue
            celulas = [c.text.strip() for c in row.cells]
            if len(celulas) >= 8:
                # ignora medalha
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
    else:
        # usa índices mapeados
        for i, row in enumerate(tabela.rows):
            if i == 0:
                continue
            cells = [c.text.strip() for c in row.cells]
            # segurança: preencher com vazio se índice não existir
            def g(k):
                idx = idx_map[k]
                try:
                    return cells[idx].strip()
                except Exception:
                    return ""
            valido = g("valido")
            equipe = g("equipe")
            funcao = g("funcao")
            escola = g("escola")
            cidade = g("cidade")
            estado = g("estado")
            nome = g("nome_participante")

            # só adiciona se equipe existir (evita linhas vazias)
            if equipe:
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
        # Monta nome da equipe e escola/cidade com quebras de linha
        nome_equipe_formatado = f"Equipe: {equipe_nome.split()[-1]}"
        nome_escola_formatado = formatar_texto(info.get("Escola", ""))
        cidade_uf_formatado = f"{formatar_texto(info.get('Cidade', ''))} / {formatar_texto(info.get('Estado', ''), True)}"

        dados_finais.append({
            dados_finais.append({
            "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info['Valido']} m",
            "{{NOME_EQUIPE}}": f"Equipe: {equipe_nome.split()[-1]}",
            "{{NOME_ESCOLA}}": f"{formatar_texto(info['Escola'])}\n{formatar_texto(info['Cidade'])} / {formatar_texto(info['Estado'], True)}",
            "{{NOMES_ALUNOS}}": nomes_formatados
            })
         })
    return dados_finais

# -------------------- PPTX helpers --------------------
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

    for paragraph in shape.text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        selected_key = None

        # Verifica qual placeholder está presente
        for k in team_data.keys():
            if k in full_text:
                selected_key = k
                break

        if not selected_key:
            continue

        new_text = full_text.replace(selected_key, team_data[selected_key])

        # Limpa o conteúdo anterior
        while paragraph.runs:
            paragraph._p.remove(paragraph.runs[0]._r)

        # --- Estilos específicos ---
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
            tf = shape.text_frame
            tf.clear()
            linhas = new_text.split("\n")

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

        elif selected_key == "{{NOME_EQUIPE}}":
            run = paragraph.add_run()
            run.text = new_text
            run.font.name = "Lexend"
            run.font.bold = True
            run.font.size = Pt(20)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        elif selected_key in ("{{NOME_ESCOLA}}", "{{CIDADE_UF}}"):
            tf = shape.text_frame
            tf.clear()

            # Divide escola e cidade/UF em linhas separadas
            escola = team_data.get("{{NOME_ESCOLA}}", "")
            cidade_uf = team_data.get("{{CIDADE_UF}}", "")

            p1 = tf.paragraphs[0]
            r1 = p1.add_run()
            r1.text = escola
            r1.font.name = "Lexend"
            r1.font.bold = True
            r1.font.size = Pt(20)
            r1.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p1.alignment = PP_ALIGN.CENTER

            p2 = tf.add_paragraph()
            r2 = p2.add_run()
            r2.text = cidade_uf
            r2.font.name = "Lexend"
            r2.font.bold = True
            r2.font.size = Pt(20)
            r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p2.alignment = PP_ALIGN.CENTER
            p2.line_spacing = None

        else:
            run = paragraph.add_run()
            run.text = new_text
            run.font.name = "Lexend"
            run.font.bold = True
            run.font.size = Pt(18)
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

        paragraph.alignment = PP_ALIGN.CENTER
        paragraph.line_spacing = None


def gerar_apresentacao(dados, template_stream):
    prs = Presentation(template_stream)
    if not dados or not prs.slides:
        return prs

    modelo = prs.slides[0]

    # Duplica os slides antes de preencher
    for _ in range(len(dados) - 1):
        duplicate_slide_with_media(prs, modelo)

    # Preenche cada slide
    for slide, team_data in zip(prs.slides, dados):
        if isinstance(team_data, dict):  # ✅ Evita erro de tipo
            for shape in slide.shapes:
                replace_placeholders_in_shape(shape, team_data)

    return prs


# -------------------- INTERFACE STREAMLIT --------------------
docx_file = st.file_uploader("📄 Arquivo DOCX (tabela)", type=["docx", "DOCX"])
pptx_file = st.file_uploader("📊 Arquivo PPTX modelo (1 slide com placeholders)", type=["pptx", "PPTX"])

if st.button("✨ Gerar Apresentação"):
    if not docx_file or not pptx_file:
        st.warning("Envie os dois arquivos (.docx e .pptx).")
    else:
        try:
            dados = extrair_dados(docx_file)
            if not dados:
                st.warning("Nenhum dado encontrado no DOCX.")
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






