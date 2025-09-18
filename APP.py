import streamlit as st
from docx import Document
from pptx import Presentation
from collections import defaultdict
from io import BytesIO
from copy import deepcopy

# === Funções Utilitárias ===
def formatar_texto(texto, maiusculo_estado=False):
    """Capitaliza nomes e coloca estados em maiúsculo."""
    texto = ' '.join(texto.strip().split())
    return texto.upper() if maiusculo_estado else ' '.join(w.capitalize() for w in texto.split())

def extrair_dados(uploaded_file):
    """
    Lê o DOCX e retorna uma lista de equipes ordenadas pelo 'Valido'.
    Ignora a coluna Medalhas.
    """
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

    # Agrupar por equipe
    equipes = defaultdict(list)
    for r in registros:
        equipes[r["Equipe"]].append(r)

    # Ordenar equipes por lancamentos válidos (crescente)
    def chave_ordem(membros):
        try:
            return float(membros[0]["Valido"].replace(",", "."))
        except:
            return float("inf")

    equipes_ordenadas = sorted(equipes.items(), key=lambda x: chave_ordem(x[1]))

    # Montar dados finais para slides
    dados_finais = []
    for equipe_nome, membros in equipes_ordenadas:
        lider = [m for m in membros if "líder" in m["Funcao"] or "lider" in m["Funcao"]]
        acompanhante = [m for m in membros if "acompanhante" in m["Funcao"]]
        alunos = sorted(
            [m for m in membros if "aluno" in m["Funcao"]],
            key=lambda m: formatar_texto(m["Nome"])
        )

        blocos_nomes = []
        if lider: blocos_nomes.append(formatar_texto(lider[0]["Nome"]))
        if acompanhante: blocos_nomes.append(formatar_texto(acompanhante[0]["Nome"]))
        blocos_nomes.extend(formatar_texto(a["Nome"]) for a in alunos)

        info = membros[0]
        dados_finais.append({
            "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info['Valido']} m",
            "{{NOME_EQUIPE}}": f"Equipe: {equipe_nome.split()[-1]}",
            "{{NOME_ESCOLA}}": formatar_texto(info["Escola"]),
            "{{CIDADE_UF}}": f"{formatar_texto(info['Cidade'])} / {formatar_texto(info['Estado'], True)}",
            "{{NOME_LIDER}}": formatar_texto(lider[0]["Nome"]) if lider else "",
            "{{NOME_ACOMPANHANTE}}": formatar_texto(acompanhante[0]["Nome"]) if acompanhante else "",
            "{{NOMES_ALUNOS}}": "\n".join(blocos_nomes[2:] if lider or acompanhante else blocos_nomes)
        })
    return dados_finais

def substituir_texto(shape, dados_time):
    if not shape.has_text_frame:
        return
    tf = shape.text_frame
    for p in tf.paragraphs:
        texto = "".join(run.text for run in p.runs)
        for chave, valor in dados_time.items():
            texto = texto.replace(chave, valor)
        for _ in p.runs:
            p._p.remove(p.runs[0]._r)
        p.add_run().text = texto

def duplicar_slide(prs, indice):
    template = prs.slides[indice]
    layout = prs.slide_layouts[6] if len(prs.slide_layouts) > 6 else prs.slide_layouts[-1]
    novo = prs.slides.add_slide(layout)
    for s in template.shapes:
        novo.shapes._spTree.insert_element_before(deepcopy(s.element), 'p:extLst')
    return novo

def gerar_apresentacao(dados, template_file):
    prs = Presentation(template_file)
    if not dados or not prs.slides:
        return prs
    # Primeiro slide
    for s in prs.slides[0].shapes:
        substituir_texto(s, dados[0])
    # Restantes
    for d in dados[1:]:
        novo = duplicar_slide(prs, 0)
        for s in novo.shapes:
            substituir_texto(s, d)
    return prs

# === Interface Streamlit ===
st.set_page_config(layout="wide")
st.title("🚀 Gerador Automático de Slides – Versão Final")
st.info("Envie o DOCX com as equipes e o PPTX modelo para gerar os slides.")

docx_file = st.file_uploader("📄 Envie o DOCX de dados", type=["docx"])
pptx_file = st.file_uploader("📊 Envie o PPTX modelo", type=["pptx"])

if st.button("✨ Gerar Apresentação"):
    if docx_file and pptx_file:
        try:
            dados = extrair_dados(docx_file)
            if dados:
                prs = gerar_apresentacao(dados, pptx_file)
                buf = BytesIO()
                prs.save(buf)
                buf.seek(0)
                st.success(f"Gerado {len(dados)} slides!")
                st.download_button(
                    label="📥 Baixar Apresentação Final",
                    data=buf,
                    file_name="Apresentacao_Final_Equipes.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )
            else:
                st.warning("Nenhum dado válido encontrado no DOCX.")
        except Exception as e:
            st.error(f"Erro: {e}")
    else:
        st.warning("Envie os dois arquivos para continuar.")
