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

# Logo e título
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
                continue  # Ignora cabeçalho
            celulas = [c.text.strip() for c in linha.cells]
            if len(celulas) >= 8:
                # Ignora a primeira coluna (MEDALHA)
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

    # Agrupa por equipe
    equipes = defaultdict(list)
    for r in registros:
        equipes[r["Equipe"]].append(r)

    # Ordena por lançamento válido (ordem crescente)
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
        alunos = sorted
