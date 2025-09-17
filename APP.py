import streamlit as st
from docx import Document
from pptx import Presentation
from collections import defaultdict
from io import BytesIO

# === Funções de Processamento de Dados (sem alterações) ===
def formatar_texto(texto, maiusculo_estado=False):
    texto = ' '.join(texto.strip().split())
    if maiusculo_estado:
        return texto.upper()
    return ' '.join(w.capitalize() for w in texto.split())

def extrair_e_estruturar_dados(uploaded_file):
    doc = Document(uploaded_file)
    dados_brutos = []
    
    if not doc.tables:
        st.error("Nenhuma tabela encontrada no arquivo DOCX.")
        return None

    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            if i == 0: continue
            
            valores = [c.text.strip() for c in linha.cells]
            if len(valores) >= 8:
                medalha, valido, equipe, funcao, escola, cidade, estado, nome = valores[:8]
                dados_brutos.append({
                    "Medalha": medalha, "Valido": valido, "Equipe": equipe, "Funcao": funcao.lower(),
                    "Escola": escola, "Cidade": cidade, "Estado": estado, "Nome": nome
                })

    equipes = defaultdict(list)
    for item in dados_brutos:
        equipes[item["Equipe"]].append(item)

    def valor_valido(membros):
        try:
            return float(membros[0]['Valido'].replace(',', '.'))
        except (ValueError, IndexError):
            return float('inf')

    equipes_ordenadas = sorted(equipes.items(), key=lambda x: valor_valido(x[1]))

    dados_finais_para_slides = []
    for equipe_nome, membros in equipes_ordenadas:
        lider_obj = [m for m in membros if "líder" in m["Funcao"] or "lider" in m["Funcao"]]
        acompanhante_obj = [m for m in membros if "acompanhante" in m["Funcao"]]
        alunos_obj = sorted([m for m in membros if "aluno" in m["Funcao"]], key=lambda m: formatar_texto(m["Nome"]))

        nome_lider = formatar_texto(lider_obj[0]["Nome"]) if lider_obj else ""
        nome_acompanhante = formatar_texto(acompanhante_obj[0]["Nome"]) if acompanhante_obj else ""
        nomes_alunos = "\n".join([formatar_texto(m["Nome"]) for m in alunos_obj])

        if membros:
            info_geral = membros[0]
            dados_finais_para_slides.append({
                "{{NOME_LIDER}}": nome_lider,
                "{{NOME_ACOMPANHANTE}}": nome_acompanhante,
                "{{NOMES_ALUNOS}}": nomes_alunos,
                "{{NOME_EQUIPE}}": f"Equipe: {equipe_nome.split()[-1]}",
                "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info_geral['Valido']} m",
                "{{NOME_ESCOLA}}": formatar_texto(info_geral["Escola"]),
                "{{CIDADE_UF}}": f"{formatar_texto(info_geral['Cidade'])} / {formatar_texto(info_geral['Estado'], maiusculo_estado=True)}",
                "{{TITULO_MEDALHA}}": formatar_texto(info_geral["Medalha"]).upper()
            })
    return dados_finais_para_slides

# === Função de Geração de PowerPoint (ATUALIZADA PARA LER TAGS) ===

def generate_presentation(team_data, template_file):
    prs = Presentation(template_file)
    
    # Pega o layout do primeiro slide como base para os novos slides
    slide_layout = prs.slide_layouts[0] 

    for team in team_data:
        slide = prs.slides.add_slide(slide_layout)

        # Itera sobre todas as formas do slide para encontrar e substituir as tags
        for shape in slide.shapes:
            if not shape.has_text_frame:
                continue
            
            # Itera sobre todas as chaves (tags) que precisamos substituir
            for key, value in team.items():
                if key in shape.text:
                    text_frame = shape.text_frame
                    for paragraph in text_frame.paragraphs:
                        for run in paragraph.runs:
                            # Substitui a tag pelo valor
                            run.text = run.text.replace(key, value)
    return prs

# === Interface Streamlit (Simplificada e Final) ===

st.set_page_config(layout="wide")
st.title("🚀 Gerador Automático de Slides")
st.info("Faça o upload da tabela de dados e do modelo de PowerPoint para gerar a apresentação final.")

st.header("1. Carregue os Arquivos")
st.write("Certifique-se que o arquivo de modelo `.pptx` (baixado do Google Slides) contém as tags de texto, como `{{NOME_LIDER}}`.")

uploaded_data_file = st.file_uploader("Arquivo .docx com a TABELA DE DADOS", type=["docx"])
uploaded_template_file = st.file_uploader("Arquivo .pptx com o MODELO DE SLIDE", type=["pptx"])

st.divider()

st.header("2. Gere a Apresentação")
if st.button("✨ Gerar Slides!", use_container_width=True):
    if uploaded_data_file and uploaded_template_file:
        with st.spinner("Mapeando Foguetes... 🚀 Processando dados e criando apresentação..."):
            try:
                teams_data = extrair_e_estruturar_dados(uploaded_data_file)
                
                if teams_data:
                    presentation = generate_presentation(teams_data, uploaded_template_file)

                    pptx_buffer = BytesIO()
                    presentation.save(pptx_buffer)
                    pptx_buffer.seek(0)
                    
                    st.session_state.pptx_buffer = pptx_buffer
                    st.session_state.generation_complete = True
                    st.success(f"Apresentação com {len(teams_data)} slides gerada com sucesso!")
                else:
                    st.warning("Não foi possível gerar os slides. Verifique o arquivo de dados.")

            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
                st.error("Dica: Verifique se a tabela no arquivo .docx está correta e se as tags no modelo .pptx estão escritas corretamente (ex: `{{NOME_EQUIPE}}`).")
    else:
        st.warning("Por favor, carregue os dois arquivos.")

if 'generation_complete' in st.session_state and st.session_state.generation_complete:
    st.download_button(
        label="📥 Baixar Apresentação Final",
        data=st.session_state.pptx_buffer,
        file_name="Apresentacao_Final_Equipes.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
