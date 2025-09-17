import streamlit as st
from docx import Document
from pptx import Presentation
from collections import defaultdict
from io import BytesIO

# === Funções de Processamento de Dados (CORRIGIDA) ===

def formatar_texto(texto, maiusculo_estado=False):
    """Formata uma string, capitalizando palavras e tratando o estado."""
    texto = ' '.join(texto.strip().split())
    if maiusculo_estado:
        return texto.upper()
    return ' '.join(w.capitalize() for w in texto.split())

def extrair_e_estruturar_dados(uploaded_file):
    """
    Lê a tabela de um arquivo .docx e retorna uma lista de dicionários,
    um para cada equipe, com os dados já processados e prontos para o slide.
    """
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

        # --- ALTERAÇÃO PRINCIPAL AQUI ---
        # Agora, cada função tem sua própria variável.
        nome_lider = formatar_texto(lider_obj[0]["Nome"]) if lider_obj else ""
        nome_acompanhante = formatar_texto(acompanhante_obj[0]["Nome"]) if acompanhante_obj else ""
        nomes_alunos = "\n".join([formatar_texto(m["Nome"]) for m in alunos_obj])

        if membros:
            info_geral = membros[0]
            # O dicionário final agora tem campos separados para cada função
            dados_finais_para_slides.append({
                "NomeLider": nome_lider,
                "NomeAcompanhante": nome_acompanhante,
                "NomesAlunos": nomes_alunos,
                "NomeEquipe": f"Equipe: {equipe_nome.split()[-1]}",
                "LancamentosValidos": f"ALCANCE: {info_geral['Valido']} m",
                "NomeEscola": formatar_texto(info_geral["Escola"]),
                "CidadeUF": f"{formatar_texto(info_geral['Cidade'])} / {formatar_texto(info_geral['Estado'], maiusculo_estado=True)}",
                "TituloMedalha": formatar_texto(info_geral["Medalha"]).upper()
            })

    return dados_finais_para_slides

# === Função de Geração de PowerPoint (sem alterações) ===

def generate_presentation(team_data, template_file, placeholder_map):
    prs = Presentation(template_file)
    slide_layout = prs.slide_layouts[0]

    for team in team_data:
        slide = prs.slides.add_slide(slide_layout)
        for shape in slide.shapes:
            if shape.name in placeholder_map:
                data_key = placeholder_map[shape.name]
                text_to_insert = team.get(data_key, "")
                
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    text_frame.clear()
                    p = text_frame.paragraphs[0]
                    run = p.add_run()
                    run.text = text_to_insert
    return prs

# === Interface Streamlit ===

st.set_page_config(layout="wide")
st.title("🚀 Gerador Automático de Slides")
st.info("Faça o upload da tabela de dados e do modelo de PowerPoint para gerar a apresentação final.")

# --- Mapeamento Fixo de Placeholders (ATUALIZADO) ---
# O programa agora espera que os shapes no seu PPT tenham EXATAMENTE estes nomes.
PLACEHOLDER_MAP_FIXO = {
    "NomeLider": "NomeLider",
    "NomeAcompanhante": "NomeAcompanhante",
    "NomesAlunos": "NomesAlunos",
    "NomeEquipe": "NomeEquipe",
    "NomeEscola": "NomeEscola",
    "CidadeUF": "CidadeUF",
    "LancamentosValidos": "LancamentosValidos",
    "TituloMedalha": "TituloMedalha"
}

st.header("1. Carregue os Arquivos")
st.write("Certifique-se que o arquivo de modelo `.pptx` já está com os placeholders nomeados corretamente conforme a convenção.")

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
                    presentation = generate_presentation(teams_data, uploaded_template_file, PLACEHOLDER_MAP_FIXO)

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
                st.error("Dica: Verifique se a tabela no arquivo .docx está correta e se os nomes dos placeholders no PowerPoint correspondem EXATAMENTE à convenção.")
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
