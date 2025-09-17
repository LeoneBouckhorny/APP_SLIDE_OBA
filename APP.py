import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO

def parse_team_data(uploaded_file):
    """
    Analisa o arquivo de texto com as informações das equipes.
    Espera-se que o arquivo siga o formato especificado no prompt.
    """
    content = uploaded_file.getvalue().decode("utf-8")
    teams = content.strip().split('\n\n')
    parsed_teams = []

    for team_block in teams:
        lines = team_block.strip().split('\n')
        team_info = {}
        
        # Extrai nomes de participantes
        participants = []
        for i, line in enumerate(lines):
            if line.startswith("Equipe:"):
                participant_end_index = i
                break
            participants.append(line)
        else:
            participant_end_index = len(lines)

        # Assumindo a ordem: Líder, Acompanhante (opcional), Alunos
        team_info['Líder'] = participants[0] if len(participants) > 0 else ""
        
        if len(participants) > 1 and not lines[1].startswith("Equipe:"):
             if "Aluno" in lines[1]: # Heurística simples
                 team_info['Acompanhante'] = ""
                 team_info['Alunos'] = "\n".join(participants[1:])
             else:
                 team_info['Acompanhante'] = participants[1]
                 if len(participants) > 2:
                    team_info['Alunos'] = "\n".join(participants[2:])
                 else:
                    team_info['Alunos'] = ""
        else:
            team_info['Acompanhante'] = ""
            if len(participants) > 1:
                team_info['Alunos'] = "\n".join(participants[1:])
            else:
                 team_info['Alunos'] = ""
        
        # Extrai outras informações
        for line in lines[participant_end_index:]:
            if line.startswith("Equipe:"):
                team_info['Equipe'] = line.split(":")[1].strip()
            elif line.startswith("Lançamentos Válidos:"):
                team_info['Lançamentos Válidos'] = line.split(":")[1].strip()
            elif "/" in line and len(line.split('/')) == 2:
                 cidade, estado = [x.strip() for x in line.split('/')]
                 team_info['Cidade'] = cidade
                 team_info['Estado'] = estado
            else: # Nome da Escola
                if 'Nome da Escola' not in team_info:
                    # A linha restante antes da cidade/estado é a escola
                    team_info['Nome da Escola'] = line

        parsed_teams.append(team_info)

    return parsed_teams

def generate_presentation(team_data, template_file, placeholder_map):
    """
    Gera a apresentação de slides a partir dos dados da equipe e do modelo.
    """
    prs = Presentation(template_file)
    slide_layout = prs.slide_layouts[0] # Assumindo o primeiro layout

    for team in team_data:
        slide = prs.slides.add_slide(slide_layout)

        for shape in slide.placeholders:
            if shape.name in placeholder_map:
                key_map = placeholder_map[shape.name]
                text_to_insert = team.get(key_map, "")
                shape.text = text_to_insert

    return prs

# --- Interface do Streamlit ---

st.title("Gerador de Slides para Equipes")

st.header("1. Carregue os Arquivos")
uploaded_data = st.file_uploader("Escolha o arquivo com a lista de equipes (.txt ou .docx)", type="txt, docx")
uploaded_template = st.file_uploader("Escolha o modelo de PowerPoint (.pptx)", type="pptx")

st.header("2. Mapeie os Campos do Slide")
st.write("Preencha com os nomes dos 'Placeholders' do seu slide modelo. Para encontrar os nomes, em seu PowerPoint, vá em `Página Inicial` > `Organizar` > `Painel de Seleção`.")

col1, col2 = st.columns(2)

with col1:
    leader_placeholder = st.text_input("Placeholder para o Líder", "NomeLider")
    accompanist_placeholder = st.text_input("Placeholder para o Acompanhante", "NomeAcompanhante")
    students_placeholder = st.text_input("Placeholder para os Alunos", "NomesAlunos")
    team_name_placeholder = st.text_input("Placeholder para o Nome da Equipe", "NomeEquipe")

with col2:
    launches_placeholder = st.text_input("Placeholder para Lançamentos Válidos", "LancamentosValidos")
    school_name_placeholder = st.text_input("Placeholder para o Nome da Escola", "NomeEscola")
    city_placeholder = st.text_input("Placeholder para a Cidade", "NomeCidade")
    state_placeholder = st.text_input("Placeholder para o Estado", "SiglaEstado")

placeholder_mapping = {
    leader_placeholder: "Líder",
    accompanist_placeholder: "Acompanhante",
    students_placeholder: "Alunos",
    team_name_placeholder: "Equipe",
    launches_placeholder: "Lançamentos Válidos",
    school_name_placeholder: "Nome da Escola",
    city_placeholder: "Cidade",
    state_placeholder: "Estado"
}

st.header("3. Gere e Baixe a Apresentação")
if st.button("Gerar Apresentação"):
    if uploaded_data is not None and uploaded_template is not None:
        with st.spinner("Processando..."):
            try:
                teams = parse_team_data(uploaded_data)
                
                presentation = generate_presentation(teams, uploaded_template, placeholder_mapping)

                # Salvar apresentação em memória
                pptx_buffer = BytesIO()
                presentation.save(pptx_buffer)
                pptx_buffer.seek(0)
                
                st.session_state.pptx_buffer = pptx_buffer
                st.session_state.generation_complete = True

                st.success("Apresentação gerada com sucesso!")

            except Exception as e:
                st.error(f"Ocorreu um erro: {e}")
    else:
        st.warning("Por favor, carregue o arquivo de dados e o modelo de PowerPoint.")

if 'generation_complete' in st.session_state and st.session_state.generation_complete:
    st.download_button(
        label="Baixar Apresentação",
        data=st.session_state.pptx_buffer,
        file_name="apresentacao_equipes.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"

    )
