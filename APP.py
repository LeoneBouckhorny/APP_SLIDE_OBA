import streamlit as st
import pandas as pd
from pptx import Presentation
from pptx.util import Pt
from io import BytesIO
import docx # Importa a nova biblioteca

def parse_team_data(uploaded_file):
    """
    Analisa o arquivo .docx com as informações das equipes.
    Espera-se que as equipes sejam separadas por uma linha em branco.
    """
    document = docx.Document(uploaded_file)
    parsed_teams = []
    current_team_block = []

    for para in document.paragraphs:
        line_text = para.text.strip()
        
        if not line_text: # Se a linha está em branco, é um separador de equipe
            if current_team_block: # Processa o bloco da equipe anterior
                team_info = process_block(current_team_block)
                parsed_teams.append(team_info)
                current_team_block = [] # Limpa para a próxima equipe
        else:
            current_team_block.append(line_text)
    
    # Processa a última equipe do arquivo (que não terá uma linha em branco depois)
    if current_team_block:
        team_info = process_block(current_team_block)
        parsed_teams.append(team_info)

    return parsed_teams

def process_block(lines):
    """Função auxiliar para processar as linhas de um único bloco de equipe."""
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
         # Heurística para diferenciar Acompanhante de uma lista de Alunos
         if "Acompanhante" in lines[0] or "Líder" in lines[0] or len(participants) > 2:
             team_info['Acompanhante'] = participants[1]
             team_info['Alunos'] = "\n".join(participants[2:]) if len(participants) > 2 else ""
         else:
             team_info['Acompanhante'] = ""
             team_info['Alunos'] = "\n".join(participants[1:])
    else:
        team_info['Acompanhante'] = ""
        team_info['Alunos'] = "\n".join(participants[1:]) if len(participants) > 1 else ""
    
    # Extrai outras informações
    for line in lines[participant_end_index:]:
        if line.startswith("Equipe:"):
            team_info['Equipe'] = line.split(":", 1)[1].strip()
        elif line.startswith("Lançamentos Válidos:"):
            team_info['Lançamentos Válidos'] = line.split(":", 1)[1].strip()
        elif "/" in line and len(line.split('/')) == 2:
             cidade, estado = [x.strip() for x in line.split('/')]
             team_info['Cidade'] = cidade
             team_info['Estado'] = estado
        else: # Nome da Escola
            if 'Nome da Escola' not in team_info:
                team_info['Nome da Escola'] = line

    return team_info


def generate_presentation(team_data, template_file, placeholder_map):
    """
    Gera a apresentação de slides a partir dos dados da equipe e do modelo.
    """
    prs = Presentation(template_file)
    
    # Seleciona o layout do primeiro slide do template como base
    slide_layout = prs.slide_layouts[0] 

    for team in team_data:
        slide = prs.slides.add_slide(slide_layout)

        for shape in slide.placeholders:
            if shape.name in placeholder_map:
                key_map = placeholder_map[shape.name]
                text_to_insert = team.get(key_map, "") # Usa .get para evitar erros se uma chave não existir
                
                # Preenche o texto no placeholder
                text_frame = shape.text_frame
                text_frame.clear() # Limpa qualquer texto padrão
                p = text_frame.paragraphs[0]
                run = p.add_run()
                run.text = text_to_insert

        # Você pode também procurar por shapes que não são placeholders, se necessário
        for shape in slide.shapes:
            if shape.name in placeholder_map:
                key_map = placeholder_map[shape.name]
                text_to_insert = team.get(key_map, "")
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    text_frame.clear()
                    p = text_frame.paragraphs[0]
                    run = p.add_run()
                    run.text = text_to_insert

    return prs

# --- Interface do Streamlit ---

st.title("Gerador de Slides para Equipes")

st.header("1. Carregue os Arquivos")
# ATUALIZADO para aceitar .docx
uploaded_data = st.file_uploader("Escolha o arquivo Word com a lista de equipes (.docx)", type=["docx"])
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
# Filtra entradas vazias do usuário para não causar erros
placeholder_mapping = {k: v for k, v in placeholder_mapping.items() if k}


st.header("3. Gere e Baixe a Apresentação")
if st.button("Gerar Apresentação"):
    if uploaded_data is not None and uploaded_template is not None:
        with st.spinner("Processando..."):
            try:
                teams = parse_team_data(uploaded_data)
                
                if not teams:
                     st.error("Nenhuma equipe encontrada no arquivo. Verifique se o formato está correto (equipes separadas por uma linha em branco).")
                else:
                    presentation = generate_presentation(teams, uploaded_template, placeholder_mapping)

                    # Salvar apresentação em memória
                    pptx_buffer = BytesIO()
                    presentation.save(pptx_buffer)
                    pptx_buffer.seek(0)
                    
                    st.session_state.pptx_buffer = pptx_buffer
                    st.session_state.generation_complete = True

                    st.success(f"Apresentação com {len(teams)} slides gerada com sucesso!")

            except Exception as e:
                st.error(f"Ocorreu um erro ao processar os arquivos: {e}")
                st.error("Verifique se os nomes dos placeholders estão corretos e se os arquivos não estão corrompidos.")
    else:
        st.warning("Por favor, carregue o arquivo de dados e o modelo de PowerPoint.")

if 'generation_complete' in st.session_state and st.session_state.generation_complete:
    st.download_button(
        label="Baixar Apresentação",
        data=st.session_state.pptx_buffer,
        file_name="apresentacao_equipes.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
