import streamlit as st
from docx import Document
from pptx import Presentation
from collections import defaultdict
from io import BytesIO
import os

# === Funções de Processamento de Dados (do primeiro script, modificadas) ===

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
    
    # 1. Extrair dados da tabela do Word
    if not doc.tables:
        st.error("Nenhuma tabela encontrada no arquivo DOCX.")
        return None

    for tabela in doc.tables:
        for i, linha in enumerate(tabela.rows):
            if i == 0: continue # Ignora cabeçalho
            
            valores = [c.text.strip() for c in linha.cells]
            if len(valores) >= 8:
                medalha, valido, equipe, funcao, escola, cidade, estado, nome = valores[:8]
                dados_brutos.append({
                    "Valido": valido, "Equipe": equipe, "Funcao": funcao.lower(),
                    "Escola": escola, "Cidade": cidade, "Estado": estado, "Nome": nome
                })

    # 2. Agrupar dados por equipe
    equipes = defaultdict(list)
    for item in dados_brutos:
        equipes[item["Equipe"]].append(item)

    # 3. Ordenar equipes pelo lançamento válido
    def valor_valido(membros):
        try:
            return float(membros[0]['Valido'].replace(',', '.'))
        except (ValueError, IndexError):
            return float('inf') # Equipes sem valor válido vão para o final

    equipes_ordenadas = sorted(equipes.items(), key=lambda x: valor_valido(x[1]))

    # 4. Estruturar os dados no formato final para cada equipe
    dados_finais_para_slides = []
    for equipe_nome, membros in equipes_ordenadas:
        lider_obj = [m for m in membros if "líder" in m["Funcao"] or "lider" in m["Funcao"]]
        acompanhante_obj = [m for m in membros if "acompanhante" in m["Funcao"]]
        alunos_obj = sorted([m for m in membros if "aluno" in m["Funcao"]], key=lambda m: formatar_texto(m["Nome"]))

        # Formata os nomes para exibição
        nome_lider = formatar_texto(lider_obj[0]["Nome"]) if lider_obj else ""
        nome_acompanhante = formatar_texto(acompanhante_obj[0]["Nome"]) if acompanhante_obj else ""
        nomes_alunos = "\n".join([formatar_texto(m["Nome"]) for m in alunos_obj])

        if membros:
            info_geral = membros[0]
            equipe_formatada = f"Equipe: {equipe_nome.split()[-1]}"
            lancamento_formatado = f"ALCANCE: {info_geral['Valido']} m"
            escola_formatada = formatar_texto(info_geral["Escola"])
            cidade_uf_formatada = f"{formatar_texto(info_geral['Cidade'])} / {formatar_texto(info_geral['Estado'], maiusculo_estado=True)}"

            dados_finais_para_slides.append({
                "Líder": nome_lider,
                "Acompanhante": nome_acompanhante,
                "Alunos": nomes_alunos,
                "Equipe": equipe_formatada,
                "Lançamentos Válidos": lancamento_formatado,
                "Nome da Escola": escola_formatada,
                "Cidade / UF": cidade_uf_formatada
            })

    return dados_finais_para_slides

# === Função de Geração de PowerPoint (do segundo script) ===

def generate_presentation(team_data, template_file, placeholder_map):
    """Gera a apresentação de slides a partir dos dados da equipe e do modelo."""
    prs = Presentation(template_file)
    slide_layout = prs.slide_layouts[0] # Usar o primeiro layout como padrão

    for team in team_data:
        slide = prs.slides.add_slide(slide_layout)

        for shape in slide.shapes:
            # Checa se o nome do shape (placeholder) está no nosso mapeamento
            if shape.name in placeholder_map:
                # Pega a chave dos nossos dados (ex: "Líder", "Alunos")
                data_key = placeholder_map[shape.name]
                # Pega o texto correspondente para a equipe atual
                text_to_insert = team.get(data_key, "")
                
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    text_frame.clear()
                    p = text_frame.paragraphs[0]
                    run = p.add_run()
                    run.text = text_to_insert

    return prs

# === Interface Streamlit Unificada ===

st.set_page_config(layout="wide")
st.title("🚀 Gerador de Slides para Jornada de Foguetes")

st.info("Este aplicativo lê uma tabela de dados de um arquivo `.docx`, processa as equipes e gera uma apresentação de slides `.pptx` a partir de um modelo.")

col1, col2 = st.columns(2)

with col1:
    st.header("1. Carregue os Arquivos")
    uploaded_data_file = st.file_uploader("Arquivo .docx com a TABELA DE DADOS", type=["docx"])
    uploaded_template_file = st.file_uploader("Arquivo .pptx com o MODELO DE SLIDE", type=["pptx"])

with col2:
    st.header("2. Mapeie os Campos do Slide")
    st.write("Preencha com os nomes dos 'Placeholders' do seu slide modelo. (Encontre em `Página Inicial > Organizar > Painel de Seleção` no PowerPoint).")
    
    leader_placeholder = st.text_input("Placeholder para o Líder", "NomeLider")
    accompanist_placeholder = st.text_input("Placeholder para o Acompanhante", "NomeAcompanhante")
    students_placeholder = st.text_input("Placeholder para os Alunos", "NomesAlunos")
    team_name_placeholder = st.text_input("Placeholder para a Equipe", "NomeEquipe")
    launches_placeholder = st.text_input("Placeholder para o Alcance", "LancamentosValidos")
    school_name_placeholder = st.text_input("Placeholder para a Escola", "NomeEscola")
    city_state_placeholder = st.text_input("Placeholder para Cidade / UF", "CidadeUF")

st.divider()

st.header("3. Gere a Apresentação")
if st.button("✨ Gerar Slides!", use_container_width=True):
    if uploaded_data_file and uploaded_template_file:
        with st.spinner("Processando dados e criando apresentação..."):
            try:
                # Mapeamento dos placeholders para as chaves de dados
                placeholder_mapping = {
                    leader_placeholder: "Líder",
                    accompanist_placeholder: "Acompanhante",
                    students_placeholder: "Alunos",
                    team_name_placeholder: "Equipe",
                    launches_placeholder: "Lançamentos Válidos",
                    school_name_placeholder: "Nome da Escola",
                    city_state_placeholder: "Cidade / UF"
                }
                # Filtra mapeamentos com chaves vazias
                placeholder_mapping = {k: v for k, v in placeholder_mapping.items() if k}

                # Passo 1: Extrair e estruturar os dados do .docx
                teams_data = extrair_e_estruturar_dados(uploaded_data_file)
                
                if teams_data:
                    # Passo 2: Gerar a apresentação com os dados estruturados
                    presentation = generate_presentation(teams_data, uploaded_template_file, placeholder_mapping)

                    # Salvar em memória para download
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
                st.error("Dica: Verifique se a tabela no arquivo .docx está correta e se os nomes dos placeholders correspondem exatamente aos do PowerPoint.")
    else:
        st.warning("Por favor, carregue o arquivo de dados e o modelo de PowerPoint.")

if 'generation_complete' in st.session_state and st.session_state.generation_complete:
    st.download_button(
        label="📥 Baixar Apresentação Final",
        data=st.session_state.pptx_buffer,
        file_name="Apresentacao_Final_Equipes.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        use_container_width=True
    )
