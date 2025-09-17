import streamlit as st
from docx import Document
from pptx import Presentation
from collections import defaultdict
from io import BytesIO
from copy import deepcopy

# === Funções de Processamento de Dados (sem o campo "Medalha") ===
def formatar_texto(texto, maiusculo_estado=False):
    """Formata uma string, capitalizando palavras e tratando o estado."""
    texto = ' '.join(texto.strip().split())
    if maiusculo_estado:
        return texto.upper()
    return ' '.join(w.capitalize() for w in texto.split())

def extrair_e_estruturar_dados(uploaded_file):
    """Lê a tabela de um arquivo .docx e retorna uma lista de dicionários para cada equipe."""
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
                # O campo "medalha" é lido mas não será usado
                medalha, valido, equipe, funcao, escola, cidade, estado, nome = valores[:8]
                dados_brutos.append({
                    "Valido": valido, "Equipe": equipe, "Funcao": funcao.lower(),
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
            # Dicionário final sem a chave da medalha
            dados_finais_para_slides.append({
                "{{NOME_LIDER}}": nome_lider,
                "{{NOME_ACOMPANHANTE}}": nome_acompanhante,
                "{{NOMES_ALUNOS}}": nomes_alunos,
                "{{NOME_EQUIPE}}": f"Equipe: {equipe_nome.split()[-1]}",
                "{{LANCAMENTOS_VALIDOS}}": f"ALCANCE: {info_geral['Valido']} m",
                "{{NOME_ESCOLA}}": formatar_texto(info_geral["Escola"]),
                "{{CIDADE_UF}}": f"{formatar_texto(info_geral['Cidade'])} / {formatar_texto(info_geral['Estado'], maiusculo_estado=True)}",
            })
    return dados_finais_para_slides

# === Funções de Geração de PowerPoint (Lógica Estável) ===

def replace_text_in_shape(shape, team_data):
    """Substitui as tags de texto em uma forma específica."""
    if not shape.has_text_frame:
        return

    text_frame = shape.text_frame
    for paragraph in text_frame.paragraphs:
        full_text = "".join(run.text for run in paragraph.runs)
        
        for key, value in team_data.items():
            if key in full_text:
                full_text = full_text.replace(key, value)
        
        # Limpa os 'runs' antigos e adiciona um novo com o texto completo
        while len(paragraph.runs) > 0:
            p = paragraph._p
            p.remove(paragraph.runs[0]._r)
        
        paragraph.add_run().text = full_text

def duplicate_slide(prs, index):
    """Duplica um slide e o adiciona no final da apresentação."""
    template = prs.slides[index]
    try:
        blank_slide_layout = prs.slide_layouts[6]
    except IndexError:
        blank_slide_layout = prs.slide_layouts[len(prs.slide_layouts) - 1]

    copied_slide = prs.slides.add_slide(blank_slide_layout)

    for shape in template.shapes:
        new_el = deepcopy(shape.element)
        copied_slide.shapes._spTree.insert_element_before(new_el, 'p:extLst')
    
    return copied_slide

def generate_presentation(team_data, template_file):
    """Gera a apresentação final."""
    prs = Presentation(template_file)
    
    if not team_data or not prs.slides:
        return prs # Retorna a apresentação vazia se não houver dados ou slides

    # Modifica o primeiro slide para a primeira equipe
    first_slide = prs.slides[0]
    first_team = team_data[0]
    for shape in first_slide.shapes:
        replace_text_in_shape(shape, first_team)

    # Para as equipes restantes, duplica o primeiro slide e modifica a cópia
    for i in range(1, len(team_data)):
        team = team_data[i]
        new_slide = duplicate_slide(prs, 0)
        for shape in new_slide.shapes:
            replace_text_in_shape(shape, team)
            
    return prs

# === Interface Streamlit ===

st.set_page_config(layout="wide")
st.title("🚀 Gerador Automático de Slides")
st.info("Faça o upload da tabela de dados e do modelo de PowerPoint para gerar a apresentação final.")

st.header("1. Carregue os Arquivos")
st.write("Certifique-se que o arquivo de modelo `.pptx` contém **apenas um slide** com as tags de texto, como `{{NOME_LIDER}}`.")

uploaded_data_file = st.file_uploader("Arquivo .docx com a TABELA DE DADOS", type=["docx"])
uploaded_template_file = st.file_uploader("Arquivo .pptx com o MODELO DE SLIDE", type=["pptx"])

st.divider()

st.header("2. Gere a Apresentação")
if st.button("✨ Gerar Slides!", use_container_width=True):
    if uploaded_data_file and uploaded_template_file:
        with st.spinner("Construindo Foguetes... 🚀 Processando dados e montando a apresentação..."):
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
                st.error("Dica: Verifique se o seu .pptx tem apenas um slide e se as tags estão escritas corretamente (ex: `{{NOME_EQUIPE}}`).")
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
