import streamlit as st
from docx import Document
from pptx import Presentation
from collections import defaultdict
from io import BytesIO
import zipfile
import re
from copy import deepcopy

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

# === FUNÇÃO DE GERAÇÃO DE SLIDES (CORRIGIDA) ===

def merge_presentations(template_bytes, all_teams_data):
    """
    Cria uma apresentação final juntando cópias do slide modelo,
    uma para cada equipe, e substituindo as tags no nível do XML.
    """
    final_pres_stream = BytesIO()

    # --- CORREÇÃO AQUI ---
    # Cria a apresentação final vazia sem usar o 'with'
    final_prs = Presentation()
    
    # Remove o slide inicial em branco que o `pptx` cria por padrão
    rId = final_prs.slides._sldIdLst[0].rId
    final_prs.part.drop_rel(rId)
    del final_prs.slides._sldIdLst[0]
    
    # Itera sobre cada equipe para criar e adicionar um slide modificado
    for team_data in all_teams_data:
        template_stream = BytesIO(template_bytes)
        
        # Abre a cópia do template em memória como um arquivo zip
        with zipfile.ZipFile(template_stream, 'a') as pptx_zip:
            slide_xml_path = 'ppt/slides/slide1.xml'
            xml_content = pptx_zip.read(slide_xml_path).decode('utf-8')
            
            # Substitui cada tag com os dados da equipe no XML
            for key, value in team_data.items():
                xml_value = value.replace('\n', '</a:t><a:br/><a:t>')
                xml_content = xml_content.replace(key, xml_value)
            
            # Escreve o XML modificado de volta no arquivo zip em memória
            pptx_zip.writestr(slide_xml_path, xml_content)

        # Abre a apresentação modificada (com um único slide)
        template_stream.seek(0)
        prs_with_one_slide = Presentation(template_stream)
        slide_to_add = prs_with_one_slide.slides[0]

        # Adiciona um slide em branco à apresentação final, usando um layout padrão
        slide_layout = final_prs.slide_layouts[0] 
        new_slide = final_prs.slides.add_slide(slide_layout)
        
        # Copia todos os elementos (formas, imagens, etc.) do slide modificado para o novo slide
        for shape in slide_to_add.shapes:
            new_el = deepcopy(shape.element)
            new_slide.shapes._spTree.add_element(new_el)
    
    # Salva a apresentação final completa no stream de memória
    final_prs.save(final_pres_stream)
    final_pres_stream.seek(0)
    return final_pres_stream

# === Interface Streamlit (sem alterações) ===

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
                    template_bytes = uploaded_template_file.getvalue()
                    final_presentation_stream = merge_presentations(template_bytes, teams_data)

                    st.session_state.pptx_buffer = final_presentation_stream
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
