
    # Texto completo da shape (mesmo se o placeholder estiver quebrado)
    full_text_shape = "".join(run.text for p in shape.text_frame.paragraphs for run in p.runs)

    for key, value in team_data.items():
        if key not in full_text_shape:
            continue

        tf = shape.text_frame
        tf.clear()

        # NOMES (líder/acompanhante/alunos)
        if key == "{{NOMES_ALUNOS}}":
            linhas = value.split("\n")
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

        # ALCANCE (com parte negrito + sublinhado)
        elif key == "{{LANCAMENTOS_VALIDOS}}":
            p = tf.paragraphs[0]
            match = re.match(r"(ALCANCE:\s*)([\d,.]+ m)", value, re.IGNORECASE)
            if match:
                prefix, numero = match.groups()
                run1 = p.add_run()
                run1.text = prefix
                run1.font.name = "Lexend"
                run1.font.bold = False
                run1.font.size = Pt(28)
                run1.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)

                run2 = p.add_run()
                run2.text = numero
                run2.font.name = "Lexend"
                run2.font.bold = True
                run2.font.underline = True
                run2.font.size = Pt(35)
                run2.font.color.rgb = RGBColor(0x00, 0x6F, 0xC0)

        # CAMPOS NORMAIS
        else:
            p = tf.paragraphs[0]
            run = p.add_run()
            run.text = value
            run.font.name = "Lexend"
            run.font.bold = True

            if key == "{{NOME_EQUIPE}}":
                run.font.size = Pt(20)
            elif key in ("{{NOME_ESCOLA}}", "{{CIDADE_UF}}"):
                run.font.size = Pt(22)
            else:
                run.font.size = Pt(18)

            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            p.alignment = PP_ALIGN.CENTER
            p.line_spacing = None

def gerar_apresentacao(dados, template_stream):
    prs = Presentation(template_stream)
    if not dados or not prs.slides:
        return prs

    modelo = prs.slides[0]
    for _ in range(len(dados) - 1):
        duplicate_slide_with_media(prs, modelo)

    for slide, team in zip(prs.slides, dados):
        for shape in slide.shapes:
            replace_placeholders_in_shape(shape, team)

    return prs

# -------------------- INTERFACE STREAMLIT --------------------
docx_file = st.file_uploader("📄 Arquivo DOCX", type=["docx", "DOCX"])
pptx_file = st.file_uploader("📊 Arquivo PPTX modelo", type=["pptx", "PPTX"])

if st.button("✨ Gerar Apresentação"):
    if not docx_file or not pptx_file:
        st.warning("Envie ambos os arquivos.")
    else:
        try:
            dados = extrair_dados(docx_file)
            if not dados:
                st.warning("Nenhum dado encontrado.")
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

