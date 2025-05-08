from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def add_textbox(slide, text, abs_x, abs_y, width=2, height=0.8, font_size=12, is_special=False):
    """Adiciona caixa de texto com formatação condicional"""
    left = Inches(abs_x)
    top = Inches(abs_y)
    text_box = slide.shapes.add_textbox(left, top, Inches(width), Inches(height))
    text_frame = text_box.text_frame
    text_frame.word_wrap = True
    paragraph = text_frame.paragraphs[0]

    # Formatação especial
    if is_special:
        fill = text_box.fill
        fill.solid()
        fill.fore_color.rgb = RGBColor(200, 200, 200)
        paragraph.alignment = PP_ALIGN.CENTER
    else:
        paragraph.alignment = PP_ALIGN.LEFT

    # Adiciona texto com quebras
    for line in text.split('|'):
        run = paragraph.add_run()
        run.text = line.strip()
        run.font.size = Pt(font_size)
        run.font.bold = True
        paragraph.add_run().text = ''
    # Remove última quebra extra
    if text_frame.paragraphs[0].runs:
        text_frame.paragraphs[0].runs[-1].text = ''

def create_slide(prs, img_path, text_input, time_input, ternary_code, 
                show_chart=False, special_message=None, 
                analysis_texts=None, show_analysis=False):
    """Cria slide com lógica condicional completa"""
    slide = prs.slides.add_slide(prs.slide_layouts[6])

    # Imagem de fundo
    slide.shapes.add_picture(
        img_path, 
        Inches(0), Inches(0),
        width=Inches(10), height=Inches(5.625)
    )

    # Gráfico (apenas no primeiro slide)
    if show_chart:
        from chart import create_questionnaire_chart
        create_questionnaire_chart(slide, ternary_code)

    # Textos fixos
    add_textbox(slide, text_input, 8.5, 0.07, font_size=14)  # nomeStartup
    add_textbox(slide, time_input, 8.5, 0.50, font_size=12)  # carimbo 

    # Lógica de exibição condicional
    if special_message and special_message['show_special']:
        # Modo resposta absoluta (caixa cinza)
        add_textbox(
            slide,
            special_message['message'],
            special_message['position'][0],
            special_message['position'][1],
            width=3,
            height=0.8,
            font_size=12,
            is_special=True
        )
    elif show_analysis and analysis_texts:
        # Modo análises normais
        positions = [(1.22, 1.535), (1.22, 1.933), (1.22, 2.4)]
        for text, pos in zip(analysis_texts, positions):
            add_textbox(
                slide,
                text,
                pos[0],
                pos[1],
                width=3,
                height=0.8,
                font_size=10,
                is_special=False
            )

    # Nova lógica para o terceiro slide (quando não é nenhum dos casos anteriores)
    elif img_path == "templates/img3.png":
        from utils import check_prototipo_fase
        if check_prototipo_fase(ternary_code):
            # Primeira caixa de texto (sem fundo cinza)
            add_textbox(
                slide,
                "Finalizar a ideação do produto",
                1.3, 1.8,
                width=4,
                height=0.5,
                font_size=14,
                is_special=False
            )
            # Segunda caixa de texto (sem fundo cinza)
            add_textbox(
                slide,
                "Nesse estágio aconselhamos formar um time com três perfis: tecnologia exigida pelo produto (\"perfil hacker\"), design de produto (\"perfil hipster\") e vendas / marketing (\"perfil hustler\")",
                1.3, 2.2,
                width=4,
                height=1.2,
                font_size=12,
                is_special=False
            )

def add_reference_grid_slide(prs):
    """
    Adiciona um slide em branco com marcações de coordenadas em polegadas.
    Útil para posicionamento manual dos elementos.
    """
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    width = 15    # polegadas
    height = 5.625

    # Desenha linhas verticais (cada 1 pol)
    for x in range(0, 11):
        left = Inches(x)
        top = Inches(0)
        shape = slide.shapes.add_shape(
            1, left, top, Inches(0.01), Inches(height)
        )
        shape.line.color.rgb = RGBColor(200, 200, 200)
        # Marca o valor X
        textbox = slide.shapes.add_textbox(left, Inches(0.1), Inches(0.5), Inches(0.3))
        tf = textbox.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f"{x}\""
        run.font.size = Pt(10)
        run.font.bold = True

    # Desenha linhas horizontais (cada 0.5 pol)
    y = 0.0
    while y <= height + 0.01:
        left = Inches(0)
        top = Inches(y)
        shape = slide.shapes.add_shape(
            1, left, top, Inches(width), Inches(0.01)
        )
        shape.line.color.rgb = RGBColor(200, 200, 200)
        # Marca o valor Y
        textbox = slide.shapes.add_textbox(Inches(0.1), top, Inches(0.5), Inches(0.3))
        tf = textbox.text_frame
        p = tf.paragraphs[0]
        run = p.add_run()
        run.text = f"{y:.2f}\""
        run.font.size = Pt(10)
        run.font.bold = True
        y += 0.5

    # Título do slide
    title_box = slide.shapes.add_textbox(Inches(2), Inches(0.2), Inches(6), Inches(0.5))
    title_frame = title_box.text_frame
    title_frame.text = "Slide de Referência - Coordenadas em Polegadas (960x540px = 10x5.625\")"
    title_frame.paragraphs[0].font.size = Pt(18)
    title_frame.paragraphs[0].font.bold = True
    title_frame.paragraphs[0].alignment = PP_ALIGN.CENTER
