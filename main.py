from pptx import Presentation
from pptx.util import Inches
from presentation import create_slide, add_reference_grid_slide
from utils import (check_respostas_absolutas, validate_ternary_input,
                  analyze_tech_control, analyze_second_box, analyze_third_box,
                  check_prototipo_fase)
import os
from datetime import datetime

def main():
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    print("\n" + "="*40)
    print("QUESTIONÁRIO DE AVALIAÇÃO DE STARTUP")
    print("="*40)
    
    # Modo desenvolvimento - Adiciona grade de referência no início
    dev_mode = input("\nAtivar modo desenvolvimento com grade de referência? (s/n): ").lower().strip() == 's'
    
    if dev_mode:
        add_reference_grid_slide(prs)
        print("\n[Modo Desenvolvimento] Slide de grade adicionado como primeiro slide")
    
    ternary_code = input("\nDigite as respostas (32 caracteres 0/1/2): ").strip()
    
    try:
        validate_ternary_input(ternary_code)
    except ValueError as e:
        print(f"\nERRO: {e}")
        return

    text_input = input("Nome da Startup: ")
    time_input = input("Carimbo data/hora: ")
    
    abs_check = check_respostas_absolutas(ternary_code)
    prototipo_msg = check_prototipo_fase(ternary_code)
    
    # Mensagem padrão caso não haja mensagens específicas
    if not prototipo_msg:
        prototipo_msg = [{
            'titulo': "Análise Complementar",
            'conteudo': "Avaliação completa das capacidades da startup"
        }]
    
    print("\n=== RESULTADOS ===")
    if abs_check['show_special']:
        print(f"RESPOSTA ABSOLUTA: {abs_check['message']}")
    
    print("\nMENSAGENS PARA SLIDE 3:")
    for msg in prototipo_msg:
        print(f"{msg['titulo']}: {msg['conteudo']}")

    # Cria slides de conteúdo (agora sempre 3 slides principais)
    
    # Slide 1 - Gráfico (obrigatório)
    create_slide(prs, "templates/img1.png", text_input, time_input, ternary_code, show_chart=True)
    
    # Slide 2 - Análise ou Resposta Absoluta (obrigatório)
    if abs_check['show_special']:
        create_slide(
            prs, "templates/img2.png", text_input, time_input, ternary_code,
            special_message=abs_check
        )
    else:
        analysis_texts = [
            analyze_tech_control(ternary_code),
            analyze_second_box(ternary_code),
            analyze_third_box(ternary_code)
        ]
        create_slide(
            prs, "templates/img2.png", text_input, time_input, ternary_code,
            analysis_texts=analysis_texts, show_analysis=True
        )
    
    # Slide 3 - Protótipo (agora obrigatório com conteúdo padrão ou específico)
    create_slide(
        prs, "templates/img3.png", text_input, time_input, ternary_code,
        prototipo_msg=prototipo_msg
    )
    
    # Adiciona grade de referência no final se em modo desenvolvimento
    if dev_mode:
        add_reference_grid_slide(prs)
        print("[Modo Desenvolvimento] Slide de grade adicionado como último slide")
    
    # Salvar apresentação
    filename = f"Apresentacao_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pptx"
    prs.save(filename)
    print(f"\n✅ Apresentação salva como: {os.path.abspath(filename)}")
    
    # Abre automaticamente no PowerPoint (Windows)
    try:
        os.startfile(filename)
    except:
        print("Não foi possível abrir o arquivo automaticamente")

if __name__ == "__main__":
    main()