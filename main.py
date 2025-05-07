from pptx import Presentation
from presentation import create_slide, add_reference_grid_slide
from utils import (check_respostas_absolutas, validate_ternary_input,
                  analyze_tech_control, analyze_second_box, analyze_third_box)
import os
from datetime import datetime
from pptx.util import Inches

def main():
    prs = Presentation()
    # Ajusta o tamanho do slide para 960x540 px (10 x 5.625 pol)
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(5.625)
    
    print("\n" + "="*40)
    print("QUESTIONÁRIO DE AVALIAÇÃO DE STARTUP")
    print("="*40)
    print("\nDigite as respostas (32 caracteres no total):")
    print("0 = Não | 1 = Sim | 2 = Sem resposta")
    ternary_code = input("\nRespostas: ").strip()
    
    try:
        validate_ternary_input(ternary_code)
    except ValueError as e:
        print(f"\nERRO: {e}")
        return

    text_input = input("\nNome da Startup: ")
    time_input = input("Carimbo data/hora: ")
    
    # Verifica respostas absolutas
    abs_check = check_respostas_absolutas(ternary_code)
    
    # Cria os slides
    create_slide(prs, "templates/img1.png", text_input, time_input, ternary_code, show_chart=True)
    
    # Segundo slide - lógica condicional
    if abs_check['show_special']:
        print(f"\nRESPOSTA ABSOLUTA: {abs_check['message'].replace('|', ' ')}")
        create_slide(
            prs, 
            "templates/img2.png", 
            text_input, 
            time_input, 
            ternary_code,
            special_message=abs_check
        )
    else:
        analysis_texts = [
            analyze_tech_control(ternary_code),
            analyze_second_box(ternary_code),
            analyze_third_box(ternary_code)
        ]
        print("\nANÁLISES NORMAIS:")
        for text in analysis_texts:
            print(f"- {text.replace('|', ' ')}")
            
        create_slide(
            prs,
            "templates/img2.png",
            text_input,
            time_input,
            ternary_code,
            analysis_texts=analysis_texts,
            show_analysis=True
        )
    
    # Terceiro slide
    create_slide(prs, "templates/img3.png", text_input, time_input, ternary_code)
    
    # ===== ATIVE OU DESATIVE O SLIDE DE REFERÊNCIA AQUI =====
    # Para exibir a grade de coordenadas, descomente a linha abaixo:
    add_reference_grid_slide(prs)
    # Para remover a grade, basta comentar a linha acima.
    # =======================================================

    # Salvar arquivo
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    file_path = f"Apresentacao_{text_input[:20]}_{timestamp}.pptx"
    
    try:
        prs.save(file_path)
        print(f"\n✅ Apresentação salva como: {os.path.abspath(file_path)}")
        os.startfile(file_path)
    except Exception as e:
        print(f"\n❌ Erro ao salvar: {str(e)}")

if __name__ == "__main__":
    main()
