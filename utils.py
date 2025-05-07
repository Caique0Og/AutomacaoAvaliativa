from pptx.dml.color import RGBColor

def validate_ternary_input(ternary_code, length=32):
    """Valida input com 3 possibilidades (0, 1, 2)"""
    if len(ternary_code) != length or not all(c in ('0', '1', '2') for c in ternary_code):
        raise ValueError(f"O código deve ter {length} caracteres (0, 1 ou 2)")

def check_respostas_absolutas(ternary_code):
    """Verifica os casos de respostas absolutas (considera APENAS '1's)"""
    grupos = {
        'tracionar': {'indices': [12,13,14,15,16,17], 'min': 2},  # Índices 0-based
        'escalar': {'indices': [5,6,7,9,10,11], 'min': 2},
        'lançar': {'indices': [20,21,22,23,24,25], 'min': 2},
        'finalizar': {'indices': [27,28,29], 'min': 2},
        'prototipar': {'indices': [30,31], 'min': 2}
    }
    
    mensagens = {
        'tracionar': "Controle para tracionar Precisa de investimento",
        'escalar': "Controle para escalar Precisa de investimento",
        'lançar': "Controle para lançar",
        'finalizar': "Controle para finalizar o produto e levar ao mercado",
        'prototipar': "Controle para prototipar"
    }
    
    for grupo, dados in grupos.items():
        # IMPORTANTE: Continua só contando '1's para respostas absolutas
        count = sum(1 for i in dados['indices'] if ternary_code[i] == '1')
        if count >= (len(dados['indices']) - dados['min'] + 1):
            return {
                'show_special': True,
                'message': mensagens[grupo],
                'position': (1.3, 1.8),
                'color': RGBColor(200, 200, 200)  # Cinza
            }
    
    return {'show_special': False}

def analyze_tech_control(ternary_code):
    """Analisa o controle de tecnologia (agora considera 2=1)"""
    tech_indices = [8, 9, 12, 13, 19, 20, 25, 26, 27, 29, 30]
    # CORREÇÃO: Agora conta tanto '1' quanto '2' como positivos
    count = sum(1 for i in tech_indices if ternary_code[i] in ('1'))
    
    if count >= 6:
        return "Controle confirmado da tecnologia (análise preliminar)|Detalhes requerem entrevista"
    elif count >= 3:
        return "Déficit parcial no controle tecnológico (produto digital)|Verificação necessária"
    else:
        return "Déficit tecnológico identificado (análise inicial)|Avaliação detalhada requerida"

def analyze_second_box(ternary_code):
    """Lógica para a segunda caixa (agora considera 2=1)"""
    relevant_indices = [5, 6, 7, 8, 11, 15, 19, 22]
    # CORREÇÃO: Agora conta tanto '1' quanto '2' como positivos
    count = sum(1 for i in relevant_indices if ternary_code[i] in ('1'))
    
    if count >= 6:
        return "Controle dos processos comerciais (análise preliminar)|Detalhes requerem entrevista"
    elif count >= 3:
        return "Possíveis lacunas identificadas (verificação necessária)"
    else:
        return "Déficit de mercado identificado (análise inicial)|Avaliação detalhada requerida"

def analyze_third_box(ternary_code):
    """Lógica para a terceira caixa (agora considera 2=1)"""
    relevant_indices = [12, 16, 17, 18, 23, 24, 29, 30, 31]
    # CORREÇÃO: Agora conta tanto '1' quanto '2' como positivos
    count = sum(1 for i in relevant_indices if ternary_code[i] in ('1'))
    
    if count >= 6:
        return "Controle nos processos operacionais (análise preliminar)"
    elif count >= 3:
        return "Possíveis lacunas operacionais (verificação necessária)"
    else:
        return "Déficit operacional identificado (avaliação detalhada requerida)"
    

#Terceiro SLide:

def check_prototipo_fase(ternary_code):
    """Verifica se alguma das bolinhas 30-32 está marcada como 1"""
    prototipo_indices = [29, 30, 31]  # Índices 30-32 no questionário (0-based)
    
    if any(ternary_code[i] == '1' for i in prototipo_indices):
        return True
    else:
        return False