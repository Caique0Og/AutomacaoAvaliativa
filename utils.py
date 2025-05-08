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
        'prototipar': "empresa controla as competências necessárias prototipar seu produto, precisando de investimento."
    }
    
    mensagemPadrao = "Avaliados na entrevista de diagnóstico"
    
    for grupo, dados in grupos.items():
        count = sum(1 for i in dados['indices'] if ternary_code[i] == '1')
        if count >= (len(dados['indices']) - dados['min'] + 1):
            return {
                'show_special': True,
                'message': mensagens[grupo],
                'messagePadrao': mensagemPadrao,  # Nova mensagem padrão
                'position': (4, 2.87),          # Posição da primeira caixa
                'positionPadrao': (6, 2.87),     # Posição da nova caixa (ajuste conforme necessário)
                'color': RGBColor(200, 200, 200), # Cor das caixas
            }
    
    return {
        'show_special': False,
        'messagePadrao': mensagemPadrao  # Retorna a mensagem padrão mesmo quando não há respostas absolutas
    }

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
    """Verifica as marcações nos grupos 20-25, 27-29 ou 30-32 (1-based) e retorna mensagens específicas."""
    # Definindo os grupos (índices 0-based)
    grupo1 = [26, 27, 28]  # Posições 27-29 no questionário
    grupo2 = [29, 30, 31]  # Posições 30-32 no questionário
    grupo3 = [19, 20, 21, 22, 23, 24]  # Posições 20-25 no questionário (23 = índice 22)
    
    # Verifica marcações em cada grupo
    marcou_grupo1 = any(ternary_code[i] == '1' for i in grupo1)
    marcou_grupo2 = any(ternary_code[i] == '1' for i in grupo2)
    marcou_grupo3 = any(ternary_code[i] == '1' for i in grupo3)
    
    # Verifica conflitos entre grupos
    grupos_marcados = sum([marcou_grupo1, marcou_grupo2, marcou_grupo3])
    if grupos_marcados > 1:
        raise ValueError(
            "Erro: Você não pode marcar opções de múltiplos grupos simultaneamente. "
            "Escolha apenas um grupo entre 20-25, 27-29 ou 30-32."
        )
    
    # Lógica para Grupo 3 (20-25)
    if marcou_grupo3:
        # Verifica se marcou todas ou se apenas o 23 (índice 22) está como 0
        marcou_todas = all(ternary_code[i] == '1' for i in grupo3)
        apenas_23_nao = (ternary_code[22] == '0') and all(ternary_code[i] == '1' for i in grupo3 if i != 22)
        
        if marcou_todas or apenas_23_nao:
            return (
                "Go to market\n"
                "Necessidade distribuída a ser identificada com o crescimento da operação, "
                "por vezes, primeiro vem vendas, por outras, produção."
            )
        else:
            # Adicione outras condições para o grupo3 aqui se necessário
            return "Condição não especificada para o grupo 20-25"
    
    # Lógica para Grupo 1 (27-29)
    if marcou_grupo1:
        marca_27 = ternary_code[26] == '1'
        marca_28 = ternary_code[27] == '1'
        marca_29 = ternary_code[28] == '1'
        
        if marca_27 and marca_28 and marca_29:
            return (
                "Finalizar o desenvolvimento do produto\n"
                "Necessidade distribuída a ser identificada com o lançamento da operação, "
                "por vezes, primeiro vem vendas, por outras, produção."
            )
        elif marca_27 and marca_28:
            return (
                "Go to market (finalizar o desenvolvimento do produto)\n"
                "Profissionais de marketing (promoção) e vendas (presumindo o controle da tecnologia)."
            )
        elif marca_28 and marca_29 or marca_27 or marca_28 or marca_29:
            return (
                "Go to market (validar o produto)\n"
                "Nesse estágio aconselhamos formar um time com três perfis: tecnologia exigida pelo produto "
                "('perfil hacker'), design de produto ('perfil hipster') e vendas/marketing ('perfil hustler')."
            )
    
    # Lógica para Grupo 2 (30-32)
    elif marcou_grupo2:
        return "Lógica para prototipagem (Grupo 30-32) - Mensagem a ser definida."
    
    return None