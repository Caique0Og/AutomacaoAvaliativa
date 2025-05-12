from pptx.dml.color import RGBColor

def validate_ternary_input(ternary_code, length=32):
    if len(ternary_code) != length or not all(c in ('0', '1', '2') for c in ternary_code):
        raise ValueError(f"O código deve ter {length} caracteres (0, 1 ou 2)")

def check_respostas_absolutas(ternary_code):
    grupos = {
        'tracionar': {'indices': [12,13,14,15,16,17], 'min': 2},
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
                'messagePadrao': mensagemPadrao,
                'position': (4.2, 3),
                'positionPadrao': (7.5, 3),
                'color': RGBColor(200, 200, 200),
            }
    
    return {
        'show_special': False,
        'messagePadrao': mensagemPadrao
    }

# ============= FUNÇÕES DE ANÁLISE ATUALIZADAS =============
def analyze_tech_control(ternary_code):
    """Analisa o controle tecnológico da startup"""
    tech_indices = [0,1,2,3]  # Índices para controle tecnológico
    score = sum(int(ternary_code[i]) for i in tech_indices)
    
    if score >= 3:
        return "Domínio tecnológico consolidado"
    elif score >= 2:
        return "Controle tecnológico adequado"
    else:
        return "Recomenda-se fortalecer a base tecnológica"

def analyze_second_box(ternary_code):
    """Analisa a segunda caixa de questões"""
    second_indices = [4,5,6,7,8,9,10,11]  # Índices para análise operacional
    score = sum(int(ternary_code[i]) for i in second_indices)
    
    if score >= 6:
        return "Processos operacionais otimizados"
    elif score >= 4:
        return "Processos operacionais funcionais"
    else:
        return "Necessidade de revisão dos processos operacionais"

def analyze_third_box(ternary_code):
    """Analisa a terceira caixa de questões"""
    third_indices = [12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29,30,31]  # Índices completos
    score = sum(int(ternary_code[i]) for i in third_indices)
    
    if score >= 15:
        return "Prontidão para mercado consolidada"
    elif score >= 10:
        return "Preparação em estágio intermediário"
    else:
        return "Recomenda-se desenvolver capacidades antes de avançar"


def check_prototipo_fase(ternary_code):
    """Verifica marcações para o terceiro slide (único com \n)"""
    grupos = {
        'grupo1': {'indices': [4,5,6,7,8,9,10], 'min': 6, 'msg': {
            'titulo': "Preparar para escalar",
            'conteudo': "Necessidade distribuída com crescimento da operação"}},
        'grupo2': {'indices': [11,12,13,14,15,16], 'min': 4, 'msg': {
            'titulo': "Tracionar receita",
            'conteudo': "Equipe comercial e marketing necessária"}},
        'grupo3': {'indices': [19,20,21,22,23,24], 'min': 5, 'msg': {
            'titulo': "Go to market",
            'conteudo': "Necessidade distribuída com crescimento"}},
        'grupo4': {'indices': [26,27,28], 'min': 3, 'msg': {
            'titulo': "Finalizar produto",
            'conteudo': "Preparação para lançamento no mercado"}},
        'grupo5': {'indices': [30,31], 'min': 2, 'msg': {
            'titulo': "Finalizar ideação",
            'conteudo': "Time com perfis hacker, hipster e hustler"}}
    }

    mensagens = []
    
    for grupo, dados in grupos.items():
        count = sum(ternary_code[i] == '1' for i in dados['indices'])
        if count >= dados['min']:
            mensagens.append(dados['msg'])
    
    return mensagens if mensagens else None