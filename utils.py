import random
from datetime import datetime


def limpar_numero(num):
    """
    Limpa o número de telefone, removendo caracteres não numéricos.
    Garante que o número comece com o código DDI (ex: +55).
    """
    num = str(num).strip()
    for char in [" ", "-", "(", ")", ".", ".0"]:
        num = num.replace(char, "")
    if not num.startswith("+"):
        num = "+" + num
    return num


def obter_mensagem_por_horario():
    """
    Gera uma mensagem de saudação aleatória de acordo com a hora do dia.
    Visa contornar detecções de spam do WhatsApp diversificando as mensagens.
    """
    hora = datetime.now().hour
    if 5 <= hora < 12:
        return random.choice(
            [
                "Bom dia!",
                "Oi, bom dia.",
                "Bom dia, tudo bem?",
                "Olá!",
                "Oi, tudo certo?",
                "Bom dia, como vai?",
                "Bom dia.",
                "Opa! Bom dia.",
                "Olá, bom dia!",
                "Bom dia, tudo ótimo?",
            ]
        )
    elif 12 <= hora < 18:
        return random.choice(
            [
                "Boa tarde!",
                "Oi, boa tarde.",
                "Boa tarde, tudo bem?",
                "Olá!",
                "Oi, tudo certo?",
                "Boa tarde, como vai?",
                "Boa tarde.",
                "Oi, pode falar?",
                "Opa! Boa tarde.",
                "Olá, boa tarde!",
            ]
        )
    else:
        return random.choice(
            [
                "Boa noite!",
                "Oi, boa noite.",
                "Boa noite, tudo bem?",
                "Olá!",
                "Oi, tudo certo?",
                "Boa noite, como vai?",
                "Boa noite.",
                "Oi, consegue falar?",
                "Opa! Boa noite.",
                "Olá, boa noite!",
            ]
        )
