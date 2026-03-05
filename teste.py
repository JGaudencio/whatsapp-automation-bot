import pywhatkit as kit
import pyautogui
import time
import random
import pandas as pd
import os
from datetime import datetime

# 1. Configuração do Caminho
diretorio = os.path.dirname(os.path.abspath(__file__))
caminho_planilha = os.path.join(diretorio, 'leads.xlsx')

def limpar_numero(num):
    num = str(num).strip()
    for char in [" ", "-", "(", ")", ".", ".0"]:
        num = num.replace(char, "")
    if not num.startswith('+'):
        num = '+' + num
    return num

def obter_mensagem_por_horario():
    hora = datetime.now().hour
    # Variáveis curtas e diretas
    if 5 <= hora < 12:
        return random.choice(["Bom dia!", "Oi, bom dia.", "Bom dia, tudo bem?", "Olá!", "Oi, tudo certo?", "Bom dia, como vai?", "Bom dia.", "Oi, consegue falar?", "Olá, bom dia!", "Bom dia, tudo ótimo?"])
    elif 12 <= hora < 18:
        return random.choice(["Boa tarde!", "Oi, boa tarde.", "Boa tarde, tudo bem?", "Olá!", "Oi, tudo certo?", "Boa tarde, como vai?", "Boa tarde.", "Oi, pode falar?", "Boa tarde, preciso de você.", "Olá, boa tarde!"])
    else:
        return random.choice(["Boa noite!", "Oi, boa noite.", "Boa noite, tudo bem?", "Olá!", "Oi, tudo certo?", "Boa noite, como vai?", "Boa noite.", "Oi, consegue falar?", "Boa noite, preciso falar com você.", "Olá, boa noite!"])

try:
    df = pd.read_excel(caminho_planilha)
    
    df.columns = [str(col).strip().lower() for col in df.columns]
    
    if 'telefone' not in df.columns:
        print("❌ ERRO: Coluna 'telefone' não encontrada.")
        exit()

    # Cria a coluna 'status' caso ela não exista na planilha
    if 'status' not in df.columns:
        df['status'] = ""

except Exception as e:
    print(f"❌ Erro ao carregar planilha: {e}")
    exit()

# 2. Loop de Processamento
for index, row in df.iterrows():
    # --- A TRAVA: Pula se o status já for 'enviado' ---
    status_atual = str(row['status']).strip().lower()
    if status_atual == "enviado":
        print(f"⏩ Pulando contato {index + 1}: Já enviado.")
        continue
    
    if pd.isna(row['telefone']):
        continue

    numero_final = limpar_numero(row['telefone'])
    texto = obter_mensagem_por_horario()
    
    try:
        print(f"🚀 [{index + 1}/{len(df)}] Enviando para {numero_final}...")
        
        # Envia a mensagem
        kit.sendwhatmsg_instantly(numero_final, texto, wait_time=20, tab_close=True)
        
        # Pequena pausa para garantir que o navegador abriu e enviou
        time.sleep(7) 
        pyautogui.hotkey('ctrl', 'w')
        
        df.at[index, 'status'] = "enviado"
        df.to_excel(caminho_planilha, index=False)
        
        atraso = random.randint(25, 54) # Delay seguro contra banimentos
        print(f"✅ Sucesso! Próximo em {atraso} segundos...")
        time.sleep(atraso)

    except Exception as e:
        print(f"❌ Erro no envio para {numero_final}: {e}")
        df.at[index, 'status'] = "erro"
        df.to_excel(caminho_planilha, index=False)
        time.sleep(5)

print("\n🎯 Processamento finalizado!")