import time
import random
import pandas as pd
import os
import requests
import urllib.parse
import re
from datetime import datetime
from playwright.sync_api import sync_playwright


# ==============================================================================
# FUNÇÕES UTILITÁRIAS
# ==============================================================================

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
                "Oi, consegue falar?",
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
                "Boa tarde, preciso de você.",
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
                "Boa noite, preciso falar com você.",
                "Olá, boa noite!",
            ]
        )


# ==============================================================================
# 1. CONFIGURAÇÃO DA PLANILHA E DOWNLOAD DOS DADOS
# ==============================================================================
print("\n--- Configuração da Base de Leads ---")
entrada_usuario = input(
    "Digite o Link ou o ID da planilha do Google Sheets (precisa estar como 'Qualquer pessoa com o link'): "
).strip()

# Expressão regular para extrair apenas o ID único do link fornecido
match = re.search(r"/d/([a-zA-Z0-9-_]+)", entrada_usuario)
if match:
    sheet_id = match.group(1)
else:
    sheet_id = entrada_usuario

# Monta a URL de exportação direta do arquivo no formato XLSX
url = f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx"

diretorio = os.path.dirname(os.path.abspath(__file__))
arquivo_planilha_local = os.path.join(diretorio, "leads_local.xlsx")

try:
    print("\nBaixando a planilha do Google Sheets para uso local...")
    resposta = requests.get(url)
    if resposta.status_code == 200:
        with open(arquivo_planilha_local, "wb") as f:
            f.write(resposta.content)
        print("✅ Download concluído e salvo como 'leads_local.xlsx'!")
    else:
        print(f"❌ Falha ao baixar a planilha. HTTP Status: {resposta.status_code}")
        exit()

    print("\nLendo os dados transferidos para o arquivo local...")
    df = pd.read_excel(arquivo_planilha_local)

    df.columns = [str(col).strip().lower() for col in df.columns]

    print("✅ Conexão bem-sucedida!")
    print(f"📊 Colunas localizadas: {', '.join(df.columns)}")
    print(f"👥 Total de contatos identificados: {len(df)}")

    if "phone" not in df.columns:
        print(
            "\n❌ ERRO: Coluna 'phone' não encontrada. Verifique o cabeçalho na sua planilha."
        )
        exit()

    # Cria a coluna 'status' caso ela não exista na planilha
    if "status" not in df.columns:
        df["status"] = ""
        
    # Força a coluna a ser do tipo object para aceitar strings e evitar TypeError com floats
    df["status"] = df["status"].astype("object")

except Exception as e:
    print(f"\n❌ Erro ao carregar planilha do Google Sheets: {e}")
    exit()

print("\n" + "-" * 50)
confirmacao = (
    input("Tudo certo! Podemos iniciar os disparos de mensagens? (s/n): ")
    .strip()
    .lower()
)
if confirmacao != "s":
    print("Processamento cancelado.")
    exit()

# O progresso (status 'enviado' ou 'erro') será salvo no seu computador no arquivo 'leads_local.xlsx' recém-baixado.
print(
    f"\n⚠️ AVISO: O progresso será salvo localmente no arquivo 'leads_local.xlsx'. Assim, você poderá copiar os dados atualizados para a nuvem.\n"
)

# Configuração da sessão persistente do Playwright (não precisa scanear QR Code toda vez)
user_data_dir = os.path.join(diretorio, "user_data_playwright")

# ==============================================================================
# 2. LOOP DE PROCESSAMENTO E DISPARO COM PLAYWRIGHT
# ==============================================================================
with sync_playwright() as p:
    print(
        "Iniciando navegador Playwright... (Verifique o QR Code do WhatsApp se for a primeira vez)"
    )
    browser_context = p.chromium.launch_persistent_context(
        user_data_dir=user_data_dir,
        headless=False,
    )
    page = browser_context.new_page()

    for index, row in df.iterrows():
        # Regra de Segurança: Ignora contatos já abordados na planilha
        status_atual = str(row["status"]).strip().lower()
        if status_atual == "enviado":
            print(f"⏩ Pulando contato {index + 1}: Já enviado.")
            continue

        if pd.isna(row["phone"]):
            continue

        numero_final = limpar_numero(row["phone"])
        texto = obter_mensagem_por_horario()

        def tentar_enviar_mensagem(numero):
            texto_encoded = urllib.parse.quote(texto)
            link = f"https://web.whatsapp.com/send?phone={numero}&text={texto_encoded}"

            page.goto(link)

            # Esperar aparecer a caixa de texto (se o número for inválido, ela não aparecerá e dará timeout)
            # O timeout padrão é 30s. Como o WhatsApp Web pode demorar no 1º carregamento, deixamos 30s.
            page.wait_for_selector('div[contenteditable="true"]', timeout=30000)

            # Pausa extra para garantir que o chat abriu corretamente
            time.sleep(3)

            # Aperta Enter para enviar a mensagem
            page.keyboard.press("Enter")

            # Espera 4 segundos para garantir o envio antes de seguir
            time.sleep(4)

        try:
            print(f"🚀 [{index + 1}/{len(df)}] Enviando para {numero_final}...")
            tentar_enviar_mensagem(numero_final)
            
            df.at[index, "status"] = "enviado"
            df.to_excel(arquivo_planilha_local, index=False)

            atraso = random.randint(25, 54)  # Delay seguro contra banimentos
            print(f"✅ Sucesso! Próximo em {atraso} segundos...")
            time.sleep(atraso)

        except Exception as e:
            # Se falhou, verifica se o número tem padrão de AL (DDD 82) e possui apenas 8 dígitos (total 13 caracteres com +55)
            # Ex: +55 (3) + 82 (2) + XXXX-XXXX (8) = 13
            if numero_final.startswith("+5582") and len(numero_final) == 13:
                # Adiciona o dígito 9 logo após o DDD
                numero_com_9 = "+55829" + numero_final[5:]
                print(f"⚠️ Número não encontrado. Tentando de novo adicionando o '9': {numero_com_9}...")
                
                try:
                    tentar_enviar_mensagem(numero_com_9)
                    
                    df.at[index, "status"] = "enviado"
                    df.to_excel(arquivo_planilha_local, index=False)
                    
                    atraso = random.randint(25, 54)
                    print(f"✅ Sucesso com o dígito 9! Próximo em {atraso} segundos...")
                    time.sleep(atraso)
                    
                except Exception as e2:
                    print(f"❌ Erro definitivo. O número {numero_com_9} não está no WhatsApp.")
                    df.at[index, "status"] = "número não está no whatsapp"
                    df.to_excel(arquivo_planilha_local, index=False)
                    time.sleep(5)
            else:
                print(f"❌ Erro. O número {numero_final} não está no WhatsApp.")
                df.at[index, "status"] = "número não está no whatsapp"
                df.to_excel(arquivo_planilha_local, index=False)
                time.sleep(5)

    browser_context.close()

print("\n🎯 Processamento finalizado!")
