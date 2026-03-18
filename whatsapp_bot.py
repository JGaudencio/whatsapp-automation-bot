import time
import random
import urllib.parse
from playwright.sync_api import sync_playwright
import os
import pandas as pd
from datetime import datetime
from utils import limpar_numero, obter_mensagem_por_horario

def iniciar_disparos(df, arquivo_planilha_local, diretorio, limite_envios):
    """
    Inicia o navegador Playwright, faz a varredura na planilha e gerencia
    os envios via WhatsApp Web com o limite especificado de contatos.
    """
    user_data_dir = os.path.join(diretorio, "user_data_playwright")
    envios_feitos = 0

    with sync_playwright() as p:
        print("\nIniciando navegador Playwright... (Verifique o QR Code do WhatsApp se for a primeira vez)")
        browser_context = p.chromium.launch_persistent_context(
            user_data_dir=user_data_dir,
            headless=False,
        )
        page = browser_context.new_page()

        for index, row in df.iterrows():
            if limite_envios is not None and envios_feitos >= limite_envios:
                print(f"\n✋ Limite de {limite_envios} envio(s) atingido. Encerrando a sessão.")
                break

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

                # O timeout padrão é 30s. Como o WhatsApp Web pode demorar no 1º carregamento, deixamos 30s.
                page.wait_for_selector('div[contenteditable="true"]', timeout=30000)
                time.sleep(3)
                page.keyboard.press("Enter")
                time.sleep(4)

            try:
                print(f"🚀 [{index + 1}/{len(df)}] Enviando para {numero_final}...")
                tentar_enviar_mensagem(numero_final)
                
                df.at[index, "status"] = "enviado"
                df.at[index, "data"] = datetime.now().strftime('%d/%m/%Y')
                df.to_excel(arquivo_planilha_local, index=False)

                envios_feitos += 1
                atraso = random.randint(25, 54)  # Delay seguro contra banimentos
                print(f"✅ Sucesso! Próximo em {atraso} segundos...")
                time.sleep(atraso)

            except Exception as _e:
                # Se falhou, verifica se o número tem padrão de AL (DDD 82) e possui apenas 8 dígitos
                if numero_final.startswith("+5582") and len(numero_final) == 13:
                    numero_com_9 = "+55829" + numero_final[5:]
                    print(f"⚠️ Número não encontrado. Tentando de novo adicionando o '9': {numero_com_9}...")
                    
                    try:
                        tentar_enviar_mensagem(numero_com_9)
                        
                        df.at[index, "status"] = "enviado"
                        df.at[index, "data"] = datetime.now().strftime('%d/%m/%Y')
                        df.to_excel(arquivo_planilha_local, index=False)
                        
                        envios_feitos += 1
                        atraso = random.randint(25, 54)
                        print(f"✅ Sucesso com o dígito 9! Próximo em {atraso} segundos...")
                        time.sleep(atraso)
                        
                    except Exception as _e2:
                        print(f"❌ Erro definitivo. O número {numero_com_9} não está no WhatsApp.")
                        df.at[index, "status"] = "número não está no whatsapp"
                        df.at[index, "data"] = datetime.now().strftime('%d/%m/%Y')
                        df.to_excel(arquivo_planilha_local, index=False)
                        time.sleep(5)
                else:
                    print(f"❌ Erro. O número {numero_final} não está no WhatsApp.")
                    df.at[index, "status"] = "número não está no whatsapp"
                    df.at[index, "data"] = datetime.now().strftime('%d/%m/%Y')
                    df.to_excel(arquivo_planilha_local, index=False)
                    time.sleep(5)

        browser_context.close()

    print("\n🎯 Processamento finalizado!")
