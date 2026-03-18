import os
import pandas as pd
from gsheets import baixar_e_ler_planilha
from whatsapp_bot import iniciar_disparos

def main():
    print("====================================")
    print("   🚀 WHATSAPP LEAD AUTOMATOR 🚀    ")
    print("====================================\n")

    diretorio = os.path.dirname(os.path.abspath(__file__))
    
    # 1. Obter e ler a planilha do Google Sheets
    df, arquivo_planilha_local = baixar_e_ler_planilha(diretorio)

    # 2. Configurar o limite e calcular estatísticas
    df["status"] = df.get("status", pd.Series(dtype="object")).fillna("").astype(str)
    
    total_leads = len(df)
    leads_contatados = len(df[df["status"].str.strip() != ""])
    leads_pendentes = total_leads - leads_contatados

    print("\n" + "-" * 50)
    print("--- Estatísticas da Planilha ---")
    print(f"👥 Total de Leads: {total_leads}")
    print(f"✅ Leads já contatados (Enviado/Erro): {leads_contatados}")
    print(f"⏳ Leads pendentes: {leads_pendentes}")
    print("-" * 50)

    print("\n--- Configuração de Limite ---")
    
    while True:
        limite_input = input("Quantos leads você quer contatar nesta sessão? (Aperte ENTER para enviar a todos os pendentes): ").strip()
        limite_envios = int(limite_input) if limite_input.isdigit() else None

        if limite_envios is not None and limite_envios > leads_pendentes:
            print(f"\n⚠️ ATENÇÃO: Você solicitou {limite_envios} envios, mas há apenas {leads_pendentes} leads pendentes.")
            print("Isso significa que o bot poderá repetir o contato com leads que já haviam sido abordados.")
            certeza = input("Tem certeza que deseja prosseguir com esse limite? (s/n): ").strip().lower()
            if certeza == 's':
                break
            else:
                print("\nVamos tentar de novo...")
        else:
            break

    if limite_envios:
        print(f"\n📌 O bot enviará mensagens para no máximo {limite_envios} lead(s).")
    else:
        print("\n📌 O bot enviará mensagens para todos os leads pendentes da planilha, sem limite numérico.")

    print("\n" + "-" * 50)
    confirmacao = input("Tudo certo! Podemos iniciar os disparos de mensagens? (s/n): ").strip().lower()
    
    if confirmacao != "s":
        print("Processamento cancelado.")
        exit(0)

    print("\n⚠️ AVISO: O progresso será salvo localmente no arquivo 'leads_local.xlsx'. Assim, você poderá copiar os dados atualizados para a nuvem.\n")

    # 3. Iniciar o disparo pelo WhatsApp Web
    iniciar_disparos(df, arquivo_planilha_local, diretorio, limite_envios)

if __name__ == "__main__":
    main()
