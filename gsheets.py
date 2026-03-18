import pandas as pd
import requests
import os
import re

def baixar_e_ler_planilha(diretorio):
    """
    Verifica se já existe uma planilha local e permite utilizá-la. Caso contrário, ou
    se o usuário preferir baixar os dados frescos, pede o Link da planilha, faz o 
    download via Google Sheets API e retorna o DataFrame pandas.
    """
    arquivo_planilha_local = os.path.join(diretorio, "leads_local.xlsx")

    # Verifica se já existe um arquivo local com envios parciais
    if os.path.exists(arquivo_planilha_local):
        print("\n--- Planilha Local Encontrada ---")
        usar_local = input(
            "Foi detectado um arquivo 'leads_local.xlsx' de sessões anteriores.\n"
            "Deseja usar essa planilha para continuar de onde parou? (s/n): "
        ).strip().lower()

        if usar_local == 's':
            print("\nLendo os dados transferidos para o arquivo local...")
            try:
                df = pd.read_excel(arquivo_planilha_local)
                df.columns = [str(col).strip().lower() for col in df.columns]
                
                print("✅ Leitura local bem-sucedida!")
                print(f"📊 Colunas localizadas: {', '.join(df.columns)}")
                print(f"👥 Total de contatos identificados: {len(df)}")
                
                if "phone" not in df.columns:
                    print("\n❌ ERRO: Coluna 'phone' não encontrada na planilha local.")
                    exit(1)
                
                if "status" not in df.columns:
                    df["status"] = ""
                df["status"] = df["status"].astype("object")
                
                if "data" not in df.columns:
                    df["data"] = ""
                df["data"] = df["data"].astype("object")
                
                return df, arquivo_planilha_local
            except Exception as e:
                print(f"\n❌ Erro ao ler planilha local: {e}")
                exit(1)

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

    try:
        print("\nBaixando a planilha do Google Sheets para uso local...")
        resposta = requests.get(url)
        if resposta.status_code == 200:
            with open(arquivo_planilha_local, "wb") as f:
                f.write(resposta.content)
            print("✅ Download concluído e salvo como 'leads_local.xlsx'!")
        else:
            print(f"❌ Falha ao baixar a planilha. HTTP Status: {resposta.status_code}")
            exit(1)

        print("\nLendo os dados transferidos para o arquivo local...")
        df = pd.read_excel(arquivo_planilha_local)

        df.columns = [str(col).strip().lower() for col in df.columns]

        print("✅ Conexão bem-sucedida!")
        print(f"📊 Colunas localizadas: {', '.join(df.columns)}")
        print(f"👥 Total de contatos identificados: {len(df)}")

        if "phone" not in df.columns:
            print("\n❌ ERRO: Coluna 'phone' não encontrada. Verifique o cabeçalho na sua planilha.")
            exit(1)

        # Cria a coluna 'status' e 'data' caso elas não existam na planilha
        if "status" not in df.columns:
            df["status"] = ""
        df["status"] = df["status"].astype("object")

        if "data" not in df.columns:
            df["data"] = ""
        df["data"] = df["data"].astype("object")

        return df, arquivo_planilha_local

    except Exception as e:
        print(f"\n❌ Erro ao carregar planilha do Google Sheets: {e}")
        exit(1)
