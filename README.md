# WhatsApp Lead Automator 🚀

Este projeto é um bot de automação em Python desenvolvido para facilitar o contato com leads. O script conecta-se a uma planilha pública do Google Sheets, baixa os contatos, realiza o tratamento dos números de telefone e automatiza o envio de mensagens personalizadas via WhatsApp Web através de um navegador controlado pelo **Playwright**.

## 🛠️ Novas Funcionalidades (v2.0)

- **Integração com Google Sheets na Nuvem:** O bot lê os dados diretamente de um link de compartilhamento público, sem necessidade de baixar a planilha manualmente.
- **Automação Segura em Segundo Plano:** Graças à migração para o `Playwright` com sessão persistente (Persistent Context), o bot abre um navegador próprio. Isso libera totalmente o seu mouse e teclado enquanto as mensagens são enviadas.
- **Controle Local de Progresso:** O script salva uma cópia local (`leads_local.xlsx`) marcando os leads como "enviado" ou "erro", permitindo que você retorne de forma segura e copie os dados atualizados para sua nuvem.
- **Tratamento de Dados e Sistema Anti-Spam:** 
  - Limpeza automática de máscaras de telefone (remove espaços, hifens e parênteses).
  - Rotação aleatória de saudações conforme o horário do dia.
  - Intervalos de tempo (delays) variáveis (25s a 54s) entre os envios para evitar bloqueios no WhatsApp.

## 📂 Estrutura Necessária no Google Sheets

Para que o bot funcione, certifique-se de que a planilha pública possua as seguintes colunas (não importando a ordem):

- `phone`: Contendo os números de telefone.
- `status`: Para registrar ou verificar se o contato já recebeu a mensagem (opcional criar previamente, o script cria localmente se não existir).

## 🚀 Como Executar

### 1. Pré-requisitos e Instalação
Certifique-se de ter o Python 3.12+ instalado.

Clone o repositório, crie um ambiente virtual e instale as dependências:
```bash
# Cria o ambiente virtual
python -m venv venv

# Ativa o ambiente (Windows)
venv\Scripts\activate

# Instala as bibliotecas de sistema
pip install -r requirements.txt

# Instala o navegador do Playwright
playwright install chromium
```

### 2. Uso

Basta rodar o arquivo principal:
```bash
python teste.py
```

O bot solicitará o ID da sua planilha do Google Sheets no terminal.
Na primeira execução, o Playwright abrirá uma janela do navegador exigindo que você escaneie o QR Code do WhatsApp Web. Nas próximas vezes, o login será feito automaticamente!
