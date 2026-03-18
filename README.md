# WhatsApp Lead Automator 🚀

Este projeto é um bot de automação em Python desenvolvido para facilitar o contato com leads. O script conecta-se a uma planilha pública do Google Sheets, baixa os contatos, realiza o tratamento dos números de telefone e automatiza o envio de mensagens personalizadas via WhatsApp Web através de um navegador controlado pelo **Playwright**.

## 🛠️ Novas Funcionalidades (v2.1)

- **Arquitetura Modular:** O código agora está dividido em múltiplos arquivos lógicos (`main.py`, `utils.py`, `gsheets.py` e `whatsapp_bot.py`), garantindo melhor manutenção e leitura.
- **Limitador de Envios:** Ao iniciar o bot, você pode escolher exatamente para quantos leads quer enviar mensagens naquela sessão. O bot fará a pausa automaticamente ao atingir o limite estipulado.
- **Auto-correção de Número (Nono Dígito):** Em caso de falha de envio para números com DDDs que frequentemente não salvam o 9º dígito (ex: DD 82 de Alagoas), o bot auto-corrige e tenta novamente de forma assíncrona.
- **Integração com Google Sheets na Nuvem:** O bot lê os dados diretamente de um link de compartilhamento público usando regex, sem necessidade de baixar a planilha manualmente.
- **Automação Segura em Segundo Plano:** Graças à migração para o `Playwright` com sessão persistente (Persistent Context), o bot abre um navegador próprio. Isso libera totalmente o seu mouse e teclado.

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

Basta rodar o arquivo principal unificado:
```bash
python main.py
```

1. O bot solicitará o Link ou ID da sua planilha do Google Sheets no terminal.
2. Em seguida, pedirá se você quer impor um **limite** na quantidade de envios daquela sessão (ou apertar ENTER para enviar a todos).
3. Na primeira execução, o Playwright abrirá uma janela do navegador exigindo que você escaneie o QR Code do WhatsApp Web. Nas próximas vezes, o login será feito automaticamente!

