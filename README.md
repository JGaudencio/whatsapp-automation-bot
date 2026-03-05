# WhatsApp Lead Automator 🚀

Este projeto é um bot de automação em Python desenvolvido para facilitar o contato com leads. O script lê uma base de dados em Excel, realiza o tratamento dos números de telefone e automatiza o envio de mensagens personalizadas via WhatsApp Web.

## 🛠️ Funcionalidades

- **Integração com Excel:** Leitura automatizada de contatos a partir de arquivos `.xlsx`.
- **Tratamento de Dados:** Limpeza automática de máscaras de telefone (remove espaços, hifens e parênteses).
- **Sistema Anti-Spam:**
  - Rotação aleatória de saudações.
  - Intervalos de tempo (delays) variáveis entre os envios para evitar bloqueios.
- **Fechamento Automático:** Gerencia as abas do navegador para evitar consumo excessivo de memória RAM.

## 📂 Estrutura da Planilha

Para que o bot funcione, crie um arquivo chamado `leads.xlsx` na raiz do projeto com a seguinte estrutura:

| telefone |
| :--- |
| +55 82 99999-9999 |
| 5582988887777 |

*Obs: A primeira linha deve ser obrigatoriamente "telefone".*

## 🚀 Como Executar

### 1. Pré-requisitos
Certifique-se de ter o Python 3.12+ instalado e o WhatsApp Web logado no seu navegador padrão.

### 2. Instalação
Clone o repositório e instale as dependências:
```bash
pip install -r requirements.txt
