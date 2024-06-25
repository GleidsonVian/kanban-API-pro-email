# Trello to Excel Exporter and Email Sender

Este projeto consiste em um script Python que exporta dados do Trello para um arquivo Excel e envia um e-mail com este arquivo anexado. O objetivo é automatizar o processo de coleta de informações de tarefas e seu compartilhamento via e-mail.

## Requisitos

- Python 3.x
- Trello API Key e Token (https://trello.com/power-ups/admin/)
- Microsoft Outlook instalado e configurado

## Instalação

1. Clone o repositório ou baixe os arquivos do projeto.

2. Navegue até o diretório do projeto e instale as dependências usando o `pip`:

    ```sh
    pip install -r requirements.txt
    ```

3. Substitua as credenciais da API do Trello no código `chamada_api` com sua própria API Key e Token:

    ```python
    api_key = 'sua_api_key'
    token = 'seu_token'
    ```

## Uso

1. Execute o script Python:

    ```sh
    python bot.py
    ```

2. O script irá:
   - Conectar-se à API do Trello.
   - Coletar os dados de tarefas dos quadros e listas especificados.
   - Exportar os dados para um arquivo Excel (`tarefas_kanban.xlsx`).
   - Enviar um e-mail com o arquivo Excel anexado para o destinatário especificado.

## Estrutura do Código

### Funções Principais

- `chamada_api()`: Conecta-se à API do Trello usando a API Key e o Token fornecidos.
- `loop_de_cards(cards)`: Itera sobre uma lista de cartões do Trello e retorna uma lista de tuplas contendo o nome do cartão e a data de criação.
- `exportar_para_excel()`: Coleta dados de várias listas no Trello, cria um DataFrame e exporta os dados para um arquivo Excel.
- `enviar_email()`: Envia um e-mail com o arquivo Excel anexado.

## Configurações

- Modifique o índice das listas em `todas_as_listas` conforme necessário para corresponder às suas listas do Trello.
- Atualize o endereço de e-mail do destinatário em `email.To` para o endereço desejado.
