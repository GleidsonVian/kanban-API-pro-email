from trello import TrelloClient
import pandas as pd
import win32com.client as win32

def chamada_api():
    api_key = 'sua chave api'
    token = 'seu token'

    client = TrelloClient(
        api_key=api_key,
        token=token)
    return client

def loop_de_cards(cards):
    lista = []
    for card in cards:
        lista.append((card.name, card.card_created_date))
    return lista

# Função para exportar dados do Trello para um arquivo Excel
def exportar_para_excel():
    client = chamada_api()

    todos_os_quadros = client.list_boards()
    meu_quadro = todos_os_quadros[0]
    todas_as_listas = meu_quadro.list_lists()

    a_fazer = todas_as_listas[2]
    cards_a_fazer = a_fazer.list_cards()

    em_andamento = todas_as_listas[3]
    cards_em_andamento = em_andamento.list_cards()

    fase_de_teste = todas_as_listas[5]
    cards_fase_de_teste = fase_de_teste.list_cards()

    concluido = todas_as_listas[6]
    tarefas_concluidas = concluido.list_cards()

    lista_a_fazer = loop_de_cards(cards_a_fazer)
    lista_em_andamento = loop_de_cards(cards_em_andamento)
    lista_fase_de_teste = loop_de_cards(cards_fase_de_teste)
    lista_concluido = loop_de_cards(tarefas_concluidas)

    # Unir todas as listas em uma só
    todas_as_tarefas = lista_a_fazer + lista_em_andamento + lista_fase_de_teste + lista_concluido

    # Criar um DataFrame com os dados
    df = pd.DataFrame(todas_as_tarefas, columns=['Nome do Card', 'Data de Criação'])

    # Exportar para Excel
    df.to_excel('tarefas_kanban.xlsx', index=False)
    print("Arquivo Excel 'tarefas_kanban.xlsx' criado com sucesso!")

# Função para enviar email com o arquivo Excel anexado
def enviar_email():
    outlook = win32.Dispatch('outlook.application')
    email = outlook.CreateItem(0)
    email.To = 'gleidson.testes1@outlook.com'
    email.Subject = 'Email de atualização do Trello'
    email.HTMLBody = """
    <p>Olá a todos,</p>
    <p>Segue uma planilha sobre as atualizações do Trello, por favor se atentem ao prazo.</p>
    <p></p>
    """

    # Verificar se o arquivo Excel existe antes de anexar e enviar
    anexo = r"C:\Users\gleid\pasta4\Python\varredura kanban\tarefas_kanban.xlsx"
    try:
        email.Attachments.Add(anexo)
    except Exception as e:
        print(f"Erro ao anexar arquivo: {e}")
    else:
        email.Send()
        print("Email enviado com sucesso!")

# Executar as funções
exportar_para_excel()
enviar_email()
