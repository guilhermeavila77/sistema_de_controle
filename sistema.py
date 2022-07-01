from tela_add import *
import PySimpleGUI as sg

def atualizar():
    # Variaveis globais
    global caminho
    caminho = "Base de de dados (1).xlsx"
    global tabela
    tabela = pd.read_excel(caminho)
    global txtFat
    global txtDesp
    global txtLuc

    # Cria as variaveis utilizadas para a busca e as "seta" como zero

    # As buscas são

    # Maior faturamento, maior despesa, maior lucro

    fatanterior = 0
    despanterior = 0
    lucroanterior = 0

    # Laços de repetição que procuram as informações
    # A cada linha ele verifica se o valor da nova linha é maior que a linha anterior
    # Caso seja, esse valor é armazenado
    # Fazendo assim no final mostrar o maior valor

    for faturamento in tabela.itertuples():
        fat = faturamento.FATURAMENTO
        if fat > fatanterior:
            fatanterior = fat
            namefat = faturamento.NOME

    for despesas in tabela.itertuples():
        desp = despesas.DESPESAS
        if desp > despanterior:
            despanterior = desp
            namedesp = despesas.NOME

    for lucro in tabela.itertuples():
        luc = lucro.LUCRO
        if luc > lucroanterior:
            lucroanterior = luc
            namelucro = lucro.NOME

    # Mostar os resultados

    # Setar os textos

    txtfaturamento = "O maior faturamento foi do "
    txtdespesas = "A maior despesa foi do "
    txtlucro = "O maior lucro foi do "
    com = " com "

    txtFat = f'{txtfaturamento}{namefat}{com}{fatanterior}'
    txtDesp = f'{txtdespesas}{namedesp}{com}{despanterior}'
    txtLuc = f'{txtlucro}{namelucro}{com}{lucroanterior}'

atualizar()

# cria a interface grafica
sg.theme('LightBrown6')

collun_center = [
    [sg.Text("SISTEMA DE CONTROLE")],
    [sg.Text("graficoFat"), sg.Text("graficoLuc")],
    [sg.Button("ADD CLIENTE"), sg.Button("ALTERAR INFORMAÇÃO"),
     sg.Button("EXCLUIR CLIENTE")],
    [sg.Text(txtFat), sg.Text(txtDesp), sg.Text(txtLuc)]
]

layout = [[sg.VPush()],
          [sg.Push(), sg.Column(collun_center, element_justification='c'), sg.Push()],
          [sg.VPush()]]

janela = sg.Window("SISTEMA DE CONTROLE", layout)

while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED:
        break

    if evento == "ADD CLIENTE":
        tela_add()

    if evento == "EXCLUIR CLIENTE":
        table()

    if evento == "ALTERAR INFORMAÇÃO":
       alt()

janela.close()
