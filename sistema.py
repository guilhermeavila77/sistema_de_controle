from tela_add import *
import PySimpleGUI as sg
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


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


def draw_figure(canvas, figure):
    figure_canvas_agg = FigureCanvasTkAgg(figure, canvas)
    figure_canvas_agg.draw()
    figure_canvas_agg.get_tk_widget().pack(side='top', fill='both', expand=1)
    return figure_canvas_agg


# cria a interface grafica
sg.theme('DarkGrey')

collun_left = [
    [sg.Button("ADD CLIENTE", size=(20, 2))],
    [sg.Button("ALTERAR INFORMAÇÃO",  size=(20, 2))],
    [sg.Button("EXCLUIR CLIENTE",  size=(20, 2))],
    [sg.Button("BUSCA DE CLIENTES",  size=(20, 2))]
]

collun_center = [
    [sg.Text("SISTEMA DE CONTROLE")],
    [sg.Canvas(key='figCanvas')]
]

collun_right = [
    [sg.Text(txtFat, size=(20, 2))],
    [sg.Text(txtDesp, size=(20, 2))],
    [sg.Text(txtLuc, size=(20, 2))],
]

layout = [
    [sg.Column(collun_left, element_justification='l'),
     sg.Column(collun_center, element_justification='c'),
     sg.Column(collun_right, element_justification='c')]
]

janela = sg.Window("SISTEMA DE CONTROLE", layout, finalize=True,
                   resizable=True, element_justification='c')

# \\  -------- PYPLOT -------- //

# Make synthetic data
dataSize = 1000
xData = tabela['DATA']
yData = tabela['FATURAMENTO']
# make fig and plot
fig = plt.figure()
plt.bar(xData, yData, color='g')
# Instead of plt.show
draw_figure(janela['figCanvas'].TKCanvas, fig)

# \\  -------- PYPLOT -------- //

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
    if evento == "BUSCA DE CLIENTES":
        buscar()

janela.close()
