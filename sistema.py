from tela_add import *
import PySimpleGUI as sg

# cria a interface grafica
sg.theme('LightBrown6')

collun_center = [
    [sg.Text("SISTEMA DE CONTROLE")],
    [sg.Text("graficoFat"), sg.Text("graficoLuc")],
    [sg.Button("ADD CLIENTE"), sg.Button("ALTERAR INFORMAÇÃO"),
     sg.Button("EXCLUIR CLIENTE")],
    [sg.Text("maiorfaturamento"), sg.Text(
        "maiorfaturamento"), sg.Text("maiorfaturamento")]
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

janela.close()
