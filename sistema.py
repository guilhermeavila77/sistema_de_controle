from tela_add import *
import PySimpleGUI as sg

layout = [
    [sg.Text("SISTEMA DE CONTROLE")],
    [sg.Text("graficoFat"), sg.Text("graficoLuc")],
    [sg.Button("ADD CLIENTE"), sg.Button("ALTERAR INFORMAÇÃO"),
     sg.Button("EXCLUIR CLIENTE")],
    [sg.Text("maiorfaturamento"), sg.Text(
        "maiorfaturamento"), sg.Text("maiorfaturamento")]
]

janela = sg.Window("SISTEMA DE CONTROLE", layout)

while True:
    evento, valores = janela.read()
    if evento == sg.WIN_CLOSED:
        break

    if evento == "ADD CLIENTE":
        tela_add()

janela.close()
