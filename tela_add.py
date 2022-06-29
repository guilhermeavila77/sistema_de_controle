import pandas as pd
from openpyxl import load_workbook
import PySimpleGUI as sg


def tela_add():

    layoutAdd = [
        [sg.Text("ADICIONAR CLIENTE")],
        [sg.InputText(key='nome', size=(15, 1)), sg.Text("NOME")],
        [sg.InputText(key='estado', size=(15, 1)), sg.Text("ESTADO")],
        [sg.InputText(key='pais', size=(15, 1)), sg.Text("PAIS")],
        [sg.InputText(key='faturamento', size=(15, 1)),
         sg.Text("FATURAMENTO")],
        [sg.InputText(key='despesas', size=(15, 1)), sg.Text("DESPESAS")],
        [sg.InputText(key='data', size=(15, 1)), sg.Text("DATA")],
        [sg.Button("ADICIONAR"), sg.Button("CANCELAR")]
    ]

    janelaAdd = sg.Window("SISTEMA DE CONTROLE", layoutAdd)

    while True:
        evento, valores = janelaAdd.read()
        if evento == sg.WIN_CLOSED or evento == "CANCELAR":
            break

        if evento == "ADICIONAR":
            nome = str(valores['nome'])
            estado = str(valores['estado'])
            pais = str(valores['pais'])
            faturamento = int(valores['faturamento'])
            despesas = int(valores['despesas'])
            data = str(valores['data'])
            # Adiciona a linha
            atualizartabela = load_workbook("Base de de dados (1).xlsx")
            aba_ativa = atualizartabela.active
            lucro = 0
            aba_ativa.append(
                [nome, estado, pais, faturamento, despesas, lucro, data])
            # Salva a planilha e atualizar o data frame
            atualizartabela.save("Base de de dados (1).xlsx")
            tabela = pd.read_excel("Base de de dados (1).xlsx")
            # Calcula o lucro
            tabela["LUCRO"] = tabela["FATURAMENTO"] - tabela["DESPESAS"]
            # Atualizar e salvar o data frame
            tabela.to_excel("Base de de dados (1).xlsx", index=False)
            tabela = pd.read_excel("Base de de dados (1).xlsx")

            # pop  out   e   fechaar a janela   e  tentar aa variavel caminho de novo

    janelaAdd.close()
