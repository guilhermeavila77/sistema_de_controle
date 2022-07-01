import pandas as pd
from openpyxl import load_workbook
import PySimpleGUI as sg

global caminho
caminho = "Base de de dados (1).xlsx"
global tabela
tabela = pd.read_excel(caminho)
global lista 
lista = tabela['NOME']



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
            atualizartabela = load_workbook(caminho)
            aba_ativa = atualizartabela.active
            lucro = 0
            aba_ativa.append(
                [nome, estado, pais, faturamento, despesas, lucro, data])
            # Salva a planilha e atualizar o data frame
            atualizartabela.save(caminho)
            tabela = pd.read_excel(caminho)
            # Calcula o lucro
            tabela["LUCRO"] = tabela["FATURAMENTO"] - tabela["DESPESAS"]
            # Atualizar e salvar o data frame
            tabela.to_excel(caminho, index=False)
            tabela = pd.read_excel(caminho)

            sg.popup('CLIENTE ADICIONADO!')
            janelaAdd.close()

    janelaAdd.close()


def table():
    cont = 0
    layoutTable = [
        [sg.Text("LISTA DE CLIENTES")],
        [sg.InputOptionMenu(values=lista, default_value='CLIENTE', size=(10,10),key='excluir')],
        [sg.Button("EXCLUIR")],
        [sg.Button("VOLTAR")]
    ]

    janelaTable = sg.Window("LISTA CLIENTES", layoutTable)

    while True:
        evento, valores = janelaTable.read()
        if evento == sg.WIN_CLOSED or evento == "VOLTAR":
            break

        if evento == "EXCLUIR":
            tabela = pd.read_excel(caminho)
            clienteExcluir = valores['excluir']
            for cliente in tabela.itertuples():
                rodagemCliente = cliente.NOME
                cont = cont + 1
                if rodagemCliente == clienteExcluir:
                    break
            
            # Exclui a linha da tabela
            cont = cont - 1
            tabela = tabela.drop(cont)
            # Salva a tabela
            tabela.to_excel(caminho, index=False)
            tabela = pd.read_excel(caminho)
            sg.popup('CLIENTE EXCLUIDO!')
            janelaTable.close()
            

    janelaTable.close()
