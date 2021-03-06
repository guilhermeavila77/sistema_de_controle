import pandas as pd
from openpyxl import load_workbook
import PySimpleGUI as sg

global caminho
caminho = "Base de de dados (1).xlsx"
global tabela1
tabela1 = pd.read_excel(caminho)
global lista
lista = tabela1['NOME']
global infos
infos = ['NOME', 'ESTADO', 'FATURAMENTO', 'DESPESAS', 'DATA',
         'CPF/CNPJ', 'PJ/PF', 'TELEFONE', 'E-MAIL', 'CLASSIFICAÇÃO']


def tela_add():

    layoutAdd = [
        [sg.Text("ADICIONAR CLIENTE")],
        [sg.InputCombo(("PESSOA FISICA", "PESSOA JURIDICA"), default_value="PJ/PF", key='PJ/PF', size=(
            15, 1)), sg.InputCombo(("A", "B", "C"), default_value="CLASSE", key='classe', size=(15, 1))],
        [sg.InputText(key='nome', size=(15, 1)), sg.Text("NOME")],
        [sg.InputText(key='estado', size=(15, 1)), sg.Text("ESTADO")],
        [sg.InputText(key='faturamento', size=(15, 1)),
         sg.Text("FATURAMENTO")],
        [sg.InputText(key='despesas', size=(15, 1)), sg.Text("DESPESAS")],
        [sg.InputText(key='data', size=(15, 1)), sg.Text("DATA")],
        [sg.InputText(key='CPF/CNPJ', size=(15, 1)), sg.Text("CPF/CNPJ")],
        [sg.InputText(key='telefone', size=(15, 1)), sg.Text("TELEFONE")],
        [sg.InputText(key='e-mail', size=(15, 1)), sg.Text("E-MAIL")],
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
            faturamento = int(valores['faturamento'])
            despesas = int(valores['despesas'])
            data = str(valores['data'])
            cnpj = str(valores['CPF/CNPJ'])
            pjpf = str(valores['PJ/PF'])
            tel = str(valores['telefone'])
            mail = str(valores['e-mail'])
            cliente = str(valores['classe'])
            # Adiciona a linha
            atualizartabela = load_workbook(caminho)
            aba_ativa = atualizartabela.active
            lucro = 0
            aba_ativa.append(
                [nome, estado, faturamento, despesas, lucro, data, cnpj, pjpf, tel, mail, cliente])
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
    btn_layout = [
        [sg.Button("EXCLUIR"), sg.Button("VOLTAR")]
    ]
    main_layout = [
        [sg.Text("EXCLUIR CLIENTE")],
        [sg.InputOptionMenu(
            values=lista, default_value='CLIENTE', size=(10, 10), key='excluir')]
    ]
    layoutTable = [
        [sg.Column(main_layout, element_justification='c')],
        [sg.Column(btn_layout, element_justification='c')]

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


def alt():
    contAlt = 0
    layoutAlt = [
        [sg.Text('ALTERAR INFORMAÇÃO')],
        [sg.InputText(key='linhaAlt', size=(30, 10)),
         sg.Text("CLIENTE DA ALTERAÇÃO")],
        [sg.InputOptionMenu(values=infos, default_value='COLUNA', size=(
            25, 10), key='altcolum'), sg.Text("INFORMAÇÃO DA ALTERAÇÃO")],
        [sg.InputText(key='alt', size=(30, 10)), sg.Text("ALTERAÇÃO")],
        [sg.Button('ALTERAR')]
    ]

    janelaAlt = sg.Window("ALTERAR INFORMAÇÃO", layoutAlt)

    while True:
        evento, valores = janelaAlt.read()
        if evento == sg.WIN_CLOSED:
            break

        if evento == "ALTERAR":
            tabela = pd.read_excel(caminho)
            linhaAlt = valores['linhaAlt']
            for verificar in tabela.itertuples():
                altCliente = verificar.NOME
                contAlt = contAlt + 1
                if altCliente == linhaAlt:
                    break

            colunaAlteração = valores['altcolum']
            print(contAlt)
            contAlt = contAlt - 1
            print(contAlt)
            new = valores['alt']
            if colunaAlteração == "FATURAMENTO" or colunaAlteração == "DESPESAS":
                new = int(new)
            tabela.loc[contAlt, colunaAlteração] = new
            # Salva a tabela
            tabela.to_excel(caminho, index=False)
            tabela = pd.read_excel(caminho)
            sg.popup('VALOR ALTERADO!')
            janelaAlt.close()

    janelaAlt.close()


def result_filt(classefilt, pessoafilt):
    titulo = "LISTA DE CLIENTES (" + classefilt + ") " + pessoafilt
    tab = tabela1.loc[(tabela1['CLASSIFICAÇÃO'] == classefilt) & (
        tabela1['PJ/PF'] == pessoafilt)]
    tab_layout = [
        [sg.Text(titulo)],
        [sg.Listbox(tab['NOME'], size=(20, 20))]
    ]
    janelaFilt = sg.Window("BUSCA DE CLIENTES",
                           tab_layout, element_justification='c')

    while True:
        evento, valores = janelaFilt.read()
        if evento == sg.WIN_CLOSED:
            break
    janelaFilt.close()


def buscar():
    tab = tabela1
    filt_layout = [
        [sg.InputCombo(('A', 'B', 'C'),
                       default_value='CLASSIFICAÇÃO', size=(15, 1), key='classes')],
        [sg.InputCombo(('PESSOA FISICA', 'PESSOA JURIDICA'),
                       default_value='PJ/PF', size=(15, 1), key='pessoa')],
        [sg.Button('BUSCAR')]
    ]
    result_layout = [
        [sg.Text("LISTA CLIENTE")],
        [sg.Listbox(tab['NOME'], size=(20, 20))]
    ]
    layoutBusca = [
        [sg.Column(filt_layout, element_justification='c'),
         sg.Column(result_layout, element_justification='c')]
    ]

    janelaTable = sg.Window("BUSCA DE CLIENTES", layoutBusca)

    while True:
        evento, valores = janelaTable.read()
        if evento == sg.WIN_CLOSED or evento == "VOLTAR":
            break
        if evento == "BUSCAR":
            classe = str(valores['classes'])
            pessoa = str(valores['pessoa'])
            result_filt(classe, pessoa)

    janelaTable.close()
