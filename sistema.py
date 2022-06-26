# Importar bibliotecas
import pandas as pd
from openpyxl import workbook, load_workbook, worksheet
import plotly.express as px

# Variaveis globais

caminho = "/content/drive/MyDrive/Curso analise de dados/Base de de dados.xlsx"
tabela = pd.read_excel(caminho)


def add_linha():
    nome = input(str('Nome: '))
    estado = input(str('Estado: '))
    pais = input(str('Pais: '))
    faturamento = input('Faturamento: ')
    despesas = input('Despesas: ')
    data = input(str('Data: '))

    # Formata as variaveis inteiras
    faturamento = int(faturamento)
    despesas = int(despesas)

    # Adiciona a linha
    atualizartabela = load_workbook(caminho)
    aba_ativa = atualizartabela.active
    lucro = 0
    aba_ativa.append([nome, estado, pais, faturamento, despesas, lucro, data])
    # Salva a planilha e atualizar o data frame
    atualizartabela.save(caminho)
    tabela = pd.read_excel(caminho)


def calc_lucro():
    tabela["LUCRO"] = tabela["FATURAMENTO"] - tabela["DESPESAS"]
    # Atualizar e salvar o data frame
    tabela.to_excel(caminho, index=False)
    tabela = pd.read_excel(caminho)


def alteracao():

    # Altera uma informação solicitada

    # Pede a coluna e a linha
    # Pede a coluna e passar o nome
    alteracaocolumn = input('Coluna de alteração: ')
    # Pede a linha e passar o numero da linha
    alteracaorow = input('Linha de alteração: ')
    # Formata a linha
    alteracaorow = int(alteracaorow)
    # Pede a alteração
    newvalue = input('Alteração: ')
    tabela.loc[alteracaorow, alteracaocolumn] = newvalue
    # Salva a planilha
    tabela.to_excel(caminho, index=False)
    tabela = pd.read_excel(caminho)


def excluir():
    # Escluir uma linha

    # Pede e formata a linha que deseja ser excluida
    excluir = input('Linha para excluir: ')
    excluir = int(excluir)

    # Mostra a linha excluida
    print(excluir)

    # Exclui a linha da tabela
    tabela = tabela.drop(excluir)

    # Salva a tabela
    tabela.to_excel(caminho, index=False)
    tabela = pd.read_excel(caminho)


def mostar_resultados():
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

    print(f'{txtfaturamento}{namefat}{com}{fatanterior}')
    print(f'{txtdespesas}{namedesp}{com}{despanterior}')
    print(f'{txtlucro}{namelucro}{com}{lucroanterior}')


def graficos():
    # Plotar os graficos

    # Relação ESTADO x FATURAMENTO
    grafico = px.line(tabela, x="ESTADO", y="FATURAMENTO")
    grafico.show()

    # Relação ESTADO X LUCRO
    grafico = px.bar(tabela, x="ESTADO", y="LUCRO")
    grafico.show()
