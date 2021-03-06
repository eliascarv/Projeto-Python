from openpyxl import load_workbook
from datetime import date
from statistics import mean, stdev, median
import pandas as pd
import requests

excel_file = 'sudocontrol-original.xlsx' 
wb = load_workbook(excel_file, data_only = True)
sh = wb['Abraçadeira (2)']
wb.sheetnames
sh['A1410'].fill.start_color.index

color_list = []
for i in range(7, sh.max_row):
    color_index = sh['A'+str(i)].fill.start_color.index
    if color_index != '00000000':
        color_list.append(color_index)


numrows = len(color_list)
wbpd = pd.read_excel(excel_file, skiprows=5, sheet_name='Abraçadeira (2)', converters={'Identif Compra': str,'Pregão':str,'Cód. Unidade':str})
wbpd = wbpd[0:numrows]
color_col = [1 if color == 'FFE2F0D9' else 0 for color in color_list]
wbpd['Ativo'] = color_col

wbpd["Número do Pregão"] = wbpd['Cód. Unidade'] + "000" + wbpd['Pregão']

quant_item = []

# def pesquisar_quant_item(num_item, itens_lista, lista):
#     for item in itens_lista:
#         item_title = item['_links']['self']['title']
#         if int(item_title.split()[1].split(":")[0]) == num_item:
#             quant = int(item['quantidade_item'])
#             lista.append(quant)
#             return True
    
#     return False


# for index, row in wbpd.iterrows():
#     if row['Ativo'] == 1:
#         if row['Modalidade Compra'] == 'Dispensa de Licitação':
#             quant_item.append("Erro: Dispensa de Licitação")

#         try:
#             num_item = int(row['Item'])

#             response = requests.get('http://compras.dados.gov.br/pregoes/doc/pregao/{}/itens.json'.format(row["Número do Pregão"]))
#             resdict = response.json()

#             itens = resdict['_embedded']['pregoes']

#             result_pesquisa = pesquisar_quant_item(num_item, itens, quant_item)

#             if result_pesquisa == False:
#                 response = requests.get('http://compras.dados.gov.br/pregoes/doc/pregao/{}/itens.json?offset=500'.format(row["Número do Pregão"]))
#                 resdict = response.json()

#                 itens = resdict['_embedded']['pregoes']

#                 result_pesquisa = pesquisar_quant_item(num_item, itens, quant_item)

#         except:
#             quant_item.append("Erro: Desconhecido")
            
#     else:
#         quant_item.append('Item não ativo')


for index, row in wbpd.iterrows():
    if row['Ativo'] == 1:
        if row['Modalidade Compra'] == 'Dispensa de Licitação':
            quant_item.append("Erro: Dispensa de Licitação")

        try:
            num_item = int(row['Item'])
            quant = 0

            response = requests.get('http://compras.dados.gov.br/pregoes/doc/pregao/{}/itens.json'.format(row["Número do Pregão"]))
            resdict = response.json()

            itens = resdict['_embedded']['pregoes']

            for item in itens:
                item_title = item['_links']['self']['title']
                if int(item_title.split()[1].split(":")[0]) == num_item:
                    quant = int(item['quantidade_item'])
                    quant_item.append(quant)

            if quant == 0:
                response = requests.get('http://compras.dados.gov.br/pregoes/doc/pregao/{}/itens.json?offset=500'.format(row["Número do Pregão"]))
                resdict = response.json()

                itens = resdict['_embedded']['pregoes']

                for item in itens:
                    item_title = item['_links']['self']['title']
                    if int(item_title.split()[1].split(":")[0]) == num_item:
                        quant = int(item['quantidade_item'])
                        quant_item.append(quant)
        
        except:
            quant_item.append("Erro: Desconhecido")
            
    else:
        quant_item.append('Item não ativo')


wbpd['Quant Item'] = quant_item

wbpd['Quant Item'] = quant_item

anexos_links = []

for index, row in wbpd.iterrows():
    anexos_links.append("https://comprasnet.gov.br/livre/pregao/anexosDosItens.asp?uasg={}&numprp={}&prgcod=863000".format(row['Cód. Unidade'], row['Pregão']))

wbpd['Anexos Pregão'] = anexos_links

meses = {
    "Jan": 1, "Fev": 2, "Mar": 3, "Abr":  4, "Mai":  5, "Jun":  6, 
    "Jul": 7, "Ago": 8, "Set": 9, "Out": 10, "Nov": 11, "Dez": 12
}

mes_ano = wbpd['Mês Resultado Compra']

mes = [meses[i.split()[0]] for i in mes_ano]
ano = [int(i.split()[1]) for i in mes_ano]

wbpd['Mês'] = mes
wbpd['Ano'] = ano

mes_atual = date.today().month
ano_atual = date.today().year

dentro_do_periodo = []
for index, row in wbpd.iterrows():
    if row['Ativo'] == 1:
        if row['Ano'] == (ano_atual - 1) and row['Mês'] in range(mes_atual + 1, 13):
           dentro_do_periodo.append(1)
        elif row['Ano'] == ano_atual and row['Mês'] in range(1, mes_atual + 1):
            dentro_do_periodo.append(1)
        else:
            dentro_do_periodo.append(0)
    else:
        dentro_do_periodo.append('Item não ativo')


wbpd['Dentro do Período'] = dentro_do_periodo

valores = []
for index, row in wbpd.iterrows():
    if row['Ativo'] == 1 and row['Dentro do Período'] == 1:
        valores.append(row['Valor'])

media = mean(valores)
desvio = stdev(valores)
mediana = median(valores)
coeficiente = desvio / media
preco = mediana if coeficiente > 0.25 else media

relatorio_lista = []
relatorio_lista.append(['Abraçadeira (2)', media, desvio, coeficiente, mediana, preco])

relatorio = pd.DataFrame(
    relatorio_lista, 
    columns = ['Item', 'Média', 'Desvio Padão', 'Coeficiente', 'Mediana', 'Preço']
)

wbpd.to_excel('resultado_consultas.xlsx', index = False)