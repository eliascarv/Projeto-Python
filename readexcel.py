from openpyxl import load_workbook
import pandas as pd
import requests

excel_file = 'sudocontrol-original.xlsx' 
wb = load_workbook(excel_file, data_only = True)
sh = wb['Abraçadeira (2)']
wb.sheetnames
sh['A1'].fill.start_color.index

color_list = []
for i in range(7, sh.max_row):
    color_index = sh['A'+str(i)].fill.start_color.index
    if color_index != '00000000':
        color_list.append(color_index)


numrows = len(color_list)
wbpd = pd.read_excel(excel_file, skiprows=5, sheet_name='Abraçadeira (2)', converters={'Pregão':str,'Cód. Unidade':str})
wbpd = wbpd[0:numrows]
color_col = [1 if i == 'FFE2F0D9' else 0 for i in color_list]
wbpd['Ativo'] = color_col

wbpd["Número do Pregão"] = wbpd['Cód. Unidade'] + "000" + wbpd['Pregão']

quant_item = []

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
        

len(quant_item)

wbpd['Quant Item'] = quant_item

anexos_links = []
for index, row in wbpd.iterrows():
    anexos_links.append("https://comprasnet.gov.br/livre/pregao/anexosDosItens.asp?uasg={}&numprp={}&prgcod=863000".format(row['Cód. Unidade'], row['Pregão']))

wbpd['Anexos Pregão'] = anexos_links

