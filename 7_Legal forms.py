import requests
import time
import pandas as pd

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0'}
okveds = ['49.3', '49.4', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2']
#'49.1', '49.2',

for okved in okveds:
    inn_list = pd.read_excel('3.0_BFO_sample.xlsx', sheet_name=okved)
    inn_list = list(inn_list.columns)
    inn_list.pop(0)
    inn_list.pop()
    inn_list = list(map(int, inn_list))

    excel = pd.read_excel('2_BFO_of_List_companies.xlsx', sheet_name=okved, usecols=['id', 'inn'], dtype=str)
    excel = excel[excel['inn'] != 'нет данных']
    excel = excel.astype({'inn': 'int64'})

    id_list = []
    for inn in inn_list:
        id_list.append(excel.loc[excel.query('inn == @inn').index[0], 'id'])

    data = pd.DataFrame({'id': [], 'legal_form': []})
    x = 1
    for org in id_list:
        elem = requests.get(f'https://bo.nalog.ru/nbo/organizations/{org}', headers=headers).json()
        data.loc[len(data.index)] = [org, elem['okopf']['name']]
        print(f'{okved} - {x} готов')
        x += 1
        time.sleep(1)

    data = data.value_counts(subset='legal_form')
    print(data)

    with pd.ExcelWriter('7.0_Legal forms.xlsx', mode='a') as writer:
        data.to_excel(writer, sheet_name=okved)
