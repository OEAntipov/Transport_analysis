import time
import requests
import pandas as pd


# Данный код собирает отчетность компаний (ББ и ОФР) по списку, собранному в предыдущем файле

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0'}

okveds = ['49.1', '49.2', '49.3', '49.4', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2']

for okved in okveds:
    excel_orgs = pd.read_excel('1_List_of_companies.xlsx', sheet_name=okved, dtype=object)
    print(f'Считан файл {okved}')
    orgs = excel_orgs['id']
    excel_orgs = excel_orgs.set_index('id')
    x = 0
    for org in orgs:
        try:
            elem = requests.get(f'https://bo.nalog.ru/nbo/organizations/{org}/bfo/',headers=headers).json()
        except:
            break

        try:
            excel_orgs.loc[org, 'inn'] = elem[0]['organizationInfo']['inn']
        except:
            excel_orgs.loc[org, 'inn'] = 'нет данных'
        excel_orgs.loc[org, 'period'] = elem[0]['period']
        excel_orgs.loc[org, 'knd'] = elem[0]['knd']

        if 'balance' in elem[0]['correction'] and 'financialResult' in elem[0]['correction']:
            for i in elem[0]['correction']['balance']:
                if 'current' in i:
                    excel_orgs.loc[org, i] = elem[0]['correction']['balance'][i]
            for i in elem[0]['correction']['financialResult']:
                if 'current' in i:
                    excel_orgs.loc[org, i] = elem[0]['correction']['financialResult'][i]
            x += 1
            print(f'{x}. По {okved} выгружен {org}')
        else:
            excel_orgs.loc[org, 'knd'] = 'данные отсутствуют'
        time.sleep(1)

    with pd.ExcelWriter('2_BFO_of_List_companies.xlsx', mode='a') as writer:
        excel_orgs.to_excel(writer, sheet_name=okved)
