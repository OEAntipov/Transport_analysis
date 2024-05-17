import time
import requests
import pandas as pd


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0'}
okveds = ['49.4']
#, '49.1', '49.2', '49.3', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2'

for okved in okveds:
    for year in ['previous', 'beforePrevious']:
        excel = pd.read_excel('2_BFO_of_List_companies.xlsx', sheet_name=okved)
        sample = excel.query('(current1600 > 400000) | (current2110 > 800000)').reset_index(drop=True)
        orgs = sample['id']
        excel_orgs = sample[['id', 'inn']]
        excel_orgs = excel_orgs.set_index('id')

        for org in orgs:
            elem = requests.get(f'https://bo.nalog.ru/nbo/organizations/{org}/bfo/',headers=headers).json()

            try:
                excel_orgs.loc[org, 'inn'] = elem[0]['organizationInfo']['inn']
            except:
                excel_orgs.loc[org, 'inn'] = 'нет данных'
            excel_orgs.loc[org, 'period'] = elem[0]['period']
            excel_orgs.loc[org, 'knd'] = elem[0]['knd']

            if 'balance' in elem[0]['correction'] and 'financialResult' in elem[0]['correction']:
                for i in elem[0]['correction']['balance']:
                    if year in i:
                        excel_orgs.loc[org, i] = elem[0]['correction']['balance'][i]
                for i in elem[0]['correction']['financialResult']:
                    if year in i:
                        excel_orgs.loc[org, i] = elem[0]['correction']['financialResult'][i]
                print(f'По {okved} выгружен {org} - {year}')
            else:
                excel_orgs.loc[org, 'knd'] = 'данные отсутствуют'
            time.sleep(1)

        with pd.ExcelWriter(f'4_BFO_List_{year}.xlsx', mode='a') as writer:
            excel_orgs.to_excel(writer, sheet_name=okved)
