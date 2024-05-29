import requests
from bs4 import BeautifulSoup
import time
import pandas as pd

headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0'}
okveds = ['49.3', '49.4']
#'49.1', ,'49.2', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2'

for okved in okveds:
    inn_list = pd.read_excel('3.0_BFO_sample.xlsx', sheet_name=okved)
    inn_list = list(inn_list.columns)
    inn_list.pop(0)
    inn_list.pop()
    #print(inn_list)
    print(f'{okved} начат')

    df = pd.DataFrame()
    for inn in inn_list:
        try:
            query = requests.get(f'https://checko.ru/search/quick_tips?query={inn}', headers=headers).json()
            try:
                html = BeautifulSoup(query[0]['content'], features="html.parser")
            except:
                df[inn] = ''
                print(f'{inn} нет на сайте')
                continue
            comp_link = html.find('a').get('href')
            time.sleep(1)
            affiliated_comp = requests.get(f'https://checko.ru{comp_link}?extra=connections&type=founded', headers=headers)
            html = BeautifulSoup(affiliated_comp.content, features="html.parser")
            if html.find('p', {'class': 'mt-8 mb-48'}) is not None:
                df[inn] = ''
                print(f'{inn} нет связей')
                continue
            time.sleep(1)
            #print(html)

            df[inn] = ''
            x = 0
            for a in html.find_all('a', {'class': 'link'}, href=True):
                if len(a['class']) == 1:
                    href = a['href']
                    #print(href)
                    elem = requests.get(f'https://checko.ru{href}', headers=headers)
                    html = BeautifulSoup(elem.content, features="html.parser")
                    inn_aff = html.find('strong', {'id': 'copy-inn'}).text
                    #print(inn.text)
                    time.sleep(1)

                    df.loc[x, inn] = inn_aff
                    x += 1
                    print(f'{inn} выгружен')
        except:
            df[inn] = ''
            print(f'{inn} ошибка')
            continue
    #print(df)
    with pd.ExcelWriter('8.0_Affiliates.xlsx', mode='a') as writer:
        df.to_excel(writer, sheet_name=okved)
