import requests
import time
import pandas as pd


headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:125.0) Gecko/20100101 Firefox/125.0',
           'Referer': 'https://bo.nalog.ru/search?allFieldsMatch=false&okved=49.5&period=2023&page=1'}

okveds = ['49.4']
#'49.1', '49.2', '49.3', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2'
for okved in okveds:

    componies = pd.DataFrame({'id1': [], 'id2': []})
    if okved != '49.4':
        page = 0
        isnot_all_pages = True
        while isnot_all_pages:
            elem = requests.get(f'https://bo.nalog.ru/advanced-search/organizations/?allFieldsMatch=false'
                                f'&okved={okved}'
                                f'&period=2023&page={page}',
                                headers=headers).json()

            if page == 0:
                for i in elem['content']:
                    componies.loc[len(componies.index)] = [i['id'], i['id']]
                print(f'ОКВЭД {okved}, страница {page} - готово')
                page += 1
            elif elem['content'] != []:
                elem2 = requests.get(f'https://bo.nalog.ru/advanced-search/organizations/?allFieldsMatch=false'
                                     f'&okved={okved}'
                                     f'&period=2023'
                                     f'&page={page}',
                                     headers=headers).json()
                while elem['content'] == elem2['content']:
                    elem2 = requests.get(f'https://bo.nalog.ru/advanced-search/organizations/?allFieldsMatch=false'
                                         f'&okved={okved}'
                                         f'&period=2023'
                                         f'&page={page}',
                                         headers=headers).json()
                    time.sleep(1)
                for i, j in zip(elem['content'], elem2['content']):
                    componies.loc[len(componies.index)] = [i['id'], j['id']]

                print(f'ОКВЭД {okved}, страница {page} - готово')
                page += 1
            else:
                print('Дубликатов', componies.duplicated().sum())
                with pd.ExcelWriter('1_List_of_companies.xlsx', mode='a') as writer:
                    componies.to_excel(writer, sheet_name=okved)
                print(f'ОКВЭД {okved} - готов на {page} страниц')
                isnot_all_pages = False

            time.sleep(1)
    else:
        componies = pd.DataFrame({'id': []})
        for char1 in range(100):
            if char1 < 10:
                char1 = f'0{char1}'
            else:
                char1 = f'{char1}'

            for char2 in range(100):
                if char2 < 10:
                    char2 = f'0{char2}'
                else:
                    char2 = f'{char2}'
                inn = f'{char1}{char2}'

                page = 0
                isnot_all_pages = True
                comp_dupl = pd.DataFrame({'id1': [], 'id2': []})
                items = 0
                while isnot_all_pages:
                    elem = requests.get(f'https://bo.nalog.ru/advanced-search/organizations/?allFieldsMatch=false'
                                        f'&inn={inn}'
                                        f'&okved={okved}'
                                        f'&period=2023&page={page}',
                                        headers=headers).json()

                    if page == 0 and elem['content'] != []:
                        for i in elem['content']:
                            comp_dupl.loc[len(comp_dupl.index)] = [i['id'], i['id']]
                        print(f'ОКВЭД {okved}, страница {page} - готово')
                        page += 1
                    elif elem['content'] != []:
                        elem2 = requests.get(f'https://bo.nalog.ru/advanced-search/organizations/?allFieldsMatch=false'
                                             f'&inn={inn}'
                                             f'&okved={okved}'
                                             f'&period=2023&page={page}',
                                             headers=headers).json()
                        y = 0
                        while elem['content'] == elem2['content'] and y <= 10:
                            elem2 = requests.get(f'https://bo.nalog.ru/advanced-search/organizations/?allFieldsMatch=false'
                                                 f'&inn={inn}'
                                                 f'&okved={okved}'
                                                 f'&period=2023&page={page}',
                                                 headers=headers).json()
                            y += 1
                            time.sleep(1)
                        for i, j in zip(elem['content'], elem2['content']):
                            comp_dupl.loc[len(comp_dupl.index)] = [i['id'], j['id']]
                        print(f'ОКВЭД {okved}, страница {page} - готово')
                        page += 1

                    else:
                        isnot_all_pages = False
                        items = elem['totalElements']

                    time.sleep(1)
                print(f'ОКВЭД {okved} по ИНН {inn} - готов')
                ser1 = comp_dupl['id1']
                ser2 = comp_dupl['id2']
                comp_no_dupl = pd.concat([ser1, ser2], ignore_index=True)
                comp_no_dupl = comp_no_dupl.drop_duplicates(ignore_index=True)
                print(f'Количество уникальных по {inn} - {comp_no_dupl.count()} из {items}')
                componies = pd.concat([componies, comp_no_dupl], ignore_index=True)
                componies = componies.drop_duplicates(ignore_index=True)
                print(f'На момент {inn} уникальных - {len(componies.index)}')

        with pd.ExcelWriter('1_List_of_companies.xlsx', mode='a') as writer:
            componies.to_excel(writer, sheet_name='49.4_inn')