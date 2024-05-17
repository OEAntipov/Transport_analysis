import numpy as np
import pandas as pd


# Данный код отбирает компании с активами выше 400 млн руб. или выручкой выше 800 млн руб. (это условия для
# обязательного аудита), после чего данные по отобранным компаниям обрабатываются (заполняются расчетные статьи
# при отсутствии данных в них, делается математическая проверка заполненных расчетных статей,
# подсчитывается средневзвешенная для ББ и среднеарифметическая для ОФР)

okveds = ['49.1', '49.2', '49.3', '49.4', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2']

for okved in okveds:
    excel = pd.read_excel('2_BFO_of_List_companies.xlsx', sheet_name=okved)

    # срез Активы > 400 млн руб. или Выручка > 800 млн руб.
    sample = excel.query('(current1600 > 400000) | (current2110 > 800000)').reset_index(drop=True)

    # приведение колонок в единый порядок
    sample = sample[['id', 'inn', 'period', 'knd',
                     'current1110', 'current1120', 'current1130', 'current1140', 'current1150', 'current1160', 'current1170', 'current1180', 'current1190', 'current1100', 'current1210', 'current1220', 'current1230', 'current1240', 'current1250', 'current1260', 'current1200',
                     'current1310', 'current1320', 'current1340', 'current1350', 'current1360', 'current1370', 'current1300', 'current13101', 'current13201', 'current13501', 'current13601', 'current13701', 'current13001', 'current1410', 'current1420', 'current1430', 'current1450', 'current1400', 'current1510', 'current1520', 'current1530', 'current1540', 'current1550', 'current1500',
                     'current1600', 'current1700',
                     'current2110', 'current2120', 'current2100', 'current2210', 'current2220', 'current2200', 'current2310', 'current2320', 'current2330', 'current2340', 'current2350', 'current2300', 'current2410', 'current2411', 'current2412', 'current2460', 'current2400', 'current2510', 'current2520', 'current2530', 'current2500', 'current2910', 'current2900']]

    # заполнение пропущенных расчетных статей (Итоги по разделам, Валовая прибыль и т.п.)
    sample.loc[sample['current1100'].isna(), 'current1100'] = sample[['current1110',
                                                                      'current1120',
                                                                      'current1130',
                                                                      'current1140',
                                                                      'current1150',
                                                                      'current1160',
                                                                      'current1170',
                                                                      'current1180',
                                                                      'current1190']].sum(1)

    sample.loc[sample['current1200'].isna(), 'current1200'] = sample[['current1210',
                                                                      'current1220',
                                                                      'current1230',
                                                                      'current1240',
                                                                      'current1250',
                                                                      'current1260']].sum(1)

    sample.loc[sample['current1300'].isna(), 'current1300'] = sample[['current1310',
                                                                      'current1320',
                                                                      'current1340',
                                                                      'current1350',
                                                                      'current1360',
                                                                      'current1370']].sum(1)

    sample.loc[sample['current1400'].isna(), 'current1400'] = sample[['current1410',
                                                                      'current1420',
                                                                      'current1430',
                                                                      'current1450']].sum(1)

    sample.loc[sample['current1500'].isna(), 'current1500'] = sample[['current1510',
                                                                      'current1520',
                                                                      'current1530',
                                                                      'current1540',
                                                                      'current1550']].sum(1)

    sample.loc[sample['current2100'].isna(), 'current2100'] = sample['current2110'].fillna(0) - sample['current2120'].fillna(0)

    sample.loc[sample['current2200'].isna(), 'current2200'] = sample['current2100'] - (sample['current2210'].fillna(0)
                                                                                       + sample['current2220'].fillna(0))

    sample.loc[sample['current2300'].isna(), 'current2300'] = sample['current2200'] + (sample['current2310'].fillna(0)
                                                                                       + sample['current2320'].fillna(0)
                                                                                       - sample['current2330'].fillna(0)
                                                                                       + sample['current2340'].fillna(0)
                                                                                       - sample['current2350'].fillna(0))

    sample.loc[sample['current2400'].isna(), 'current2400'] = sample['current2300'] + (sample['current2410'].fillna(0)
                                                                                       + sample['current2460'].fillna(0))

    sample.loc[sample['current2500'].isna(), 'current2500'] = sample['current2400'] + (sample['current2510'].fillna(0)
                                                                                       + sample['current2520'].fillna(0)
                                                                                       + sample['current2530'].fillna(0))

    # удаление неиспользуемых пояснительных колонок и заполнение пропусков нулями
    sample = sample.drop(['current13101', 'current13201', 'current13501', 'current13601', 'current13701', 'current13001'], axis=1)
    sample = sample.fillna(0)

    # количество интервалов (bins) для расчета средневзвешенной по частоте
    steps = int(len(sample.index) / 2)

    # транспонирование и удаление ненужных строк
    sample = sample.transpose(copy=True)
    sample.rename(columns=sample.iloc[1], inplace=True)
    sample.drop(['id', 'period', 'inn', 'knd'], inplace=True)

    # конвертация названий колонок (ИНН) в формат строки для дальнейшего обращения к ним
    columns = []
    for column in list(sample.columns):
        columns.append(str(column))
    sample.columns = columns

    # конвертация финансовых значений в числовой формат
    for k in list(sample):
        sample[k] = pd.to_numeric(sample[k], errors='ignore')

    # исправления в выборках
    for column in columns:

        if abs(sample.loc['current2110', column] - sample.loc['current2120', column] - sample.loc['current2100', column]) > 5:
            sample.loc['current2120', column] = -sample.loc['current2120', column]

        if abs(sample.loc['current2100', column] - sample.loc['current2210', column] - sample.loc['current2220', column] - sample.loc['current2200', column]) > 5:
            sample.loc['current2210', column] = -sample.loc['current2210', column]
            sample.loc['current2220', column] = -sample.loc['current2220', column]

        if abs(sample.loc['current2200', column] + sample.loc['current2310', column] + sample.loc['current2320', column] - sample.loc['current2330', column] + sample.loc['current2340', column] - sample.loc['current2350', column] - sample.loc['current2300', column]) > 5:
            sample.loc['current2330', column] = -sample.loc['current2330', column]
            sample.loc['current2350', column] = -sample.loc['current2350', column]

        if abs(sample.loc['current2300', column] + sample.loc['current2410', column] + sample.loc['current2460', column] - sample.loc['current2400', column]) > 5:
            sample.loc['current2410', column] = -sample.loc['current2410', column]

    try:
        sample.loc['current2220', '2309121212'] = -sample.loc['current2220', '2309121212']
    except:
        pass

    try:
        sample = sample.drop('7802003841', axis=1)
    except:
        pass

    # подсчет средневзвешенной по частоту
    weighted_average = []
    for index, row in sample.iterrows():
        count, division = np.histogram(row, bins=steps)
        sum_prod = 0
        for i in range(steps):
            sum_prod += ((division[i] + division[i+1]) / 2) * count[i]
        weighted_average.append(sum_prod / count.sum())

    sample[f'weighted_average_bins_{steps}'] = weighted_average

    # для ОФР считаем среднеарифметическую
    for index, row in sample.loc['current2110':'current2900'].iterrows():
        sample.loc[index, f'weighted_average_bins_{steps}'] = ((sample.loc[index].sum(0)
                                                               - sample.loc[index, f'weighted_average_bins_{steps}'])
                                                               / (sample.loc[index].count() - 1))

    # добавляем строки с проверкой расчетных значений (целевое значение проверки - 0)
    sample.loc['check1100'] = sample.loc[['current1110',
                                          'current1120',
                                          'current1130',
                                          'current1140',
                                          'current1150',
                                          'current1160',
                                          'current1170',
                                          'current1180',
                                          'current1190']].sum(0) - sample.loc['current1100']

    sample.loc['check1200'] = sample.loc[['current1210',
                                          'current1220',
                                          'current1230',
                                          'current1240',
                                          'current1250',
                                          'current1260']].sum(0) - sample.loc['current1200']

    sample.loc['check1300'] = sample.loc[['current1310',
                                          'current1320',
                                          'current1340',
                                          'current1350',
                                          'current1360',
                                          'current1370']].sum(0) - sample.loc['current1300']

    sample.loc['check1400'] = sample.loc[['current1410',
                                          'current1420',
                                          'current1430',
                                          'current1450']].sum(0) - sample.loc['current1400']

    sample.loc['check1500'] = sample.loc[['current1510',
                                          'current1520',
                                          'current1530',
                                          'current1540',
                                          'current1550']].sum(0) - sample.loc['current1500']

    sample.loc['check2100'] = (sample.loc['current2110'].fillna(0)
                               - sample.loc['current2120'].fillna(0)
                               - sample.loc['current2100'])

    sample.loc['check2200'] = (sample.loc['current2100']
                               - sample.loc['current2210'].fillna(0)
                               - sample.loc['current2220'].fillna(0)
                               - sample.loc['current2200'])

    sample.loc['check2300'] = (sample.loc['current2200']
                               + sample.loc['current2310'].fillna(0)
                               + sample.loc['current2320'].fillna(0)
                               - sample.loc['current2330'].fillna(0)
                               + sample.loc['current2340'].fillna(0)
                               - sample.loc['current2350'].fillna(0)
                               - sample.loc['current2300'])

    sample.loc['check2400'] = sample.loc[['current2300',
                                          'current2410',
                                          'current2460']].sum(0) - sample.loc['current2400']

    sample.loc['check2500'] = sample.loc[['current2400',
                                          'current2510',
                                          'current2520',
                                          'current2530']].sum(0) - sample.loc['current2500']

    # ссумируем результаты проверки
    sample.loc['check1100', f'weighted_average_bins_{steps}'] = sample.loc['check1100'].sum(0) - sample.loc['check1100', f'weighted_average_bins_{steps}']
    sample.loc['check1200', f'weighted_average_bins_{steps}'] = sample.loc['check1200'].sum(0) - sample.loc['check1200', f'weighted_average_bins_{steps}']
    sample.loc['check1300', f'weighted_average_bins_{steps}'] = sample.loc['check1300'].sum(0) - sample.loc['check1300', f'weighted_average_bins_{steps}']
    sample.loc['check1400', f'weighted_average_bins_{steps}'] = sample.loc['check1400'].sum(0) - sample.loc['check1400', f'weighted_average_bins_{steps}']
    sample.loc['check1500', f'weighted_average_bins_{steps}'] = sample.loc['check1500'].sum(0) - sample.loc['check1500', f'weighted_average_bins_{steps}']
    sample.loc['check2100', f'weighted_average_bins_{steps}'] = sample.loc['check2100'].sum(0) - sample.loc['check2100', f'weighted_average_bins_{steps}']
    sample.loc['check2200', f'weighted_average_bins_{steps}'] = sample.loc['check2200'].sum(0) - sample.loc['check2200', f'weighted_average_bins_{steps}']
    sample.loc['check2300', f'weighted_average_bins_{steps}'] = sample.loc['check2300'].sum(0) - sample.loc['check2300', f'weighted_average_bins_{steps}']
    sample.loc['check2400', f'weighted_average_bins_{steps}'] = sample.loc['check2400'].sum(0) - sample.loc['check2400', f'weighted_average_bins_{steps}']
    sample.loc['check2500', f'weighted_average_bins_{steps}'] = sample.loc['check2500'].sum(0) - sample.loc['check2500', f'weighted_average_bins_{steps}']

    with pd.ExcelWriter('3.0_BFO_sample.xlsx', mode='a') as writer:
        sample.to_excel(writer, sheet_name=okved)
