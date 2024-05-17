import numpy as np
import pandas as pd


pd.set_option('display.max_columns', None)
pd.set_option('display.max_rows', None)
pd.set_option('display.float_format', lambda x: '%.3f' % x)
okveds = ['49.1', '49.2', '49.3', '49.4', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2']
#

for okved in okveds:
    for year in ['previous', 'beforePrevious']:
        excel = pd.read_excel(f'4_BFO_List_{year}.xlsx', sheet_name=okved)
        sample = excel.reset_index(drop=True)

        sample = sample[['id', 'inn', 'period', 'knd',
                         f'{year}1110', f'{year}1120', f'{year}1130', f'{year}1140', f'{year}1150', f'{year}1160', f'{year}1170', f'{year}1180', f'{year}1190', f'{year}1100', f'{year}1210', f'{year}1220', f'{year}1230', f'{year}1240', f'{year}1250', f'{year}1260', f'{year}1200',
                         f'{year}1310', f'{year}1320', f'{year}1340', f'{year}1350', f'{year}1360', f'{year}1370', f'{year}1300', f'{year}13101', f'{year}13201', f'{year}13501', f'{year}13601', f'{year}13701', f'{year}13001', f'{year}1410', f'{year}1420', f'{year}1430', f'{year}1450', f'{year}1400', f'{year}1510', f'{year}1520', f'{year}1530', f'{year}1540', f'{year}1550', f'{year}1500',
                         f'{year}1600', f'{year}1700',
                         f'{year}2110', f'{year}2120', f'{year}2100', f'{year}2210', f'{year}2220', f'{year}2200', f'{year}2310', f'{year}2320', f'{year}2330', f'{year}2340', f'{year}2350', f'{year}2300', f'{year}2410', f'{year}2411', f'{year}2412', f'{year}2460', f'{year}2400', f'{year}2510', f'{year}2520', f'{year}2530', f'{year}2500', f'{year}2910', f'{year}2900']]

        sample.loc[sample[f'{year}1100'].isna(), f'{year}1100'] = sample[[f'{year}1110',
                                                                          f'{year}1120',
                                                                          f'{year}1130',
                                                                          f'{year}1140',
                                                                          f'{year}1150',
                                                                          f'{year}1160',
                                                                          f'{year}1170',
                                                                          f'{year}1180',
                                                                          f'{year}1190']].sum(1)

        sample.loc[sample[f'{year}1200'].isna(), f'{year}1200'] = sample[[f'{year}1210',
                                                                          f'{year}1220',
                                                                          f'{year}1230',
                                                                          f'{year}1240',
                                                                          f'{year}1250',
                                                                          f'{year}1260']].sum(1)

        sample.loc[sample[f'{year}1300'].isna(), f'{year}1300'] = sample[[f'{year}1310',
                                                                          f'{year}1320',
                                                                          f'{year}1340',
                                                                          f'{year}1350',
                                                                          f'{year}1360',
                                                                          f'{year}1370']].sum(1)

        sample.loc[sample[f'{year}1400'].isna(), f'{year}1400'] = sample[[f'{year}1410',
                                                                          f'{year}1420',
                                                                          f'{year}1430',
                                                                          f'{year}1450']].sum(1)

        sample.loc[sample[f'{year}1500'].isna(), f'{year}1500'] = sample[[f'{year}1510',
                                                                          f'{year}1520',
                                                                          f'{year}1530',
                                                                          f'{year}1540',
                                                                          f'{year}1550']].sum(1)

        sample.loc[sample[f'{year}2100'].isna(), f'{year}2100'] = sample[f'{year}2110'].fillna(0) - sample[f'{year}2120'].fillna(0)

        sample.loc[sample[f'{year}2200'].isna(), f'{year}2200'] = sample[f'{year}2100'] - (sample[f'{year}2210'].fillna(0)
                                                                                           + sample[f'{year}2220'].fillna(0))

        sample.loc[sample[f'{year}2300'].isna(), f'{year}2300'] = sample[f'{year}2200'] + (sample[f'{year}2310'].fillna(0)
                                                                                           + sample[f'{year}2320'].fillna(0)
                                                                                           - sample[f'{year}2330'].fillna(0)
                                                                                           + sample[f'{year}2340'].fillna(0)
                                                                                           - sample[f'{year}2350'].fillna(0))

        sample.loc[sample[f'{year}2400'].isna(), f'{year}2400'] = sample[f'{year}2300'] + (sample[f'{year}2410'].fillna(0)
                                                                                           + sample[f'{year}2460'].fillna(0))

        sample.loc[sample[f'{year}2500'].isna(), f'{year}2500'] = sample[f'{year}2400'] + (sample[f'{year}2510'].fillna(0)
                                                                                           + sample[f'{year}2520'].fillna(0)
                                                                                           + sample[f'{year}2530'].fillna(0))

        sample = sample.drop([f'{year}13101', f'{year}13201', f'{year}13501', f'{year}13601', f'{year}13701', f'{year}13001'], axis=1)
        sample = sample.fillna(0)

        steps = int(len(sample.index) / 2)

        sample = sample.transpose(copy=True)
        sample.rename(columns=sample.iloc[1], inplace=True)
        sample.drop(['id', 'period', 'inn', 'knd'], inplace=True)

        columns = []
        for column in list(sample.columns):
            columns.append(str(column))
        sample.columns = columns

        for k in list(sample):
            sample[k] = pd.to_numeric(sample[k], errors='ignore')

        # исправления в выборках
        for column in columns:

            if abs(sample.loc[f'{year}2110', column] - sample.loc[f'{year}2120', column] - sample.loc[f'{year}2100', column]) > 5:
                sample.loc[f'{year}2120', column] = -sample.loc[f'{year}2120', column]

            if abs(sample.loc[f'{year}2100', column] - sample.loc[f'{year}2210', column] - sample.loc[f'{year}2220', column] - sample.loc[f'{year}2200', column]) > 5:
                sample.loc[f'{year}2210', column] = -sample.loc[f'{year}2210', column]
                sample.loc[f'{year}2220', column] = -sample.loc[f'{year}2220', column]

            if abs(sample.loc[f'{year}2200', column] + sample.loc[f'{year}2310', column] + sample.loc[f'{year}2320', column] - sample.loc[f'{year}2330', column] + sample.loc[f'{year}2340', column] - sample.loc[f'{year}2350', column] - sample.loc[f'{year}2300', column]) > 5:
                sample.loc[f'{year}2330', column] = -sample.loc[f'{year}2330', column]
                sample.loc[f'{year}2350', column] = -sample.loc[f'{year}2350', column]

            if abs(sample.loc[f'{year}2300', column] + sample.loc[f'{year}2410', column] + sample.loc[f'{year}2460', column] - sample.loc[f'{year}2400', column]) > 5:
                sample.loc[f'{year}2410', column] = -sample.loc[f'{year}2410', column]

        try:
            sample.loc[f'{year}2220', '2309121212'] = -sample.loc[f'{year}2220', '2309121212']
        except:
            pass

        try:
            sample = sample.drop('7802003841', axis=1)
        except:
            pass

        weighted_average = []
        for index, row in sample.iterrows():
            count, division = np.histogram(row, bins=steps)
            sum_prod = 0
            for i in range(steps):
                sum_prod += ((division[i] + division[i+1]) / 2) * count[i]
            weighted_average.append(sum_prod / count.sum())

        sample[f'weighted_average_bins_{steps}'] = weighted_average

        # для ОФР считаем среднеарифметическую
        for index, row in sample.loc[f'{year}2110':f'{year}2900'].iterrows():
            sample.loc[index, f'weighted_average_bins_{steps}'] = ((sample.loc[index].sum(0)
                                                                   - sample.loc[index, f'weighted_average_bins_{steps}'])
                                                                   / (sample.loc[index].count() - 1))

        sample.loc['check1100'] = sample.loc[[f'{year}1110',
                                              f'{year}1120',
                                              f'{year}1130',
                                              f'{year}1140',
                                              f'{year}1150',
                                              f'{year}1160',
                                              f'{year}1170',
                                              f'{year}1180',
                                              f'{year}1190']].sum(0) - sample.loc[f'{year}1100']

        sample.loc['check1200'] = sample.loc[[f'{year}1210',
                                              f'{year}1220',
                                              f'{year}1230',
                                              f'{year}1240',
                                              f'{year}1250',
                                              f'{year}1260']].sum(0) - sample.loc[f'{year}1200']

        sample.loc['check1300'] = sample.loc[[f'{year}1310',
                                              f'{year}1320',
                                              f'{year}1340',
                                              f'{year}1350',
                                              f'{year}1360',
                                              f'{year}1370']].sum(0) - sample.loc[f'{year}1300']

        sample.loc['check1400'] = sample.loc[[f'{year}1410',
                                              f'{year}1420',
                                              f'{year}1430',
                                              f'{year}1450']].sum(0) - sample.loc[f'{year}1400']

        sample.loc['check1500'] = sample.loc[[f'{year}1510',
                                              f'{year}1520',
                                              f'{year}1530',
                                              f'{year}1540',
                                              f'{year}1550']].sum(0) - sample.loc[f'{year}1500']

        sample.loc['check2100'] = (sample.loc[f'{year}2110'].fillna(0)
                                   - sample.loc[f'{year}2120'].fillna(0)
                                   - sample.loc[f'{year}2100'])

        sample.loc['check2200'] = (sample.loc[f'{year}2100']
                                   - sample.loc[f'{year}2210'].fillna(0)
                                   - sample.loc[f'{year}2220'].fillna(0)
                                   - sample.loc[f'{year}2200'])

        sample.loc['check2300'] = (sample.loc[f'{year}2200']
                                   + sample.loc[f'{year}2310'].fillna(0)
                                   + sample.loc[f'{year}2320'].fillna(0)
                                   - sample.loc[f'{year}2330'].fillna(0)
                                   + sample.loc[f'{year}2340'].fillna(0)
                                   - sample.loc[f'{year}2350'].fillna(0)
                                   - sample.loc[f'{year}2300'])

        sample.loc['check2400'] = sample.loc[[f'{year}2300',
                                              f'{year}2410',
                                              f'{year}2460']].sum(0) - sample.loc[f'{year}2400']

        sample.loc['check2500'] = sample.loc[[f'{year}2400',
                                              f'{year}2510',
                                              f'{year}2520',
                                              f'{year}2530']].sum(0) - sample.loc[f'{year}2500']

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

        with pd.ExcelWriter(f'5.0_BFO_sample_{year}.xlsx', mode='a') as writer:
            sample.to_excel(writer, sheet_name=okved)