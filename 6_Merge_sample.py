import pandas as pd


# Данный код собирает данные по отобранным компаниям за 2023, 2022, 2021 в одну таблицу

okveds = ['49.1', '49.2', '49.3', '49.4', '49.5', '50.10', '50.2', '50.3', '50.4', '51.10', '51.2']

for okved in okveds:
    excel_current = pd.read_excel('3.1_BFO_sample.xlsx', sheet_name=okved)
    excel_current = excel_current[['Наименование статьи', 'Код статьи', 'Итого']]
    excel_current['Код статьи'] = excel_current['Код статьи'].str.replace('current', '')
    excel_current.set_index(['Наименование статьи', 'Код статьи'], inplace=True)
    excel_current.columns = ['2023']

    excel_previous = pd.read_excel('5.1_BFO_sample_previous.xlsx', sheet_name=okved)
    excel_previous = excel_previous[['Наименование статьи', 'Код статьи', 'Итого']]
    excel_previous['Код статьи'] = excel_previous['Код статьи'].str.replace('previous', '')
    excel_previous.set_index(['Наименование статьи', 'Код статьи'], inplace=True)
    excel_previous.columns = ['2022']

    excel_beforePrevious = pd.read_excel('5.1_BFO_sample_beforePrevious.xlsx', sheet_name=okved)
    excel_beforePrevious = excel_beforePrevious[['Наименование статьи', 'Код статьи', 'Итого']]
    excel_beforePrevious['Код статьи'] = excel_beforePrevious['Код статьи'].str.replace('beforePrevious', '')
    excel_beforePrevious.set_index(['Наименование статьи', 'Код статьи'], inplace=True)
    excel_beforePrevious.columns = ['2021']

    merge_sample = excel_current.merge(right=excel_previous, how='left', left_index=True, right_index=True)
    merge_sample = merge_sample.merge(right=excel_beforePrevious, how='left', left_index=True, right_index=True)
    merge_sample.dropna(axis=0, how='all', inplace=True)

    with pd.ExcelWriter('6_Merge_sample.xlsx', mode='a') as writer:
        merge_sample.to_excel(writer, sheet_name=okved)
