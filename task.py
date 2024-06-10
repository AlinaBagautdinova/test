import pandas as pd
import matplotlib.pyplot as plt

# Загружаем данные из Excel файла
file = 'data.xlsx'
data = pd.read_excel(
    file, header=0, usecols=[0, 1, 2, 3, 4, 6, 7], skiprows=[1, 130, 259, 370, 485, 595]
)

# Добавляем новый столбец 'period', определяющий период для каждой строки на основе ее индекса
def period(index):
    if 0 <= index <= 127:
        return 'май 2021'
    elif 128 <= index <= 255:
        return 'июнь 2021'
    elif 256 <= index <= 365:
        return 'июль 2021'
    elif 366 <= index <= 479:
        return 'август 2021'
    elif 480 <= index <= 588:
        return 'сентябрь 2021'
    elif 589 <= index:
        return 'октябрь 2021'
    else:
        return None
data['period'] = data.index.map(period)

# Обрабатываем данные
# Преобразуем значения в нижний регистр 
data['status'] = data['status'].str.lower()
data['document'] = data['document'].str.lower()

# Заполняем пропущенные значения в столбце 'document' 
# Если известна дата получения оригинала, то в столбце 'document' ставим значение 'оригинал'
# В остальных случаях заполняем значением 'нет'
data.loc[(data.document.isnull()) & (data.receiving_date.notnull()), 'document'] = 'оригинал'
data['document'] = data['document'].fillna('нет')
data.loc[data['document'] == '-', 'document'] = 'нет'

# Преобразуем значения 'receiving_date' в тип datetime 
data.loc[data['receiving_date'] == '-', 'receiving_date'] = ''
data['receiving_date'] = pd.to_datetime(data['receiving_date'])

# 1. Вычислите общую выручку за июль 2021 по тем сделкам, приход денежных средств которых не просрочен.
print('Общая выручка за июль 2021 по тем сделкам, приход денежных средств которых не просрочен, равна ', end='')
print(data.loc[(data['period'] == 'июль 2021') & (data['status'] != 'просрочено'), 'sum'].sum())

# 2. Как изменялась выручка компании за рассматриваемый период?
# Проиллюстрируйте графиком.

period_order = ['май 2021', 'июнь 2021', 'июль 2021', 'август 2021', 'сентябрь 2021', 'октябрь 2021']
data['period'] = pd.Categorical(data['period'], categories=period_order, ordered=True)

revenue_by_period = data.groupby('period', observed=True)['sum'].sum()
revenue_by_period = revenue_by_period.sort_index()

# Строим график
plt.figure(figsize=(10,6))
plt.bar(revenue_by_period.index, revenue_by_period, color='navy', alpha=0.4)

plt.title("Изменение выручки компании за период май-октябрь 2021 г.")
plt.grid(axis='y',alpha=0.2)
plt.ylabel("Размер выручки")
plt.xlabel("Месяц")
plt.show()

# 3. Кто из менеджеров привлек для компании больше всего денежных средств в сентябре 2021?
sale_september = data.loc[data['period']=='сентябрь 2021']
sale_september = sale_september.groupby('sale')['sum'].sum().sort_values(ascending=False)
sale_september.index[0]
print('В сентябре 2021 г. больше всего денежных средств в размере {} привлек {}.'.format(
    sale_september.iloc[0], sale_september.index[0])
)

# 4. Какой тип сделок (новая/текущая) был преобладающим в октябре 2021?
type_october = data.loc[data['period']=='октябрь 2021']
type_october = type_october.groupby('new/current')['client_id'].count().sort_values(ascending=False)
print('В октябре 2021 г. преобладающим типом была {} сделка ({} ед. из {}).'.format(
    type_october.index[0], type_october.iloc[0], type_october.iloc[0] + type_october.iloc[1])
)

# 5. Сколько оригиналов договора по майским сделкам было получено в июне 2021?
document_june = data.loc[
    (data['period']=='май 2021') & 
    (data['receiving_date'] >= '2021-06-01') & 
    (data['receiving_date'] <= '2021-06-30')
]
document_june = document_june.groupby('document')['client_id'].count().sort_values(ascending=False)
print('{} оригиналов договора по майским сделкам было получено в июне 2021 г.'.format(
    document_june.iloc[0])
)

# Задание

# Рассматриваемый месяц - июнь
# Рассчитываем бонусы для каждой сделки
def bonus(row):
    if row['new/current'] == 'новая' and row['status'] == 'оплачено' and row['document'] == 'оригинал':
        return row['sum'] * 0.07
    elif row['new/current'] == 'текущая' and row['status'] != 'просрочено' and row['document'] == 'оригинал':
        if row['sum'] > 10000:
            return row['sum'] * 0.05
        else:
            return row['sum'] * 0.03
    return 0

# Фильтрация данных для сделок, оригиналы для которых пришли позже 01.07.2021
data_bonus = data.loc[(data['period'] == 'июнь 2021') & (data['receiving_date'] >= '2021-07-01')].copy()

# Остаток по сделкам, оригиналы для которых пришли позже 
data_bonus['bonus'] = data_bonus.apply(bonus, axis=1)

remaining_bonus = data_bonus.groupby('sale')['bonus'].sum()

# Вывод остатка каждого менеджера на 01.07.2021
print('Остаток менеджеров на 01.07.2021:')
print(remaining_bonus)
