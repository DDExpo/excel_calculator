import re
from pathlib import Path

import pandas as pd

from const import COLUMNS_MAP, DATE_NAME
from product_class import Product


def calculate_profit_excel(
        excel: Path, save_to: Path, settings: dict[str: str],
        diff_kaspi_same_product: dict[str: str],
        products: dict[str: Product, str]
        ) -> None:

    start_row: int = int(settings['row_start'])
    col_main_sum: str = settings['column_main_sum']
    col_expenses: list[str] = settings['column_expenses'].split(',')
    col_kaspi_name: str = settings['column_name']
    col_amount: str = settings['column_amount']
    tax: int = int(settings['tax'])

    total_sum: int = 0
    total_sum_expenses: int = 0
    total_sum_clean: int = 0
    refunds: int = 0

    columns_indicies = [
        COLUMNS_MAP[col] for col in [col_main_sum, *col_expenses,
                                     col_kaspi_name]]

    if col_amount:
        columns_indicies.append(COLUMNS_MAP[col_amount])

    xlsx = pd.ExcelFile(excel)
    df = pd.read_excel(xlsx,
                       usecols=columns_indicies,
                       skiprows=start_row, engine='openpyxl')
    xlsx.close()

    lenght = len(df.columns) if not col_amount else len(df.columns) - 1
    tax_col = lenght-1
    pprice_col = lenght
    amount_col = lenght + 1
    res_col = lenght + 2

    df = df.fillna(0)
    df.insert(loc=tax_col, column=f'Налог {tax}%', value=None)
    df.insert(loc=pprice_col, column='Закупочная цена', value=None)
    if not col_amount:
        df.insert(loc=amount_col, column='Количество', value=None)
    df.insert(loc=res_col, column='Итог', value=None)

    for index, row in df.iterrows():
        values = row.values
        operation_price = int(values[0])
        total_sum += operation_price if operation_price >= 0 else 0
        kaspi_name = values[-1]
        quantity = 1
        splitted = (kaspi_name.split(','))

        if len(splitted) > 1:
            raw_quantity = splitted[-1].split()
            if re.search(pattern=r'шт.*', string=raw_quantity[-1]):
                kaspi_name = ''.join(splitted[:-1])
                quantity = int(raw_quantity[-2])
            else:
                kaspi_name = ''.join(splitted)
        else:
            kaspi_name = ''.join(splitted)

        key = kaspi_name.lower().strip()
        amount = ''

        if col_amount:
            quantity = int(re.search(r'(\d+)', values[-3]).group(1))

        if key in diff_kaspi_same_product:
            key, amount, type_amount = diff_kaspi_same_product[key]
            if type_amount == 'ONLY_USER_AMOUNT':
                quantity = int(amount)
            elif type_amount == 'BOTH_AMOUNT':
                quantity = int(amount) * quantity
            elif type_amount == 'ONLY_KASPI_AMOUNT':
                pass

        product_price = 0
        if key in products:
            product_price = int(products[key].price)
        curr_sum = operation_price
        price_for_amount = 0

        for i in range(1, len(df.columns)-1):
            if i <= len(col_expenses):
                expense = int(df.iat[index, i])
                curr_sum += expense
                df.iat[index, i] = curr_sum

                if operation_price < 0:
                    continue

                total_sum_expenses += expense
            elif tax_col == i:
                if operation_price < 0:
                    continue

                curr_tax = int((operation_price / 100) * tax)
                total_sum_expenses += -curr_tax
                curr_sum -= curr_tax
                df.iat[index, i] = curr_sum
            elif pprice_col == i:
                df.iat[index, i] = product_price
            elif amount_col == i:
                price_for_amount = quantity * product_price
                df.iat[index, i] = f'{quantity} шт.'
            elif res_col == i:
                if operation_price < 0:
                    refunds += curr_sum
                    df.iat[index, i] = curr_sum
                else:
                    total = curr_sum - price_for_amount
                    total_sum_clean += total
                    df.iat[index, i] = total

    df = pd.concat([pd.DataFrame([[None] * 5],
                                 columns=range(5), index=['']).T, df])

    totals = [('Общая сумма продаж', total_sum),
              ('Общие затраты', total_sum_expenses),
              ('Чистая выручка', total_sum_clean),
              ('Отмены', refunds),]

    for i, val in enumerate(totals, start=1):
        df.iat[i, 0] = val[0]
        df.iat[i, 1] = val[1]

    copy_excel = 0
    while (save_to / (f'{DATE_NAME}_result_{copy_excel}.xlsx')).exists():
        copy_excel += 1

    df.to_excel(save_to / (f'{DATE_NAME}_result_{copy_excel}.xlsx'),
                index=False)
