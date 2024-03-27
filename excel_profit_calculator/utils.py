import os
import glob
import re
from pathlib import Path
from collections import defaultdict

import pandas as pd

from product_class import Product
from custom_exceptions import NoneInColumnName
from const import (
    BASE_DIR, LAST_BACKUP, USER_SETTINGS, DATE_CUR, DEFAULT_SETTINGS, BACKUP)


def read_settings() -> dict[str: [str, list[str]]]:

    settings: dict[str, str] = defaultdict(str)

    if USER_SETTINGS.exists():
        with open(USER_SETTINGS, 'r', encoding='utf-8') as f:
            for line in f:
                key, val = line.strip('\n').split('|')
                val = int(val) if val.isdecimal() else val
                if key == 'last_backup' and not val:
                    val = str(DATE_CUR)
                settings[key] = val
    else:
        with open(USER_SETTINGS, 'w', encoding='utf-8') as f:
            pass
        settings = DEFAULT_SETTINGS

    return settings


def get_kaspi_names_excel(path_to_excel: Path, row: int,
                          column_name: str) -> set[str]:

    uniq_kaspi_names: set[str] = set()

    xlsx = pd.ExcelFile(path_to_excel)
    df = pd.read_excel(xlsx, usecols=column_name, skiprows=row,
                       engine='openpyxl')
    values = df.values
    xlsx.close()

    for value in values:
        if not value:
            raise NoneInColumnName
        splitted_name = (value[0].split(','))
        if len(splitted_name) > 1:
            raw_quantity = splitted_name[-1].split()
            if re.search(pattern=r'шт.*', string=raw_quantity[-1]):
                kaspi_name = ''.join(splitted_name[:-1])
            else:
                kaspi_name = ''.join(splitted_name)
        else:
            kaspi_name = ''.join(splitted_name)
        uniq_kaspi_names.add(kaspi_name.lower().strip())

    return uniq_kaspi_names


def get_last_backup_data(
        load_notlast: bool = False) -> tuple[dict[str: Product],
                                             set[str], dict[str: str],
                                             dict[str: str]]:

    data: dict[str: Product] = {}
    kaspi_names: set[str] = set()
    product_names: dict[str: str] = {}
    same_product: dict[str: tuple[str, str, str]] = {}
    backup_path = LAST_BACKUP

    if load_notlast:
        files = glob.glob(str(BACKUP/'*'))
        files.sort(key=os.path.getctime)

        if not backup_path.exists() or len(files) < 2:
            return (None, None, None, None)

        backup_path = BACKUP / files[1]

    if backup_path.exists():
        with open(backup_path, 'r', encoding='utf-8') as f:
            for line in f:
                if line.strip():
                    values_dict = {}

                    kaspi, name, price, values = line.split('||')
                    kaspi_name = kaspi.strip().lower()
                    name = name.strip().lower()
                    values = values.rstrip('\n').split('#$#')

                    for val in values:
                        if val.strip():
                            v = val.split('|')
                            key = v[0].lower().strip()
                            quantity = v[1].strip()
                            type_amount = v[2]
                            values_dict[key] = (quantity, type_amount)
                            same_product[key] = (kaspi_name, quantity,
                                                 type_amount)
                            kaspi_names.add(key)

                    kaspi_names.add(kaspi_name)

                    if name:
                        product_names[name] = kaspi_name

                    data[kaspi_name] = Product(
                        name=name, price=price, main_kaspi_name=kaspi_name,
                        connected_kaspi_names=values_dict)

    else:
        with open(backup_path, 'w', encoding='utf-8') as f:
            pass

    return (data, kaspi_names, product_names, same_product)


def check_dirs_ifnot_create(directories: list[str]) -> None:

    for directory in directories:
        path_to_dir = BASE_DIR / directory

        if path_to_dir.exists():
            return

        path_to_dir.mkdir(parents=True, exist_ok=True)
        return


def delete_latest_backup(directory: Path) -> None:

    files = glob.glob(str(directory/'*'))

    if len(files) > 99:
        latest_file = max(files, key=os.path.getctime)
        os.remove(latest_file)
