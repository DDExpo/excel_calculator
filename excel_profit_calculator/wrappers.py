import traceback
from inspect import signature
from functools import wraps
from typing import Callable

from pandas.errors import ParserError
from PySide6.QtWidgets import QMessageBox

from custom_exceptions import NoneInColumnName
from const import BASE_DIR


def cool_wrapper(func: Callable) -> Callable:
    @wraps(func)
    def wrapper(*args, **kwargs):
        try:
            if len(signature(func).parameters) > 1:
                return func(*args, **kwargs)
            return func(*args)

        except AttributeError:
            QMessageBox.warning(
                None, 'Ошибка',
                'Убедитесь что выбрали правильную колонну для названий!\n')
        except NoneInColumnName:
            QMessageBox.warning(
                None, 'Ошибка',
                'Убедитесь что выбрали правильную колонну для названий!\n')
        except ParserError as e:
            str_e = str(e)
            if 'Defining usecols with out-of-bounds' in str_e:
                QMessageBox.warning(
                    None, 'Ошибка',
                    'Значение "Колонна с наименованиями"'
                    ' находится вне досигаемости!\n'
                    'Понизьте значение колонны!')
        except Exception as e:
            QMessageBox.critical(None, 'Ошибка', 'Что-то пошло не так!')
            with open(BASE_DIR / 'Logs.txt', 'w') as logs:
                logs.write(f'{traceback.format_exc()}\n{e}')
        return wrapper
    return wrapper
