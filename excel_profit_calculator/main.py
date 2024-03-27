import sys
import os
import ctypes
import traceback

from PySide6.QtWidgets import QApplication
from PySide6.QtGui import QIcon

from const import APP_ICON, BACKUP, FILE_LOCK, BASE_DIR
from utils import (
    check_dirs_ifnot_create, read_settings, delete_latest_backup)
from user_interface import ExcelCalculator


def main(window=None) -> None:

    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon(APP_ICON))

    check_dirs_ifnot_create(['backup'])
    delete_latest_backup(BACKUP)
    user_settings = read_settings()

    window = ExcelCalculator(user_settings)
    window.show()
    app.exec()


if __name__ == '__main__':

    try:
        if not FILE_LOCK.exists():
            with open(FILE_LOCK, 'w') as f:
                main()
        else:
            ctypes.windll.user32.MessageBoxW(
                0, 'Программа уже запущена!', 'Предупреждение', 0x0 | 0x40)

    except BaseException as e:
        with open(BASE_DIR / 'Logs.txt', 'w', encoding='utf-8') as logs:
            logs.write(f'{traceback.format_exc()}\n{e}')
        ctypes.windll.user32.MessageBoxW(0, 'Что то пошло не так',
                                         'Ошибка', 0x0 | 0x40)
    finally:
        try:
            os.remove(FILE_LOCK)
        except IOError:
            pass
