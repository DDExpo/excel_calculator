from pathlib import Path
from datetime import datetime, timedelta
from typing import Union

from utility_functions import get_desktop_path

BASE_DIR: Path = Path(__file__).parent.parent
DESKTOP_PATH: Path = Path(get_desktop_path())
ICONS: Path = BASE_DIR / 'icons/'
DATE_NAME: datetime = str(datetime.now().strftime('%d-%m-%Y'))
DATE_CUR: datetime = datetime.now()
TO_BACKUP_TIME: datetime = timedelta(hours=2)
USER_SETTINGS: Path = BASE_DIR / 'user_settings.txt'
BACKUP:  Path = BASE_DIR / 'backup'
LAST_BACKUP: Path = BASE_DIR / 'backup/last.txt'
COLUMNS: int = 5

FILE_LOCK: Path = BASE_DIR / 'file_lock'

COLUMNS_MAP: dict[str: int] = {
    'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4, 'F': 5, 'G': 6, 'H': 7, 'I': 8,
    'J': 9, 'K': 10, 'L': 11, 'M': 12, 'N': 13, 'O': 14, 'P': 15, 'Q': 16,
    'R': 17, 'S': 18, 'T': 19, 'U': 20, 'V': 21, 'W': 22, 'X': 23, 'Y': 24,
    'Z': 25, 'AA': 26, 'AB': 27, 'AC': 28, 'AD': 29, 'AE': 30, 'AF': 31,
    'AG': 32, 'AH': 33, 'AI': 34, 'AJ': 35, 'AK': 36, 'AL': 37, 'AM': 38,
    'AN': 39, 'AO': 40, 'AP': 41, 'AQ': 42, 'AR': 43, 'AS': 44, 'AT': 45,
    'AU': 46, 'AV': 47, 'AW': 48, 'AX': 49, 'AY': 50, 'AZ': 51, 'BA': 52,
    'BB': 53, 'BC': 54, 'BD': 55, 'BE': 56, 'BF': 57, 'BG': 58, 'BH': 59,
    'BI': 60, 'BJ': 61, 'BK': 62, 'BL': 63, 'BM': 64, 'BN': 65, 'BO': 66,
    'BP': 67, 'BQ': 68, 'BR': 69, 'BS': 70, 'BT': 71, 'BU': 72, 'BV': 73,
    'BW': 74, 'BX': 75, 'BY': 76, 'BZ': 77, 'CA': 78, 'CB': 79, 'CC': 80,
    'CD': 81, 'CE': 82, 'CF': 83, 'CG': 84, 'CH': 85, 'CI': 86, 'CJ': 87,
    'CK': 88, 'CL': 89, 'CM': 90, 'CN': 91, 'CO': 92, 'CP': 93, 'CQ': 94,
    'CR': 95, 'CS': 96, 'CT': 97, 'CU': 98
}

DEFAULT_SETTINGS: dict[str: Union[str, int]] = {
    'column_main_sum': 'S', 'column_expenses': 'U,W,Y,AA,AD',
    'column_name': 'AF', 'column_amount': '', 'row_start': 6, 'tax': 3,
    'last_backup': str(DATE_CUR),
}

EXTRACT_DATA_DIR: Path = BASE_DIR / 'data/'

APP_ICON: str = str(BASE_DIR / 'icons/app_icon.jpg')

AMOUNT_NAME: dict[str: str] = {'ONLY_USER_AMOUNT': 'Пользовательское',
                               'ONLY_KASPI_AMOUNT': 'Каспи',
                               'BOTH_AMOUNT': 'Оба'}

NAME_AMOUNT: dict[str: str] = {'Пользовательское': 'ONLY_USER_AMOUNT',
                               'Каспи': 'ONLY_KASPI_AMOUNT',
                               'Оба': 'BOTH_AMOUNT'}
