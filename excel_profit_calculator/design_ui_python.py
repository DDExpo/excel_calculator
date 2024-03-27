from PySide6.QtCore import QRect, QSize, Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import (
    QFrame, QPushButton, QGridLayout, QGroupBox, QWidget, QVBoxLayout, QLabel,
    QProgressBar, QTableWidget, QLineEdit, QHeaderView, QListWidget,
    QHBoxLayout, QSplitter, QSpinBox, QComboBox, QDialog, QRadioButton
)

from const import ICONS, COLUMNS


class Ui_MainWindow(object):

    def setupUi(self, MainWindow):
        if not MainWindow.objectName():
            MainWindow.setObjectName(u'MainWindow')
        MainWindow.resize(433, 390)
        MainWindow.setWindowTitle('ExcelProfitCalculator')

        self.centralwidget = QWidget(MainWindow)
        self.centralwidget.setObjectName(u'centralwidget')

        self._add_item_window()
        self._setting_window(MainWindow)
        self.create_grid_group_box()
        self.create_filter_table()
        self._combine_item_window()
        self._connected_kaspi_to_products_window()
        self._change_amount_type_window()

        self.filter_table.setColumnHidden(4, True)

        icon_backup = QIcon()
        icon_backup.addFile(str(ICONS / 'backup.svg'))
        self.backup_butt = QPushButton()
        self.backup_butt.setEnabled(True)
        self.backup_butt.setMaximumSize(QSize(29, 29))
        self.backup_butt.setIcon(icon_backup)

        icon_menu = QIcon()
        icon_menu.addFile(str(ICONS / 'menu.svg'))
        self.button_menu = QPushButton()
        self.button_menu.setEnabled(True)
        self.button_menu.setMaximumSize(QSize(29, 29))
        self.button_menu.setIcon(icon_menu)

        icon_box = QIcon()
        icon_box.addFile(str(ICONS / 'box.svg'))
        self.combine_butt = QPushButton()
        self.combine_butt.setEnabled(True)
        self.combine_butt.setMaximumSize(QSize(29, 29))
        self.combine_butt.setIcon(icon_box)

        icon_sett = QIcon()
        icon_sett.addFile(str(ICONS / 'settings.svg'))
        self.button_settings = QPushButton()
        self.button_settings.setEnabled(True)
        self.button_settings.setMaximumSize(QSize(29, 29))
        self.button_settings.setIcon(icon_sett)

        self.button_layout = QHBoxLayout()
        self.button_layout.addWidget(self.backup_butt)
        self.button_layout.addWidget(self.combine_butt)
        self.button_layout.addWidget(self.button_menu)
        self.button_layout.addWidget(self.button_settings)

        self.progressbar = QProgressBar()
        self.progressbar.setGeometry(QRect(100, 155, 226, 45))
        self.progressbar.setMinimum(0)
        self.progressbar.setMaximum(100)
        self.progressbar.hide()

        self.splitter = QSplitter()
        self.splitter.addWidget(self._grid_group_box)
        self.splitter.addWidget(self._grid_filter_table)

        self.main_layout = QGridLayout(self.centralwidget)
        self.main_layout.setContentsMargins(5, 5, 5, 0)
        self.main_layout.addWidget(self.splitter, 1, 0)
        self.main_layout.addLayout(
            self.button_layout, 0, 0, alignment=Qt.AlignmentFlag.AlignRight)

        MainWindow.setLayout(self.main_layout)
        MainWindow.setCentralWidget(self.centralwidget)

    def create_grid_group_box(self):
        self._grid_group_box = QGroupBox()
        self._grid_group_box.setStyleSheet(
            '''QGroupBox { padding: 0; padding-bottom: 10px;
               border: none; }''')

        layout = QVBoxLayout()

        frame = QFrame()
        frame.setEnabled(True)
        frame.setFrameShape(QFrame.Box)
        frame.setMouseTracking(True)
        frame.setAcceptDrops(True)

        self.file_label = QLabel('<center>Перетащите сюда Excel</center>',
                                 frame)
        self.file_label.setWordWrap(True)

        frame_layout = QVBoxLayout(frame)
        frame_layout.addWidget(
            self.file_label, alignment=Qt.AlignmentFlag.AlignCenter |
            Qt.AlignmentFlag.AlignVCenter)

        self.button_choose = QPushButton('Выбрать excel')
        self.button_start = QPushButton('Старт')
        self.button_choose_res = QPushButton(
            'Выбрать куда сохранять результат')

        layout.addWidget(frame)
        layout.setContentsMargins(5, 0, 1, 1)
        layout.addWidget(self.button_start)
        layout.addWidget(self.button_choose)
        layout.addWidget(self.button_choose_res)

        self._grid_group_box.setLayout(layout)

    def create_filter_table(self):
        self._grid_filter_table = QGroupBox()
        self._grid_filter_table.setStyleSheet(
            '''QGroupBox { padding: 0px;}''')

        self.filter_table = QTableWidget()
        self.filter_table.setColumnCount(COLUMNS)
        self.filter_table.setHorizontalHeaderLabels(
            ['Каспи Название', 'Название', '', 'Цена', '',])
        self.filter_table.setEditTriggers(QTableWidget.DoubleClicked)
        self.filter_table.setSortingEnabled(True)

        horizontal_header = self.filter_table.horizontalHeader()
        horizontal_header.setSectionResizeMode(
            0, QHeaderView.ResizeMode.Interactive)
        horizontal_header.setSectionResizeMode(
            1, QHeaderView.ResizeMode.Interactive)
        horizontal_header.setSectionResizeMode(
            2, QHeaderView.ResizeMode.ResizeToContents)
        horizontal_header.setSectionResizeMode(
            3, QHeaderView.ResizeMode.Stretch)
        horizontal_header.setSectionResizeMode(
            4, QHeaderView.ResizeMode.Stretch)

        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText('Поиск...')
        self.search_input.setStyleSheet('QLineEdit{ color: black; }')

        self.button_additem = QPushButton('Добавить товар')
        self.button_deleteitem = QPushButton('Удалить товар')
        self.button_deleteitem.setStyleSheet(
            'QPushButton {border: 2px solid red; color: red;}'
            'QPushButton:pressed {background-color: #FF3333;}')
        self.button_save_data = QPushButton('Сохранить')

        layout = QVBoxLayout()
        layout.addWidget(self.search_input)
        layout.addWidget(self.filter_table)
        layout.addWidget(self.button_additem)
        layout.addWidget(self.button_save_data)
        layout.addWidget(self.button_deleteitem)

        self._grid_filter_table.setLayout(layout)
        self._grid_filter_table.hide()

    def _setting_window(self, MainWindow):

        self.setting_window = QWidget()
        self.setting_window.setObjectName(u'settings_window')
        self.setting_window.setWindowTitle('Настройки')

        layout = QVBoxLayout(self.setting_window)

        self.col_main_sum = QLineEdit()
        self.col_main_sum.setPlaceholderText('Например: A, AA, AB')
        self.col_main_sum.setStyleSheet('QLineEdit{ color: black; }')
        self.col_main_sum.setText(MainWindow.user_settings['column_main_sum'])
        self.col_expensses = QLineEdit()
        self.col_expensses.setPlaceholderText(
            'Введите колонны через запятую: A,B,C')
        self.col_expensses.setStyleSheet('QLineEdit{ color: black; }')
        self.col_expensses.setText(MainWindow.user_settings['column_expenses'])
        self.col_name = QLineEdit()
        self.col_name.setText(MainWindow.user_settings['column_name'])
        self.col_amount = QLineEdit()
        self.col_amount.setText(MainWindow.user_settings['column_amount'])
        self.start_rows = QSpinBox()
        self.start_rows.setValue(MainWindow.user_settings['row_start'])
        self.start_rows.setRange(0, 100000000)
        self.tax = QSpinBox()
        self.tax.setValue(MainWindow.user_settings['tax'])
        self.tax.setRange(0, 100)
        self.submit_settings_button = QPushButton('Сохранить')

        lab1 = QLabel('Коллонна с суммой')
        lab2 = QLabel('Коллонны расходов')
        lab3 = QLabel('Колонна с наименованиями')
        lab4 = QLabel('Начальный ряд (Не включительно!)')
        lab5 = QLabel('Налог % (От первоначальной суммы)')
        lab6 = QLabel('Колонна с количеством')
        lab1.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab2.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab3.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab4.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab5.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab6.setStyleSheet("font-weight: bold; font-size: 10pt;")

        layout.addWidget(lab1)
        layout.addWidget(self.col_main_sum)
        layout.addWidget(lab2)
        layout.addWidget(self.col_expensses)
        layout.addWidget(lab3)
        layout.addWidget(self.col_name)
        layout.addWidget(lab6)
        layout.addWidget(self.col_amount)
        layout.addWidget(lab4)
        layout.addWidget(self.start_rows)
        layout.addWidget(lab5)
        layout.addWidget(self.tax)
        layout.addWidget(self.submit_settings_button)

        self.setting_window.hide()

    def _add_item_window(self):

        self.add_item_window = QWidget()
        self.add_item_window.setObjectName(u'add_item_window')
        self.add_item_window.setWindowTitle('Добавления товара')

        layout = QVBoxLayout(self.add_item_window)

        self.kaspi_field = QLineEdit()
        self.name_field = QLineEdit()
        self.int_field = QSpinBox()
        self.int_field.setRange(0, 100000000)
        self.submit_button = QPushButton('Сохранить')

        lab1 = QLabel('Каспи название')
        lab2 = QLabel('Название')
        lab3 = QLabel('Цена')
        lab1.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab2.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab3.setStyleSheet("font-weight: bold; font-size: 10pt;")

        layout.addWidget(lab1)
        layout.addWidget(self.kaspi_field)
        layout.addWidget(lab2)
        layout.addWidget(self.name_field)
        layout.addWidget(lab3)
        layout.addWidget(self.int_field)
        layout.addWidget(self.submit_button)

        self.add_item_window.hide()

    def _combine_item_window(self):

        self.combine_item_window = QWidget()
        self.combine_item_window.setObjectName(u'combine_window')
        self.combine_item_window.setWindowTitle('Соединение товара')
        self.combine_item_window.setMinimumSize(QSize(250, 200))

        layout = QVBoxLayout(self.combine_item_window)

        self.main_kaspi_name = QLabel('')
        self.product_main = QComboBox()
        self.product_main.setEditable(True)
        self.product_main.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.kaspi_field_to = QComboBox()
        self.kaspi_field_to.setEditable(True)
        self.kaspi_field_to.setInsertPolicy(QComboBox.InsertPolicy.NoInsert)
        self.amount_field = QSpinBox()
        self.amount_field.setRange(1, 100000000)
        self.submit_button_combine = QPushButton('Соединить')

        lab1 = QLabel('Товары')
        lab2 = QLabel('Основное каспи имя')
        lab3 = QLabel('Каспи имена')
        lab4 = QLabel('Количество')
        lab5 = QLabel('Считать количество:')
        lab1.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab2.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab3.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab4.setStyleSheet("font-weight: bold; font-size: 10pt;")
        lab5.setStyleSheet("font-weight: bold; font-size: 10pt;")

        radioLayout = QHBoxLayout()

        self.radioButton1 = QRadioButton('Только каспи')
        self.radioButton1.setProperty('custom_data', 'ONLY_KASPI_AMOUNT')
        self.radioButton2 = QRadioButton('Установленное пользователем')
        self.radioButton2.setProperty('custom_data', 'ONLY_USER_AMOUNT')
        self.radioButton2.setChecked(True)
        self.radioButton3 = QRadioButton('Оба')
        self.radioButton3.setProperty('custom_data', 'BOTH_AMOUNT')

        radioLayout.addWidget(self.radioButton1)
        radioLayout.addWidget(self.radioButton2)
        radioLayout.addWidget(self.radioButton3)

        layout.addWidget(lab1)
        layout.addWidget(self.product_main)
        layout.addWidget(lab2)
        layout.addWidget(self.main_kaspi_name)
        layout.addWidget(lab3)
        layout.addWidget(self.kaspi_field_to)
        layout.addWidget(lab4)
        layout.addWidget(self.amount_field)
        layout.addWidget(lab5)
        layout.addLayout(radioLayout)
        layout.addWidget(self.submit_button_combine)

        self.combine_item_window.hide()

    def _connected_kaspi_to_products_window(self):

        self.connected_kaspi_products = QDialog()
        self.connected_kaspi_products.setObjectName(
            u'connected_kaspi_products')
        self.connected_kaspi_products.setWindowTitle(
            'Объединенные каспи имена')
        self.connected_kaspi_products.setModal(True)

        layout = QVBoxLayout(self.connected_kaspi_products)
        button_layout = QHBoxLayout()

        self.kaspi_names_connected_list = QListWidget()
        self.save_list_connected_names = QPushButton('Сохранить')
        self.label_connected = QLabel('')
        self.label_connected.setStyleSheet(
            "font-weight: bold; font-size: 11pt;")
        label_product = QLabel('Товар:')
        label_product.setStyleSheet("font-weight: bold; font-size: 12pt;")

        label_layout = QHBoxLayout()
        label_layout.addWidget(label_product)
        label_layout.addWidget(self.label_connected)
        label_layout.setAlignment(Qt.AlignmentFlag.AlignLeft)

        self.edit_connected_name = QPushButton('Изменить')
        self.delete_button = QPushButton('Отсоединить')
        self.delete_button.setStyleSheet(
            'QPushButton {border: 2px solid red; color: red;}'
            'QPushButton:pressed {background-color: #FF3333;}')

        button_layout.addWidget(self.delete_button)
        button_layout.addWidget(self.edit_connected_name)
        button_layout.addWidget(self.save_list_connected_names)

        layout.addLayout(label_layout)
        layout.addWidget(QLabel('Привязанные каспи имена | Количество | Тип'))
        layout.addWidget(self.kaspi_names_connected_list)
        layout.addLayout(button_layout)

        self.connected_kaspi_products.hide()

    def _change_amount_type_window(self):

        self.change_amount_window = QDialog()
        self.change_amount_window.setObjectName(
            u'change_amount_window')
        self.change_amount_window.setWindowTitle('Изменить количество')
        self.change_amount_window.setModal(True)

        layout = QVBoxLayout(self.change_amount_window)

        self.label_change = QLabel('')
        self.label_change.setStyleSheet("font-weight: bold; font-size: 12pt;")
        self.amount_change_field = QSpinBox()
        self.amount_change_field.setRange(1, 100000000)
        layout_amount = QLabel('Считать количество:')
        radioLayout = QHBoxLayout()

        self.radioButton1 = QRadioButton('Только каспи')
        self.radioButton1.setProperty('custom_data', 'ONLY_KASPI_AMOUNT')
        self.radioButton2 = QRadioButton('Установленное пользователем')
        self.radioButton2.setProperty('custom_data', 'ONLY_USER_AMOUNT')
        self.radioButton2.setChecked(True)
        self.radioButton3 = QRadioButton('Оба')
        self.radioButton3.setProperty('custom_data', 'BOTH_AMOUNT')

        radioLayout.addWidget(self.radioButton1)
        radioLayout.addWidget(self.radioButton2)
        radioLayout.addWidget(self.radioButton3)

        self.save_butt_change_amount = QPushButton('Сохранить')

        layout.addWidget(self.label_change)
        layout.addWidget(QLabel('Количество'))
        layout.addWidget(self.amount_change_field)
        layout.addWidget(layout_amount)
        layout.addLayout(radioLayout)
        layout.addWidget(self.save_butt_change_amount)

        self.change_amount_window.hide()

    def _get_amount_type(self) -> tuple[str, str]:
        dialog = QDialog()
        dialog.setMinimumSize(QSize(159, 100))
        dialog.setWindowTitle('Количество')
        layout = QVBoxLayout(dialog)
        button_ok = QPushButton('OK')
        spin_box = QSpinBox()
        spin_box.setRange(1, 100000000)

        layout_amount = QLabel('Считать количество:')
        radioLayout = QHBoxLayout()

        radioButton1 = QRadioButton('Только каспи')
        radioButton1.setProperty('custom_data', 'ONLY_KASPI_AMOUNT')
        radioButton2 = QRadioButton('Установленное пользователем')
        radioButton2.setProperty('custom_data', 'ONLY_USER_AMOUNT')
        radioButton2.setChecked(True)
        radioButton3 = QRadioButton('Оба')
        radioButton3.setProperty('custom_data', 'BOTH_AMOUNT')

        radioLayout.addWidget(radioButton1)
        radioLayout.addWidget(radioButton2)
        radioLayout.addWidget(radioButton3)

        layout.addWidget(QLabel('Количество'))
        layout.addWidget(spin_box)
        layout.addWidget(layout_amount)
        layout.addLayout(radioLayout)
        layout.addWidget(button_ok)

        button_ok.clicked.connect(dialog.accept)

        if dialog.exec() == QDialog.Accepted:
            if radioButton1.isChecked():
                data = radioButton1.property('custom_data')
            elif radioButton2.isChecked():
                data = radioButton2.property('custom_data')
            else:
                data = radioButton3.property('custom_data')

            return (str(spin_box.value()), data)
        else:
            return None
