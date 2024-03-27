import itertools
import traceback
import shutil
import re
from pathlib import Path
from datetime import datetime

from PySide6.QtWidgets import (
    QMainWindow, QFileDialog, QMessageBox, QTableWidgetItem, QPushButton)
from PySide6.QtCore import Slot, QSize, Qt
from qt_material import apply_stylesheet, QtStyleTools

from product_class import Product
from custom_pyside import IntTableWidgetItem
from const import (
    BASE_DIR, DESKTOP_PATH, DATE_NAME, LAST_BACKUP, TO_BACKUP_TIME,
    USER_SETTINGS, BACKUP, DATE_CUR, AMOUNT_NAME)
from utils import (
    get_kaspi_names_excel, get_last_backup_data)
from calculate_profit import calculate_profit_excel
from design_ui_python import Ui_MainWindow
from wrappers import cool_wrapper


class ExcelCalculator(QMainWindow, QtStyleTools):

    def __init__(self, user_settings: dict[str: str]):
        super().__init__()

        self.excel_path: str | None = None
        self.path_to_save_folder: Path = DESKTOP_PATH
        self.user_settings: dict[str: str] = user_settings
        self.row_settings_before: int = user_settings['row_start']

        (self.back_up_data, self.kaspi_names_combined, self.product_names,
         self.diff_names_same_product) = get_last_backup_data()

        self.kaspi_names: set[str] = set()
        self.incorrect_products: set[str] = set()
        self.kaspi_names_to_unmerge: list[str] = []

        self.is_changed: bool = False
        self.populated: bool = False

        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)
        apply_stylesheet(self, 'light_darkgreen.xml')

        self.message_box = QMessageBox(self)

        self.ui.button_additem.clicked.connect(self._open_add_item)
        self.ui.button_settings.clicked.connect(self._open_settings)
        self.ui.combine_butt.clicked.connect(self._open_combine_window)
        self.ui.button_menu.clicked.connect(self._show_filter_table)

        self.ui.button_save_data.clicked.connect(self._save_data)
        self.ui.submit_settings_button.clicked.connect(self._save_settings)
        self.ui.submit_button_combine.clicked.connect(self._combine_products)
        self.ui.save_butt_change_amount.clicked.connect(
            self._change_item_connected_list)
        self.ui.save_list_connected_names.clicked.connect(
            self._unmerge_products)

        self.ui.search_input.textChanged.connect(self._filtering_table)
        self.ui.filter_table.itemChanged.connect(self._on_item_changed)
        self.ui.submit_button.clicked.connect(self._confirm_add_item)
        self.ui.button_deleteitem.clicked.connect(self._delete_item)
        self.ui.backup_butt.clicked.connect(self._load_backup)

        self.ui.button_choose.clicked.connect(self._get_new_path)
        self.ui.button_choose_res.clicked.connect(self._get_new_path_for_res)
        self.ui.button_start.clicked.connect(self._excel_calculate)

        self.ui.delete_button.clicked.connect(self._delete_item_connected_list)
        self.ui.edit_connected_name.clicked.connect(
            self._open_change_amount_window)
        self.ui.product_main.currentIndexChanged.connect(
            self._change_text_label)
        self.ui.product_main.currentIndexChanged.connect(
            self._change_kaspi_names_to_connect)

        self._populate_table()

    @cool_wrapper
    @Slot()
    def _change_kaspi_names_to_connect(self, _: str = '') -> None:

        name = self.ui.main_kaspi_name.text()
        if name.strip():
            to_include = set(self.back_up_data.keys())
            if name in to_include:
                to_include.remove(name)

            self.ui.kaspi_field_to.clear()
            self.ui.kaspi_field_to.addItems(to_include)

    @cool_wrapper
    @Slot()
    def _change_text_label(self, _) -> None:

        name = self.ui.product_main.currentText()
        if name.strip():
            self.ui.main_kaspi_name.setText(self.product_names[name])

    @Slot()
    def _confirm_add_item(self) -> None:
        try:
            self.ui.filter_table.itemChanged.disconnect(self._on_item_changed)
            self.ui.filter_table.setSortingEnabled(False)

            row = self.ui.filter_table.rowCount()
            kaspi = self.ui.kaspi_field.text().strip().lower()
            name = self.ui.name_field.text().strip().lower()
            value = str(self.ui.int_field.value()).strip()

            if not kaspi or not name:
                self.message_box.information(
                    self.ui.add_item_window, 'Предупреждение',
                    'Каспи имя не должно быть пустыми!')
            elif kaspi in self.kaspi_names_combined:
                self.message_box.information(
                    self.ui.add_item_window, 'Предупреждение',
                    '<b>Каспи имя</b> должно быть уникальным!')

            elif name in self.product_names:
                self.message_box.information(
                    self.ui.add_item_window, 'Предупреждение',
                    'Товар с таким <b>именем</b> уже существует!')

            else:
                if not name:
                    name = ''
                elif not value:
                    value = '0'

                kaspi_item = QTableWidgetItem(kaspi)
                name_item = QTableWidgetItem(name)
                value_item = IntTableWidgetItem(value)
                kaspi_item_2 = QTableWidgetItem(kaspi)
                self.ui.filter_table.setRowCount(row+1)
                self.ui.filter_table.setItem(row, 0, kaspi_item)
                self.ui.filter_table.setItem(row, 1, name_item)
                self.ui.filter_table.setItem(row, 3, value_item)
                self.ui.filter_table.setItem(row, 4, kaspi_item_2)
                self.back_up_data[kaspi] = Product(name=name, price='0',
                                                   main_kaspi_name=kaspi,
                                                   connected_kaspi_names={})
                self.kaspi_names_combined.add(kaspi)
                self.product_names[name] = kaspi

                self.message_box.information(
                    self.ui.add_item_window, 'Уведомление',
                    'Данные успешно добавлены!')

                if self.ui.combine_item_window.isVisible():
                    self._change_kaspi_names_to_connect()
                    self._populate_combine_window()

        except Exception as e:
            QMessageBox.critical(None, 'Ошибка', 'Что-то пошло не так!')
            with open(BASE_DIR / 'Logs.txt', 'w') as logs:
                logs.write(f'{traceback.format_exc()}\n{e}')
        finally:
            self.ui.filter_table.itemChanged.connect(self._on_item_changed)
            self.ui.filter_table.setSortingEnabled(True)
            self.is_changed = True

    @cool_wrapper
    @Slot()
    def _load_backup(self):

        answer = self.message_box.question(
            self, 'Уведомление', 'Вы уверены что хотите загрузить бэкап?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if answer == QMessageBox.Yes:

            first, second, third, fourth = get_last_backup_data(
                load_notlast=True)

            if any(x is None for x in (first, second, third, fourth)):
                self.message_box.information(
                    self, 'Уведомление',
                    'Бэкаппа не существует!')
                return

            (self.back_up_data, self.kaspi_names_combined,
                self.product_names,
                self.diff_names_same_product) = (first, second, third, fourth)

            self.ui.filter_table.clearContents()
            self.ui.filter_table.setRowCount(0)

            self.kaspi_names = set()
            self.incorrect_products = set()
            self.kaspi_names_to_unmerge = []
            self.is_changed = False
            self.populated = False

            self._populate_table()

    @cool_wrapper
    @Slot()
    def _delete_item_connected_list(self):
        selected_item = self.ui.kaspi_names_connected_list.currentItem()
        if selected_item:
            self.kaspi_names_to_unmerge.append(
                selected_item.text().split('||')[0].strip())
            self.ui.kaspi_names_connected_list.takeItem(
                self.ui.kaspi_names_connected_list.row(selected_item))
            if self.ui.combine_item_window.isVisible():
                self._change_kaspi_names_to_connect()

            self.is_changed = True

    @cool_wrapper
    @Slot()
    def _change_item_connected_list(self):

        selected_item = self.ui.kaspi_names_connected_list.currentItem()

        if selected_item:
            name, _, _ = selected_item.text().split('||')
            name = name.strip()
            old = self.diff_names_same_product[name]
            new = str(self.ui.amount_change_field.value())

            if self.ui.radioButton1.isChecked():
                new_type = self.ui.radioButton1.property('custom_data')
            elif self.ui.radioButton2.isChecked():
                new_type = self.ui.radioButton2.property('custom_data')
            else:
                new_type = self.ui.radioButton3.property('custom_data')

            self.back_up_data[old[0]].connected_kaspi_names[name] = (new,
                                                                     new_type)
            self.diff_names_same_product[name] = (old[0], new, new_type)

            self._open_connected_names(
                name=self.diff_names_same_product[name][0], flag=True)
            self.ui.change_amount_window.hide()

    @cool_wrapper
    @Slot()
    def _merge_products(
         self, product: Product, to_merge: str,
         row_to_delete: int, amount: int, type_amount: str) -> None:

        product_to_merge = self.back_up_data[to_merge]
        con_names_to_merge = product_to_merge.connected_kaspi_names
        product_to_merge.connected_kaspi_names = {}

        if con_names_to_merge:
            for name, values in con_names_to_merge.items():

                product.connected_kaspi_names[name] = (values[0],
                                                       values[1])
                self.diff_names_same_product[name] = (product.main_kname,
                                                      values[0], values[1])

        if product_to_merge.name:
            self.product_names.pop(product_to_merge.name)

        product.connected_kaspi_names[to_merge] = (amount, type_amount)
        self.diff_names_same_product[to_merge] = (product.main_kname, amount,
                                                  type_amount)

        self.ui.filter_table.removeRow(row_to_delete)
        self.ui.kaspi_field_to.removeItem(
            self.ui.kaspi_field_to.currentIndex())
        row = self._find_row_of_item(text=product.main_kname, column=0)

        button = QPushButton(product.main_kname)
        button.setFixedSize(QSize(49, 22))
        button.setStyleSheet("color: rgba(0, 0, 0, 0);")
        button.clicked.connect(
            lambda checked=None,
            button=button: self._open_connected_names(button=button))
        self.ui.filter_table.setCellWidget(row, 2, button)

        self.is_changed = True
        self.back_up_data.pop(to_merge)
        if self.ui.combine_item_window.isVisible():
            self._change_kaspi_names_to_connect()
            self._populate_combine_window()

    @cool_wrapper
    @Slot()
    def _unmerge_products(self, main_was_deleted: bool = False) -> None:

        products_to_umnerge: list[str] = self.kaspi_names_to_unmerge

        if products_to_umnerge:

            kaspi_key = self.diff_names_same_product[products_to_umnerge[0]][0]

            if not main_was_deleted:
                product_con_kname = self.back_up_data[
                    kaspi_key].connected_kaspi_names

                if len(product_con_kname) == len(products_to_umnerge):
                    row = self.ui.filter_table.currentRow()
                    empty = QTableWidgetItem()
                    empty.setFlags(empty.flags() & ~Qt.ItemIsEditable)
                    self.ui.filter_table.removeCellWidget(row, 2)
                    self.ui.filter_table.setItem(row, 2, empty)

            for kaspi_name_amount in products_to_umnerge:
                name = kaspi_name_amount.strip()

                if not main_was_deleted:
                    product_con_kname.pop(name)

                self.diff_names_same_product.pop(name)
                self.kaspi_names_combined.remove(name)
                self.kaspi_names.add(name)
                self.back_up_data[name] = Product(
                    main_kaspi_name=name, price='0',
                    connected_kaspi_names={}, name='')

            if not main_was_deleted:
                self.back_up_data[
                    kaspi_key].connected_kaspi_names = product_con_kname

            self.is_changed = True
            self._populate_table()
            self._populate_combine_window()
            self._change_kaspi_names_to_connect()

        self.kaspi_names_to_unmerge = []

        self.message_box.information(
            self.ui.connected_kaspi_products, 'Оповещвние',
            'Данные успешно сохранены!')

    @cool_wrapper
    @Slot()
    def _combine_products(self) -> None:

        answer = self.message_box.question(
            self.ui.combine_item_window, 'Уведомление',
            'Вы хотите соединить товар?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        if answer == QMessageBox.Yes:
            product_main_text = self.ui.product_main.currentText()
            to_merge_kaspi = self.ui.kaspi_field_to.currentText()

            if not product_main_text or not to_merge_kaspi:
                self.message_box.warning(
                    self.ui.combine_item_window, 'Предупреждение',
                    'Поля не могут быть пусты!')
            elif (product_main_text not in self.product_names
                  or to_merge_kaspi not in self.back_up_data):
                self.message_box.warning(
                    self.ui.combine_item_window, 'Предупреждение',
                    'Не валидный товар или имя!')
            else:
                product = self.back_up_data[
                    self.product_names[product_main_text]]
                amount = str(self.ui.amount_field.value()).strip()
                if self.ui.radioButton1.isChecked():
                    new_type = self.ui.radioButton1.property('custom_data')
                elif self.ui.radioButton2.isChecked():
                    new_type = self.ui.radioButton2.property('custom_data')
                else:
                    new_type = self.ui.radioButton3.property('custom_data')
                row = self._find_row_of_item(text=to_merge_kaspi, column=0)

                self._merge_products(
                    product=product, to_merge=to_merge_kaspi, amount=amount,
                    row_to_delete=row, type_amount=new_type)

    @cool_wrapper
    @Slot()
    def _delete_item(self) -> None:
        self.ui.filter_table.setSortingEnabled(False)

        row = self.ui.filter_table.currentRow()
        if row >= 0:
            name_item = self.ui.filter_table.item(row, 1)
            price_item = self.ui.filter_table.item(row, 3)
            if not name_item:
                name = ''
            else:
                name = name_item.text()
            if not price_item:
                price = '0'
            else:
                price = price_item.text()

            kaspi_name = self.ui.filter_table.item(row, 0).text()

            answer = self.message_box.question(
                self, 'Уведомление',
                '<p>Вы уверены, что хотите удалить товар?</p>'
                f'<p>Каспи название: <b>{kaspi_name}</b><br>'
                f'Название: <b>{name}</b><br>'
                f'Цена: <b>{price}</b></p>',
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

            if answer == QMessageBox.Yes:
                product = self.back_up_data.pop(kaspi_name)
                connected_names = product.connected_kaspi_names

                if connected_names:
                    k_name = ', '.join(itertools.islice(connected_names.keys(),
                                                        2))
                    answer = self.message_box.question(
                        self, 'Уведомление',
                        '<p>С данным товаром связаны каспи имена</p>'
                        '<p>Удалить их тоже?</p>'
                        '<p>Каспи имена: '
                        f'{k_name} ...<p>',
                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    if answer == QMessageBox.Yes:
                        for con_name in connected_names.keys():
                            self.kaspi_names_combined.remove(con_name)
                            self.diff_names_same_product.pop(con_name)
                    else:
                        self.kaspi_names_to_unmerge = list(
                            connected_names.keys())
                        self._unmerge_products(main_was_deleted=True)

                if kaspi_name in self.kaspi_names_combined:
                    self.kaspi_names_combined.remove(kaspi_name)

                if kaspi_name in self.incorrect_products:
                    self.incorrect_products.remove(kaspi_name)

                if name and name in self.product_names:
                    self.product_names.pop(product.name)

                self.ui.filter_table.removeRow(row)
                self.kaspi_names = set()
                self.is_changed = True
                if self.ui.combine_item_window.isVisible():
                    self._change_kaspi_names_to_connect()
                    self._populate_combine_window()
        self.ui.filter_table.setSortingEnabled(True)

    @cool_wrapper
    @Slot()
    def _excel_calculate(self) -> None:
        if self.excel_path:
            if not self.incorrect_products:
                calculate_profit_excel(
                    excel=self.excel_path, save_to=self.path_to_save_folder,
                    settings=self.user_settings, products=self.back_up_data,
                    diff_kaspi_same_product=self.diff_names_same_product)
                self.message_box.information(self, 'Уведомление',
                                             'Процесс завершен успешно!')
            else:
                self.message_box.information(
                    self, 'Уведомление',
                    'Есть товар с некоректной ценной!\n'
                    f'Каспи именна: {self.incorrect_products}')

    @cool_wrapper
    @Slot()
    def _get_new_path_for_res(self) -> None:
        directory_path = QFileDialog.getExistingDirectory(
            caption='Выберите папку',
            options=QFileDialog.ShowDirsOnly)

        if not directory_path:
            pass
        else:
            self.path_to_save_folder = Path(directory_path)

    @cool_wrapper
    @Slot()
    def _get_new_path(self) -> None:
        directory_path = QFileDialog.getOpenFileName(
            filter='Excel (*.xlsx *.xls)')[0]

        if not directory_path:
            pass
        elif not directory_path.endswith(('.xlsx', '.xls')):
            self.message_box.warning(
                self, 'Ошибка',
                'Выберите файл с расширением ".xlsx" или ".xls"',
            )
        else:
            if not Path(directory_path).exists():
                self.message_box.warning(
                    self, 'Ошибка',
                    'Фаила не существует!',
                )
            else:
                self.excel_path = Path(directory_path)
                self.kaspi_names = get_kaspi_names_excel(
                    self.excel_path,
                    row=self.user_settings['row_start'],
                    column_name=self.user_settings['column_name'])
                if (self.kaspi_names.difference(self.kaspi_names_combined)
                   or not self.populated):
                    self._populate_table()

                self.ui.file_label.setText(
                    f'<center>{directory_path}</center>')

    @cool_wrapper
    @Slot()
    def _open_change_amount_window(self) -> None:

        if self.ui.change_amount_window.isVisible():
            self.ui.change_amount_window.hide()
        else:
            selected_item = self.ui.kaspi_names_connected_list.currentItem()

            if selected_item:
                name, _, _ = selected_item.text().split('||')
                name = name.strip()
                self.ui.label_change.setText(name)
                self.ui.amount_change_field.setValue(
                    int(self.diff_names_same_product[name][1]))
                self.ui.change_amount_window.show()
            else:
                self.message_box.information(
                    self.ui.connected_kaspi_products,
                    'Уведомление', 'Товар не выбран')

    @cool_wrapper
    @Slot()
    def _open_connected_names(self, button=None, name: str = None,
                              flag: bool = False) -> None:

        self.ui.kaspi_names_connected_list.clear()
        self.kaspi_names_to_unmerge = []
        name = name or button.text()

        if self.ui.connected_kaspi_products.isVisible() and not flag:
            self.ui.connected_kaspi_products.hide()
        else:
            product = self.back_up_data[name]
            self.ui.label_connected.setText(product.name)
            self.ui.kaspi_names_connected_list.addItems(
                [f'{name_amount_type[0]} || {name_amount_type[1][0]} '
                 f'|| {AMOUNT_NAME[name_amount_type[1][1]]}'
                 for name_amount_type
                 in product.connected_kaspi_names.items()])
            self.ui.connected_kaspi_products.show()

    @cool_wrapper
    @Slot()
    def _open_combine_window(self) -> None:

        if self.ui.combine_item_window.isVisible():
            self.ui.combine_item_window.hide()
        else:
            self._populate_combine_window()
            self.ui.combine_item_window.show()

    @cool_wrapper
    @Slot()
    def _open_settings(self) -> None:

        if self.ui.setting_window.isVisible():
            self.ui.setting_window.hide()
        else:
            self.ui.setting_window.show()

    @cool_wrapper
    @Slot()
    def _open_add_item(self) -> None:

        if self.ui.add_item_window.isVisible():
            self.ui.add_item_window.hide()
        else:
            self.ui.add_item_window.show()

    @Slot()
    def _on_item_changed(self, item: QTableWidgetItem) -> None:

        try:
            self.ui.filter_table.itemChanged.disconnect(self._on_item_changed)

            if item.column() == 2 or item.column() == 4:
                return

            curr = item.text().strip().lower()
            row = item.row()
            column = item.column()
            key = self.ui.filter_table.item(row, 0).text().lower().strip()

            if key in self.incorrect_products:
                self.incorrect_products.remove(key)

            if column == 3:
                if not curr.isdecimal():
                    self.message_box.critical(
                        self, 'Ошибка',
                        'Поле с ценой должно быть числом\n'
                        f'Ряд: {row+1}')
                    self.ui.filter_table.setItem(row, column,
                                                 IntTableWidgetItem('0'))

                else:
                    product = self.back_up_data[key]
                    product.price = curr

            elif column == 1:
                old_product = self.back_up_data[key]

                if curr in self.product_names:
                    answer = self.message_box.question(
                        self, 'Оповещвние',
                        'Товар с таким именем уже существует!\n'
                        'Хотите объединить этот товар?',
                        QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                    if answer == QMessageBox.Yes:
                        amount, type_amount = self.ui._get_amount_type()
                        if amount:
                            self._merge_products(
                                product=self.back_up_data[
                                    self.product_names[curr]],
                                to_merge=key, row_to_delete=row,
                                amount=amount, type_amount=type_amount)
                        else:
                            self.ui.filter_table.setItem(
                                row, column,
                                QTableWidgetItem(old_product.name))
                    else:
                        self.ui.filter_table.setItem(
                            row, column,
                            QTableWidgetItem(old_product.name))

                else:
                    if (old_product.name and old_product.name
                       in self.product_names):
                        self.product_names.pop(old_product.name)

                    if curr:
                        self.product_names[curr] = key
                        self.ui.filter_table.item(row, column).setText(curr)
                        old_product.name = curr
                    else:
                        old_product.name = curr
                        self.ui.filter_table.setItem(
                            row, column,
                            QTableWidgetItem(''))

                if self.ui.combine_item_window.isVisible():
                    self._change_kaspi_names_to_connect()
                    self._populate_combine_window()

            elif column == 0:
                old = self.ui.filter_table.item(row,
                                                4).text().lower().strip()

                if not curr:
                    self.message_box.critical(
                        self, 'Ошибка', 'Поле не должно быть пустым!')
                    self.ui.filter_table.item(row, column).setText(old)

                elif curr in self.kaspi_names_combined:
                    self.message_box.critical(
                        self, 'Ошибка',
                        '<b>Каспи имена</b> должны быть уникальными!')
                    self.ui.filter_table.setItem(row, column,
                                                 QTableWidgetItem(old))

                else:
                    self.kaspi_names_combined.remove(old)
                    self.kaspi_names_combined.add(curr)
                    product = self.back_up_data.pop(old)

                    product.main_kname = curr
                    self.product_names[product.name] = curr
                    self.back_up_data[curr] = product

                    but_cell = self.ui.filter_table.cellWidget(row, 2)
                    if but_cell:
                        but_cell.setText(curr)

                    self.ui.filter_table.item(row, 4).setText(curr)
                    self.ui.filter_table.item(row, column).setText(curr)

                    if self.ui.combine_item_window.isVisible():
                        self.ui.main_kaspi_name.setText(
                            self.back_up_data[self.product_names[
                                self.ui.product_main.currentText()
                                ]].main_kname)
                        self._change_kaspi_names_to_connect()

        except Exception as e:
            QMessageBox.critical(None, 'Ошибка', 'Что-то пошло не так!')
            with open(BASE_DIR / 'Logs.txt', 'w') as logs:
                logs.write(f'{traceback.format_exc()}\n{e}')
        finally:
            self.ui.filter_table.itemChanged.connect(self._on_item_changed)
            self.is_changed = True

    @cool_wrapper
    @Slot()
    def _save_settings(self) -> None:

        flag: bool = False

        settings_excel = [
            (self.ui.col_main_sum.text(), 'column_main_sum'),
            (self.ui.col_expensses.text(), 'column_expenses'),
            (self.ui.col_name.text(), 'column_name'),
            (self.ui.col_amount.text(), 'column_amount'),
            (self.ui.start_rows.text(), 'row_start'),
            (self.ui.tax.text(), 'tax'),]

        for settings, settings_name in settings_excel:

            settings = settings.strip()

            if settings_name == 'column_amount' and not settings.strip():
                self.user_settings[settings_name] = settings
                continue

            if settings_name == 'column_expenses':
                orig_settings = settings
                settings = ''.join(settings.split(','))

            if (settings_name in (
                'column_main_sum', 'column_expenses',
                'column_name', 'column_amount')
               and not bool(re.search(r'^[A-Z]+$', settings))):
                self.message_box.information(
                    self.ui.setting_window, 'Предупреждение',
                    'Некоректное имя колонны!')
                break

            if settings_name in ('column_name', 'row_start'):
                if self.user_settings[settings_name] != settings:
                    flag = True
                if settings_name == 'row_start':
                    self.row_settings_before = (
                        self.user_settings[settings_name])

            if settings_name in ('tax', 'row_start'):
                settings = int(settings)

            self.user_settings[settings_name] = (
                orig_settings if settings_name == 'column_expenses'
                else settings)
        else:
            if flag and self.excel_path:
                self.kaspi_names = get_kaspi_names_excel(
                    self.excel_path,
                    row=self.user_settings['row_start'],
                    column_name=self.user_settings['column_name'])
                if (self.kaspi_names.difference(self.kaspi_names_combined)
                   or not self.populated):
                    self._populate_table()

            self.message_box.information(
                self.ui.setting_window, 'Уведомление',
                'Данные успешно сохранены!')

    @cool_wrapper
    @Slot()
    def _save_data(self) -> int:

        if not LAST_BACKUP.exists():
            open(LAST_BACKUP, 'w', encoding='utf-8').close()

        if self.is_changed:

            pre_writed_data: str = ''

            with open(LAST_BACKUP, 'r', encoding='utf-8') as f:
                pre_writed_data = f.read()

            with open(LAST_BACKUP, 'w', encoding='utf-8') as f:

                for kaspi, product in self.back_up_data.items():
                    if not product.price.isdecimal():
                        self.message_box.warning(
                            self, 'предупреждение',
                            f'Цена должна быть числом а не: {product.price}\n'
                            f'Каспи имя: {kaspi}\n'
                            f'Имя: {product.name}\n')
                        break
                    elif not kaspi.strip():
                        continue

                    values = [f'{key}|{val[0]}|{val[1]}' for key, val in
                              product.connected_kaspi_names.items()]

                    f.write(
                        f'{kaspi}||{product.name}||'
                        f'{int(product.price)}||'
                        f'{"#$#".join(values)}\n')
                else:
                    self.is_changed = False
                    self.message_box.information(
                        self, 'Уведомление', 'Данные успешно сохранены!')
                    return 1

            with open(LAST_BACKUP, 'w', encoding='utf-8') as f:
                f.write(pre_writed_data)
            return 0
        return 1

    @cool_wrapper
    @Slot()
    def _show_filter_table(self) -> None:
        if self.ui._grid_filter_table.isVisible():
            self.ui._grid_filter_table.hide()
            self.resize(self.size())
        else:
            self.ui._grid_filter_table.show()
            self.resize(self.size())

    @cool_wrapper
    @Slot()
    def _populate_combine_window(self):

        self.ui.product_main.clear()
        self.ui.product_main.addItems(self.product_names.keys())

    @Slot()
    def _populate_table(self):
        try:
            self.ui.filter_table.itemChanged.disconnect(self._on_item_changed)
            self.ui.filter_table.setSortingEnabled(False)

            flag = False

            new = self.kaspi_names.difference(self.kaspi_names_combined)
            curr_len = self.ui.filter_table.rowCount()

            if not new:
                new = self.kaspi_names_combined

            new = new.difference(set(self.diff_names_same_product.keys()))
            self.ui.filter_table.setRowCount(curr_len + len(new))

            for i, kaspi_name in enumerate(new):

                if kaspi_name not in self.back_up_data:
                    product = Product(name='', main_kaspi_name=kaspi_name,
                                      price='0', connected_kaspi_names={})
                    self.back_up_data[kaspi_name] = product
                else:
                    product = self.back_up_data[kaspi_name]

                self.kaspi_names_combined.add(kaspi_name)
                product_name = product.name.strip()
                if product_name and product_name not in self.product_names:
                    self.product_names[product.name] = kaspi_name

                if not product.price.isdecimal():
                    flag = True
                    self.incorrect_products.add(kaspi_name)

                if product.connected_kaspi_names:
                    button = QPushButton(kaspi_name)
                    button.setFixedSize(QSize(49, 22))
                    button.setStyleSheet("color: rgba(0, 0, 0, 0);")
                    button.clicked.connect(
                        lambda checked=None,
                        button=button: self._open_connected_names(
                            button=button))
                    self.ui.filter_table.setCellWidget(curr_len+i, 2, button)
                else:
                    empty = QTableWidgetItem()
                    empty.setFlags(empty.flags() & ~Qt.ItemIsEditable)
                    self.ui.filter_table.setItem(curr_len+i, 2, empty)

                kaspi_item = QTableWidgetItem(kaspi_name)
                item_item = QTableWidgetItem(product.name)
                value_item = IntTableWidgetItem(product.price)
                kaspi_item_2 = QTableWidgetItem(kaspi_name)
                self.ui.filter_table.setItem(curr_len+i, 0, kaspi_item)
                self.ui.filter_table.setItem(curr_len+i, 1, item_item)
                self.ui.filter_table.setItem(curr_len+i, 3, value_item)
                self.ui.filter_table.setItem(curr_len+i, 4, kaspi_item_2)

            if flag:
                self.message_box.warning(
                    self, 'Предупреждение',
                    'В таблице есть товар с некоректной ценной!',)

            if self.ui.combine_item_window.isVisible():
                self._change_kaspi_names_to_connect()
                self._populate_combine_window()

        except Exception as e:
            QMessageBox.critical(None, 'Ошибка', 'Что-то пошло не так!')
            with open(BASE_DIR / 'Logs.txt', 'w') as logs:
                logs.write(f'{traceback.format_exc()}\n{e}')
        finally:
            self.kaspi_names = set()
            self.is_changed = True
            self.populated = True
            self.ui.filter_table.setSortingEnabled(True)
            self.ui.filter_table.itemChanged.connect(self._on_item_changed)

    @cool_wrapper
    @Slot()
    def _filtering_table(self, search_text: str):
        search_text = self.ui.search_input.text().strip().lower()
        for row in range(self.ui.filter_table.rowCount()):
            kaspi_text = self.ui.filter_table.item(row, 0).text()
            item_text = self.ui.filter_table.item(row, 1).text()
            value_text = self.ui.filter_table.item(row, 3).text()
            if (search_text in item_text or search_text in value_text
               or search_text in kaspi_text):
                self.ui.filter_table.setRowHidden(row, False)
            else:
                self.ui.filter_table.setRowHidden(row, True)

    @cool_wrapper
    @Slot()
    def dragEnterEvent(self, event) -> None:
        if event.mimeData().hasUrls():
            event.accept()
        else:
            event.ignore()

    @cool_wrapper
    @Slot()
    def dropEvent(self, event) -> None:
        if event:
            url = event.mimeData().urls()[0]
            file_path = url.toLocalFile()
            if not file_path.endswith(('.xlsx', '.xls')):
                self.message_box.warning(
                    self, 'Ошибка',
                    'Выберите файл с расширением ".xlsx" или ".xls"',
                )
            else:
                self.excel_path = Path(file_path)
                self.kaspi_names = get_kaspi_names_excel(
                    self.excel_path,
                    row=self.user_settings['row_start'],
                    column_name=self.user_settings['column_name'])
                if (self.kaspi_names.difference(self.kaspi_names_combined)
                   or not self.populated):
                    self._populate_table()

                self.ui.file_label.setText(f'<center>{file_path}</center>')
        else:
            event.ignore()

    @Slot()
    def closeEvent(self, event) -> None:
        answer = self.message_box.question(
            self, 'Выход', 'Вы хотите закрыть программу?',
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)

        count: int = 0

        if answer == QMessageBox.Yes:
            try:
                response = self._save_data()
                count += 1

                if not response and count < 2:
                    event.ignore()
                    return

                if (DATE_CUR - datetime.strptime(
                   self.user_settings['last_backup'],
                   '%Y-%m-%d %H:%M:%S.%f')) > TO_BACKUP_TIME:
                    copy_backup = 0
                    while (BACKUP /
                            (f'{DATE_NAME}_backup_{copy_backup}.txt')
                           ).exists():
                        copy_backup += 1

                    shutil.copy(LAST_BACKUP, BACKUP /
                                f'{DATE_NAME}_backup_{copy_backup}.txt')

                    self.user_settings['last_backup'] = datetime.now()

                with open(USER_SETTINGS, 'w', encoding='utf-8') as f:
                    for key, val in self.user_settings.items():
                        f.write(f'{key}|{val}\n')

                if self.ui.setting_window.isVisible():
                    self.ui.setting_window.hide()
                if self.ui.add_item_window.isVisible():
                    self.ui.add_item_window.hide()
                if self.ui.combine_item_window.isVisible():
                    self.ui.combine_item_window.hide()
                if self.ui.connected_kaspi_products.isVisible():
                    self.ui.connected_kaspi_products.hide()
                if self.ui.change_amount_window.isVisible():
                    self.ui.change_amount_window.hide()

                event.accept()
            except Exception:
                answer = QMessageBox.critical(
                    self,
                    'Ошибка', 'Что-то пошло не так!\n Закрыть программу?',
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
                with open(BASE_DIR / 'Logs.txt',
                          'w', encoding='utf-8') as logs:
                    logs.write(traceback.format_exc())
                if answer == QMessageBox.Yes:
                    event.accept()
                else:
                    event.ignore()
        else:
            event.ignore()

    def _find_row_of_item(self, text: str, column: int) -> int:
        for row in range(self.ui.filter_table.rowCount()):
            item = self.ui.filter_table.item(row, column)
            if item is not None and item.text() == text:
                return row
