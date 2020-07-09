import os

import xlsxwriter


class XlsxFile:

    def __init__(self, text, orders_list):
        self.text = text
        self.orders_list = orders_list
        self.sky_header = ''
        self.new_post = ''
        self.sky_workbook = ''
        self.sky_sheet = ''
        self.client_data = ''
        self.product_data = ''
        self.file_data = ''
        self.empty_line_index = 1
        self.cell_format = ''
        self.initialize_sky_header()
        self.initialize_new_post()
        self.create_sky_xlsx()
        self.create_header()
        self.reading()

    def initialize_sky_header(self):
        self.sky_header = ['Область', 'Город', 'Адрес',
                           'Квартира', 'ФИО получателя', 'Номер телефона',
                           'Номер счета продажи', 'Декларируемая стоимость',
                           'Валюта', 'Exist COD?', 'Сумма наложеного платежа (UAH)',
                           'Валюта', 'Наименование', 'Количество', 'Вес']

    def initialize_new_post(self):
        self.new_post = ['НОВА', 'ПОШТА', 'НОВАЯ', 'ОТДЕЛЕНИЕ',
                         'ПОЧТА', 'ОТДЕЛЕНИЯ', 'ВІДДІЛЕННЯ',
                         'НОВОЙ', 'НОВОЇ', 'ПОШТИ', 'ПОЧТЫ',
                         'ПОЧТОЙ', 'ПОЧТОЮ', 'НОВОЮ', 'НОВУЮ',
                         'НОВУ', 'NOVA', 'POSHTA', 'NOVAYA', 'POCHTA']

    def create_sky_xlsx(self):
        desktop_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        self.sky_workbook = xlsxwriter.Workbook(r'{}\{}'.format(desktop_path, f'{self.text}_sky.xlsx'))
        self.sky_sheet = self.sky_workbook.add_worksheet()

    def create_header(self):
        for index, header in enumerate(self.sky_header):
            self.sky_sheet.write(0, index, header)

    def read_order_data(self, order):
        self.client_data = order.CLIENT_DATA
        self.product_data = order.PRODUCT_DATA

    def reset_data_order(self):
        self.client_data = ''
        self.product_data = ''

    def insert_order(self):
        self.insert_index()
        self.insert_town()
        self.insert_address()
        self.insert_apartment()
        self.insert_client()
        self.insert_phone()
        self.insert_fs()
        self.insert_comment()
        self.insert_products()

    def reading(self):
        for order in self.orders_list:
            self.read_order_data(order)
            self.insert_order()
            self.reset_data_order()

    def insert_index(self):
        self.sky_sheet.write(self.empty_line_index, 0, self.client_data.get('index'))

    def insert_town(self):
        self.insert_item_np('town', 1)

    def insert_address(self):
        self.insert_item_np('address', 2)

    def insert_apartment(self):
        self.insert_item_np('apartment', 3)

    def insert_client(self):
        self.insert_item('client', 4)

    def insert_phone(self):
        self.insert_item('phone', 5)

    def insert_fs(self):
        self.insert_item('FS', 6)

    def insert_comment(self):
        self.insert_item_np('comment', 15)
        if self.client_data.get('comment') == '' and not self.is_new_post('comment'):
            self.cell_format_white()
            self.sky_sheet.write(self.empty_line_index, 15, self.client_data.get('comment'), self.cell_format)
        elif self.client_data.get('comment') != '' and not self.is_new_post('comment'):
            self.cell_format_green()
            self.sky_sheet.write(self.empty_line_index, 15, self.client_data.get('comment'), self.cell_format)

    def insert_products(self):
        for product in self.product_data:
            self.sky_sheet.write(self.empty_line_index, 7, product.get('price_eur'))
            self.sky_sheet.write(self.empty_line_index, 8, 'EUR')
            self.sky_sheet.write(self.empty_line_index, 9, product.get('exist_code'))
            self.sky_sheet.write(self.empty_line_index, 10, product.get('uah price by product'))
            self.sky_sheet.write(self.empty_line_index, 11, 'UAH')
            self.sky_sheet.write(self.empty_line_index, 12, product.get('Title'))
            self.sky_sheet.write(self.empty_line_index, 13, product.get('quantity'))
            self.sky_sheet.write(self.empty_line_index, 14, 1.00)
            self.empty_line_index += 1

    def insert_item_np(self, marker, index):
        if self.is_new_post(marker):
            self.sky_sheet.write(self.empty_line_index, index, self.client_data.get(marker), self.cell_format)
        else:
            self.sky_sheet.write(self.empty_line_index, index, self.client_data.get(marker))

    def insert_item(self, marker, index):
        self.sky_sheet.write(self.empty_line_index, index, self.client_data.get(marker))

    def split_data(self, marker):
        for r in ((',', ' '), ('.', ' '), (':', ' ')):
            string = self.client_data.get(f'{marker}').replace(*r)
        return string.upper().split()

    def is_new_post(self, marker):
        if any(item in self.split_data(marker) for item in self.new_post):
            self.cell_format_red()
            return True
        else:
            return False

    def cell_format_red(self):
        self.cell_format = self.sky_workbook.add_format({'fg_color': 'red'})

    def cell_format_green(self):
        self.cell_format = self.sky_workbook.add_format({'fg_color': 'green'})

    def cell_format_white(self):
        self.cell_format = self.sky_workbook.add_format({'fg_color': 'white'})

    def stop(self):
        self.sky_workbook.close()
