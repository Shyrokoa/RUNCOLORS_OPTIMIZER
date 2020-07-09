import os

from bs4 import BeautifulSoup


class Order:

    def __init__(self, file, currency):
        self.currency = currency
        self.file = file
        self.soup = ''
        self.table = ''
        self.unique_products = ''
        self.delivery_cost = ''
        self.total_price_uah = ''
        self.pln_uan_currency = ''
        self.exist_code = ''
        self.product_index = 0
        self.product_name = ''
        self.product_title = ''
        self.title = ''
        self.product_symbol = ''
        self.product_quantity = ''
        self.product_price_zl = ''
        self.product_price_eur = ''
        self.PRODUCT_DATA = []
        self.PRICE_DATA = []
        self.CLIENT_DATA = []
        self.tmp = 0
        self.client = ''
        self.address = ''
        self.apartment = ''
        self.comment = ''
        self.email = ''
        self.phone = ''
        self.index = ''
        self.town = ''
        self.open_html()
        self.check_quantity()
        self.read_delivery_cost()
        self.read_total_price_uah()
        self.get_exist_code()
        self.read_products()
        self.read_client_data()

    def open_html(self):
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop') + f'\DATA\{self.file}'
        html_file = open(desktop, 'r', encoding='utf-8')
        source_code = html_file.read()
        self.soup = BeautifulSoup(source_code, 'html.parser')
        self.table = self.soup.find('table', class_='objectListing objectListing--bordered')

    # method check how many unique products order consists
    def check_quantity(self):
        quantity = 0
        while True:
            row = self.table.find_all('tr')[quantity + 1]
            col = row.find_all('td')
            # first line in runcolors order
            if col[0].text.isdigit():
                quantity += 1
            else:
                self.unique_products = quantity - 1
                break

    # read Delivey Cost
    def read_delivery_cost(self):
        row = self.table.find_all('tr')[self.unique_products + 1]
        col = row.find_all('td')
        self.delivery_cost = float(col[9].text.split(' ')[0].replace(',', '.'))

    def read_total_price_uah(self):
        page_content = self.soup.find('table', class_="objectDetails")
        self.total_price_uah = round(float(page_content.find_all('td')[3].text.split(' ')[1].replace(',', '.')))
        self.pln_uan_currency = float(page_content.find_all('td')[5].text.split(' ')[0].replace(',', '.'))

    def get_exist_code(self):
        if self.delivery_cost == 19.99:
            self.exist_code = 'YES'
        elif self.delivery_cost == 14.99:
            self.exist_code = 'NO'
        else:
            self.exist_code = 'ERROR'

    def read_products(self):
        for i in range(self.unique_products):
            self.product_index += 1
            self.read_single_product(i + 1)

    def read_single_product(self, line_number):
        row = self.table.find_all('tr')[line_number]
        col = row.find_all('td')
        self.product_name = col[2].a.text
        self.concat_title()
        self.product_symbol = ''.join(col[2].div.text.split(' ')[1:])
        self.product_quantity = ''.join(col[3].text.split(' ')[:1])
        self.product_price_zl = col[9].text.split(' ')[0].replace(',', '.')
        self.product_price_eur = round(float(self.product_price_zl) / float(self.currency), 2)
        self.product_price_uah = self.convert_to_uah()
        self.add_product_data()

    def concat_title(self):
        self.is_shoes()
        self.title = self.product_title + self.product_name

    def is_shoes(self):
        a = self.product_name.split('/')
        if len(a) > 1:
            self.product_title = 'BUTY SPORTOWE '
        else:
            pass

    def convert_to_uah(self):
        product_price_uah = 0
        if self.exist_code == 'NO':
            pass
        elif self.exist_code == 'YES':
            if self.unique_products == 1:
                product_price_uah = self.total_price_uah
            else:
                if self.product_index != self.unique_products:
                    product_price_uah = round(float(self.product_price_zl) / self.pln_uan_currency)
                    self.PRICE_DATA.append(product_price_uah)

                else:
                    self.sum_in_uah_without_last()
                    product_price_uah = self.total_price_uah - self.tmp
        else:
            product_price_uah = 'ERROR'
        return product_price_uah

    def sum_in_uah_without_last(self):
        for i in self.PRICE_DATA:
            self.tmp += i

    def add_product_data(self):
        self.PRODUCT_DATA.append({
            "symbol": self.product_symbol,
            "quantity": self.product_quantity,
            "price_zl": self.product_price_zl,
            "price_eur": self.product_price_eur,
            'CURRENCY_EUR': 'EUR',
            'exist_code': self.exist_code,
            'uah price by product': self.product_price_uah,
            'CURRENCY_UAH': 'UAH',
            'Title': self.title,
            'weight': 1.00
        })

    def read_client_data(self):
        client_data = self.soup.find('fieldset', class_="form-fieldset")
        self.read_full_name(client_data)
        self.read_full_address(client_data)
        self.read_comments()
        self.read_email()
        self.read_phone()
        self.read_index(client_data)
        self.read_town(client_data)
        self.add_client_data()

    def read_full_name(self, data):
        self.client = data.select_one('#order_customer_delivery_full_name')["value"].upper()

    def read_full_address(self, data):
        street = data.select_one('#order_customer_delivery_street_name')['value']
        house = data.select_one('#order_customer_delivery_building_number')['value']
        self.address = ', '.join([street, house])
        self.apartment = data.select_one('#order_customer_delivery_apartment_number')['value']

    def read_comments(self):
        office_frame = \
            self.soup.find_all('div', class_='form-group')[1].select_one('#order_customer_delivery_company_name')[
                'value']
        if not self.soup.find_all('fieldset', class_='form-fieldset')[7].find_all('p'):
            comment = ''
        else:
            comment = ' '.join(self.soup.find_all('fieldset', class_='form-fieldset')[7].find_all('p')[0].text).replace(
                "\t", "")
        self.formate_comments(office_frame, comment)

    def formate_comments(self, office, msg):
        if len(office + msg) != 0:
            self.comment = msg.replace("  ", "~!~").replace(" ", "").replace("~!~", " ") + ' / ' + office
        else:
            self.comment = ''

    def read_email(self):
        self.email = self.soup.select_one('#order_customer_email')["value"]

    def read_phone(self):
        phone = self.soup.select_one('#order_customer_telephone')["value"]
        phone_number = []
        for i in list(phone):
            if i.isnumeric():
                phone_number.append(i)
        try:
            self.phone = '+38' + ''.join(phone_number[-10:])
        except:
            self.phone = 'ERROR'

    def read_index(self, data):
        self.index = data.select_one('#order_customer_delivery_postal_code')['value']

    def read_town(self, data):
        self.town = data.select_one('#order_customer_delivery_city')['value']

    def add_client_data(self):
        self.CLIENT_DATA = {
            'client': self.client,
            'apartment': self.apartment,
            'index': self.index,
            'town': self.town,
            'address': self.address,
            'comment': self.comment,
            'email': self.email,
            'phone': self.phone,
            'total_uan_price': self.total_price_uah,
            'pln_uan_currency': self.pln_uan_currency,
            'FS': 'FS'
        }
