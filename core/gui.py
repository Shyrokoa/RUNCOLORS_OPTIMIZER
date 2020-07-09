import os
from tkinter import *

from core.order import Order
from core.xlsx_file import XlsxFile


class FrameGui:

    def __init__(self):
        self.window = ''
        self.enter_frame = ''
        self.gen_btm = ''
        self.subject = ''
        self.declaration = ''
        self.text_field = ''
        self.a = ''
        self.currency = ''
        self.b = ''
        self.collected_data = ''
        self.create_window()

    def create_window(self):
        self.window = Tk()
        self.window.title('Skyptimizer by Andrii Shyrokov')
        self.window.resizable(0, 0)
        self.window.geometry('{}x{}'.format(300, 166))
        self.initiate()
        self.window.mainloop()

    def initiate(self):
        # create all of the beta containers
        self.enter_frame = Frame(self.window, bg='black', width=900, height=70, pady=3)
        self.enter_frame.grid(row=0, sticky="ew")

        self.gen_btm = Button(self.enter_frame, fg='black', bg='green yellow', text='sky', width=25,
                              font='none 14 bold', command=self.generate_sky).grid(row=3, columnspan=2)
        self.text_field = Entry(self.enter_frame, width=15, text='enter warehouse', fg='BLUE')
        self.text_field.grid(column=1, row=1)
        self.a = Label(self.enter_frame, text="Warehouse:", font='none 10 bold', bg='black', fg='pink').grid(column=0,
                                                                                                             row=1)

        self.currency = Entry(self.enter_frame, width=15, text='EUR/ZL:', fg='BLUE')
        self.currency.grid(column=1, row=2)
        self.b = Label(self.enter_frame, text="EUR/ZL:", font='none 10 bold', bg='black', fg='pink').grid(column=0,
                                                                                                          row=2)

    def generate_sky(self):
        currency = self.currency.get()
        string = self.text_field.get()
        clients = []
        desktop = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')
        entries = os.listdir('{}\DATA'.format(desktop))
        self.collected_data = entries

        for a in self.collected_data:
            gg = Order(str(a), currency)
            clients.append(gg)

        file = XlsxFile(string, clients)
        file.stop()

    def generate_sub(self):
        pass

    def generate_declaration(self):
        pass
