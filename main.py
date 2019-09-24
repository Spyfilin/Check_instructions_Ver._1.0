from tkinter import *
from tkinter import messagebox
import xlrd
import datetime

window = Tk()
window.withdraw()  # Сокрытие основного окна
messagebox.showinfo('Программа проверки актуальности инструкций Ver. 1.0', 'Если файл base_of_instructions.xlsx '
                                                                           'находится не в папке программы и не '
                                                                           'заполнен по образцу, как указано в Листе 2,'
                                                                           ' программа будет работать неправильно! В '
                                                                           'данной версии нет защиты от "дурака".'
                                                                           '\n\n© Яргункин Е.А., 2019')

date_today = datetime.date.today()

excel_data_file = xlrd.open_workbook('./base_of_instructions.xlsx')
sheet = excel_data_file.sheet_by_index(0)
row_number = sheet.nrows

bad_instruction = ""

if row_number > 0:
    for row in range(0, row_number):
        if str(sheet.row(row)[0]).replace("text:", "").replace("'", "") < str(date_today):
            bad_instruction += str(sheet.row(row)[0]).replace("text:", "").replace("'", "") + ' - ' \
                               + str(sheet.row(row)[1]).replace("text:", "").replace("'", "") + "\n"
    if bad_instruction != "":
        messagebox.showwarning('Инструкции с истекшим сроком', bad_instruction)
    else:
        messagebox.showinfo('Инструкции с истекшим сроком', 'Все инструкции актуальны!')
else:
    messagebox.showerror('Ошибка файла base_of_instructions.xlsx', 'Файл пустой или заполнен неверно.\n'
                                                                   'См. правильность заполнения файла '
                                                                   'base_of_instructions.xlsx на Листе 2.')
