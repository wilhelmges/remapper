import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import sqlite3
import re
from core import is_valid_date
from datetime import datetime

start_row = 5995
sourcefile = "source.xlsx"
headers = ['рао', 'рао збб та р', 'зас ураж', 'бпла', 'ппо', 'нсо', 'реб', 'овт та мсп', 'реч', 'інж', 'зв', 'рхбз', 'ас', 'прод', 'мед', 'пмм', 'гео', 'кес', 'елтех', 'пожежна', 'метрол']
sheets = ['БпЛА', 'ОВТ', 'ЗВ', 'ЗББ', 'ЗУ', 'РЕЧ', 'НСО (БТ)', 'Ел-тех', 'ІС', 'ГЕО', 'прод', 'пмм', 'СВТ (АС)', 'мед', 'КЕС( СІ-ІЗ)', 'Метрологія']
values = [None, None, None, None, None, 1122000.64, None, 180.05, 16947.73, None, None, None, None, None, None, None, None, None, None, None, None]

if __name__=="__main__":
    print(len(headers), len(values))


    exit()
    wb = load_workbook("output.xlsx")

    sheet_names = wb.sheetnames
    print(sheet_names);

    wb = load_workbook(sourcefile, data_only=True);
    ws: Worksheet = wb["Sheet1"]

    values = [
        ws.cell(row=1, column=col).value
        for col in range(4, 25)
    ]
    print(values)