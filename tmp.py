import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import sqlite3
import re
from core import is_valid_date
from datetime import datetime
import re
import re

def get_order(s= "№106 (ОД) 08.01.2026"):
    match = re.search(r"\d+", s)
    num = int(match.group()) if match else None
    return num


# start_row = 5995
# sourcefile = "source.xlsx"
# headers = ['рао', 'рао збб та р', 'зас ураж', 'бпла', 'ппо', 'нсо', 'реб', 'овт та мсп', 'реч', 'інж', 'зв', 'рхбз', 'ас', 'прод', 'мед', 'пмм', 'гео', 'кес', 'елтех', 'пожежна', 'метрол']
# sheets = ['БпЛА', 'ОВТ', 'ЗВ', 'ЗББ', 'ЗУ', 'РЕЧ', 'НСО (БТ)', 'Ел-тех', 'ІС', 'ГЕО', 'прод', 'пмм', 'СВТ (АС)', 'мед', 'КЕС( СІ-ІЗ)', 'Метрологія']
# values = [None, None, None, None, None, 1122000.64, None, 180.05, 16947.73, None, None, None, None, None, None, None, None, None, None, None, None]

if __name__=="__main__":
    vs = [None, 24015.24, 553295.36, None, 1479304.3, None, None, None, None, None, None, None, None, None, None, None, None,
     None, None, None, None]
    for i,v in enumerate(vs):
        print(i,v)


