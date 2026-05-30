import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import sqlite3
import re
from core import is_valid_date
from datetime import datetime

start_row = 5995
sourcefile = "source.xlsx"

# Відкрити базу (або створити, якщо її немає)
conn = sqlite3.connect("data.db")
# Створити курсор
cur = conn.cursor()


def get_last_number(s="(зміни в 2431)                               3719"):
    if s is None:
        return None
    s = str(s)
    parts = s.split()
    if not parts:
        return None

    last = parts[-1]
    return int(last) if last.isdigit() else None

def find_root_order_info(ws, row):
    while True:
        row+=1
        changed = ws.cell(row=row, column=2).font.strike
        if changed:
            continue
        textdate = ws.cell(row=row, column=2).value
        operation = ws.cell(row=row, column=3).value
        order_id = get_last_number(operation)
        if not is_valid_date(textdate) or order_id is None:
            return None
        date = datetime.date(textdate)
        return date, order_id

def extract():
    wb = load_workbook(sourcefile, data_only=True); ws:Worksheet = wb["Sheet1"]
    last_row = 6011 #ws.max_row #9475
    row = 5995

    while True:
        if row > last_row:
            break

        textdate = ws.cell(row=row, column=2).value
        operation = ws.cell(row=row, column=3).value
        order_id = get_last_number(operation)

        if not is_valid_date(textdate) or order_id is None:
            row += 1
            print(row, textdate, order_id, operation)
            continue

        values = [
        ws.cell(row=row, column=col).value
        for col in range(4, 25)
        ]

        date = datetime.date(textdate)
        changed = ws.cell(row=row, column=2).font.strike
        if changed:
            root_info = find_root_order_info(ws, row)
            if root_info is None:
                print(row, "cant be processed")
                row += 1
                continue
            else:
                print(values, row, date, order_id, changed, root_info[0], root_info[1])
        else:
            print(values,row, date, order_id, changed)






        operation = ws.cell(row=row, column=3).value

        # ws.cell(row=row, column=4).value = order_id
        # ws.cell(row=row, column=5).value = effect
        # ws.cell(row=row, column=6).value = ws.cell(row=row, column=28).value

        row += 1


if __name__=="__main__":
    extract()