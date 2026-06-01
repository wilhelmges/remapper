import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import sqlite3
import re
from core import is_valid_date
from datetime import datetime

from core import headers, department_to_sheet

start_row = 5995
sourcefile = "source.xlsx"

# Відкрити базу (або створити, якщо її немає)
conn = sqlite3.connect("wasted.db")
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
        return order_id, date

def extract():
    cur.execute("DELETE FROM wasted");  conn.commit()
    wb = load_workbook(sourcefile, data_only=True); ws:Worksheet = wb["Sheet1"]
    last_row = ws.max_row #9475 6055 #
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

        changed = None; root_order_id = None; root_date = None
        date = datetime.date(textdate)
        changed = ws.cell(row=row, column=2).font.strike
        if changed:
            root_info = find_root_order_info(ws, row)
            root_order_id, root_date = root_info
            if root_info is None:
                print(row, "cant be processed")
                row += 1
                continue
            else:
                print(row, order_id, date, changed, root_order_id, root_date, values )
        else:
            print(row, order_id, date, changed, root_order_id, root_date, values )

        for i, value in enumerate(values):
            sheet_name = department_to_sheet[headers[i]]
            if value is not None and sheet_name is not None:
                pass
                cur.execute(
                    "INSERT INTO wasted (order_id, order_date, root_order_id, root_date, department,sheet_name, money, row) "
                    "VALUES (?, ?, ?, ?, ?,?, ?, ?)",
                    (order_id, date, root_order_id, root_date, headers[i],sheet_name, value, row)
                )

        operation = ws.cell(row=row, column=3).value

        row += 1

    conn.commit(); conn.close()


if __name__=="__main__":
    extract()