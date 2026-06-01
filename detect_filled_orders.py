import shutil
from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
import sqlite3
import re
from core import is_valid_date
from datetime import datetime

from core import headers, department_to_sheet
import re

from tmp import get_order


def get_order(s= "№106 (ОД) 08.01.2026"):
    match = re.search(r"\d+", s)
    num = int(match.group()) if match else None
    return num

outputfile = "output.xlsx"

# Відкрити базу (або створити, якщо її немає)
conn = sqlite3.connect("wasted.db")
cur = conn.cursor()

if __name__=="__main__":
    year = 2026
    wb = load_workbook(outputfile, data_only=True)
    for ws in wb.worksheets:
        cur.execute("""
            SELECT COUNT(*)
            FROM wasted
            WHERE 
            sheet_name= ?
            AND order_date >= ?
            AND order_date < ?
        """, (f"{ws.title}",f"{year}-01-01", f"{year + 1}-01-01"))

        count = cur.fetchone()[0]
        if count<1:
            continue
        sheet_name = ws.title
        print(sheet_name, count)

        row = 8
        last_row = ws.max_row #14#
        while True:
            if row > last_row:
                break
            date = ws.cell(row=row, column=1).value
            if not is_valid_date(date):
                row += 1
                continue
            date = date.date()
            order_id = get_order(ws.cell(row=row, column=2).value)
            d = (ws.cell(row=row, column=2).value.split()[-1])
            try:
                if len(d)==10:
                    date2 = datetime.strptime(d, "%d.%m.%Y").date()
                elif len(d)==8:
                    date2 = datetime.strptime(d, "%d.%m.%y").date()
                else:
                    date2 = None; print(" cant convert date: ", repr(d), 'len ',len(d))
            except Exception as e:
                print(" cant convert date: ", d)
                date2 = date
            if date!=date2:
                print(date, date2)
                print("different dates", order_id, date, sheet_name)
            row += 1

            cur.execute(
                "UPDATE wasted SET filled = 1 WHERE order_id = ? AND order_date = ? AND sheet_name = ?",
                (order_id, date, sheet_name)
            )
            updated_rows = cur.rowcount
            if updated_rows == 0:
                print("Запис не знайдено", order_id, date, sheet_name)
            elif updated_rows == 1:
                pass #print("Оновлено 1 запис")

            else:
                print("Оновлено більше одного записа", order_id, date, sheet_name)


    conn.commit()
    conn.close()

