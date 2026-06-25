from pathlib import Path
from dataclasses import dataclass
from decimal import Decimal
import datetime
from types import SimpleNamespace


from openpyxl import load_workbook
from openpyxl.worksheet.worksheet import Worksheet
from utils.core import cell_to_sqlite_date, get_order_from_comment, cell_to_decimal
from utils.row_marker import Warning_color, mark_row_wcolor

from config import wasted_backup as file_path, sample_xlsx, wasted_backup  # "data.xlsx"
sheets = ['РАО', 'ЗББ та Р', 'Зас ураж', 'НСО', 'БПЛА', 'ППО', 'ОВТ та МСП', 'РЕБ', 'Інж', 'РХБЗ', 'Реч', 'Звяз', 'Прод', 'ПММ', 'Мед', 'Авто', 'КЕС', 'Елек', 'Пожежна', 'Засоби розвідки', 'Гео', 'Метр']

# @dataclass()
# class Title_waste(init=False):
#     date: str
#     ordernum: int
#     updated: bool
#     title: str
#     titlesum: Decimal

def parse_sheet(ws: Worksheet):
    last_row = 25# ws.max_row # 8000 #9475 6078 #
    row = 7
    previous_order_num = None; set_total_sum = 0
    for row in range(row, last_row):
        date = cell_to_sqlite_date(ws.cell(row=row, column=2))
        ordernum = int(get_order_from_comment(ws.cell(row=row, column=3).value))
        if type(ordernum) == str:
            print(("Сталася помилка"))
            raise Exception("Сталася помилка")
        print(row, date, ordernum, type(ordernum))
        if (date is None or ordernum is None) and previous_order_num is None:
            print('cant define date or ordernum in row ', row)
            continue

        elif (date and ordernum) and (previous_order_num != ordernum):
            print('new set started with ordernum ', ordernum)
            if previous_order_num is not None:
                set_total_sum2 = cell_to_decimal(ws.cell(row=row-1, column=9))
                if set_total_sum != set_total_sum2:
                    print('does not equal for ', row-1, set_total_sum, set_total_sum2); exit(1)

            previous_order_num = ordernum
            set_total_sum = 0
            changed = True if ws.cell(row=row, column=2).font.strike else False
            title = ws.cell(row=row, column=4).value
            title_total_sum = cell_to_decimal(ws.cell(row=row, column=8))
            set_total_sum += title_total_sum

            print(row, date, ordernum, changed, title, title_total_sum)

        elif (date and ordernum) and (previous_order_num == ordernum and previous_order_num):
            print('set continued')
            changed = True if ws.cell(row=row, column=2).font.strike else False
            title = ws.cell(row=row, column=4).value
            title_total_sum = cell_to_decimal(ws.cell(row=row, column=8))
            set_total_sum+=title_total_sum

            print(row, date, ordernum, changed, title, title_total_sum)

            if title_total_sum is None:
                pass #print('file may be need to recalculate formulas'); exit(1)
                continue
        else:
            print('else')
            changed = True if ws.cell(row=row, column=2).font.strike else False
            title = ws.cell(row=row, column=4).value
            title_total_sum = cell_to_decimal(ws.cell(row=row, column=8))
            set_total_sum += title_total_sum

            print(row, date, ordernum, changed, title, title_total_sum)



def merged(ws: Worksheet):
    cell = "D6"
    for merged_range in ws.merged_cells.ranges:
        if cell in merged_range:
            print("Головна комірка:", merged_range.start_cell.coordinate)

if __name__ == '__main__':
    wb = load_workbook(wasted_backup, data_only=True); ws:Worksheet = wb["ЗББ та Р"]
    parse_sheet(ws)



