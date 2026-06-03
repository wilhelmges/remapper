from openpyxl.styles import PatternFill

from enum import StrEnum

class WorningColor(StrEnum):
    ORDER_NOT_FOUND = "D9D9D9"
    MULTIPLE_RECORD = "A9D9A9"
    DIFFERENT_DATES = "0000FF"

def mark_neutral(ws, row, fgColor="D9D9D9"):
    gray_fill = PatternFill(
        fill_type="solid",
        fgColor=fgColor  # світло-сірий
    )

    for cell in ws[row]:
        cell.fill = gray_fill

