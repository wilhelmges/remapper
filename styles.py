from openpyxl.styles import PatternFill

def mark_neutral(ws, row):
    gray_fill = PatternFill(
        fill_type="solid",
        fgColor="D9D9D9"  # світло-сірий
    )

    for cell in ws[row]:
        cell.fill = gray_fill