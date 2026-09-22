import win32com.client
from config import wasted_backup

def recalc(filepath):
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    wb = excel.Workbooks.Open(r"C:\progs\remapper\order_tracer\backups\sample.xlsx")
    excel.CalculateFull()
    wb.Save()
    wb.Close()
    excel.Quit()

if __name__ == "__main__":
    recalc(wasted_backup)