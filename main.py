from openpyxl import load_workbook
import pandas as pd
import shutil

if __name__ == "__main__":
    shutil.copy2("книга втрат електронний варіант.xlsx", "output.xlsx")

    exit()
    df = pd.read_excel("накази_втрати майна  А4007.xlsx")

    # 2. дивимось структуру
    print(df.head())
    print(df.columns)
    print(df.dtypes)