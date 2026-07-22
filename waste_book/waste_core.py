from enum import Enum

from openpyxl.workbook import Workbook
from utils.core import remove_duplicate_spaces

sheets = ['РАО', 'ЗББ та Р', 'Зас ураж', 'НСО', 'БПЛА', 'ППО', 'ОВТ та МСП', 'РЕБ', 'Інж', 'РХБЗ', 'Реч', 'Звяз', 'Прод', 'ПММ', 'Мед', 'Авто', 'КЕС', 'Елек', 'Пожежна', 'Засоби розвідки', 'Гео', 'Метр']
sheets_normal = [ 'ЗББ та Р','Зас ураж','НСО','БПЛА', 'ППО', 'ОВТ та МСП', 'РЕБ', 'Інж', 'РХБЗ', 'Реч', 'Звяз', 'Прод', 'ПММ', 'Мед', 'Авто', 'КЕС', 'Елек', 'Пожежна', 'Засоби розвідки', 'Гео', ]
sheets_perversed=['РАО','Метр']

def get_status_for_row(wb: Workbook, sheetname, row):
    if sheetname not in sheets:
        raise Exception('cant find a sheet {}'.format(sheetname))
    if not wb.data_only:
        raise Exception('unmaintainable file mode')

    ws = wb[sheetname]

class Waste_status(Enum):
    BLANK = 'blank'
    CANCELED = "canceled"
    UPDATED = "updated"
    WROTEOFF = "wroteoff"
    GUILTED = "guilted"
    INQUIRY = "inquiry"
    ETC = "etc"

def waste_status(value:str):
    if value is None:
        return Waste_status.BLANK
    if not isinstance(value, str):
        return Waste_status.ETC
    value = remove_duplicate_spaces(value.strip().lower())
    if value.startswith("єас №"):
        return Waste_status.WROTEOFF
    elif value.startswith("зміни") or  value.startswith("змінено"):
        return Waste_status.UPDATED
    elif value.startswith("скасування") or "втрату" in value or "втратив" in value:
        return Waste_status.CANCELED
    elif "розслідування" in value or "рослідування" in value or "розслідуваня" in value:
        return Waste_status.INQUIRY
    elif ("винн" in value) or ("стягнути" in value):
        return Waste_status.GUILTED
    elif len(value) > 1:
        return Waste_status.ETC
    else:
        return Waste_status.BLANK

if __name__=='__main__':
    print(waste_status('вважати таким що втратив чинність на підставі наказу №907 від 10.02.2026').value)

#стягнути з Бухнацевич Д.В.
