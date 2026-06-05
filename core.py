from datetime import datetime, date
import re

sourcefile = "накази_втрати майна  А4007.xlsx"
outputfile = "книга втрат електронний варіант.xlsx"


def is_valid_date(value):
    if value is None:
        return False

    if isinstance(value, (datetime, date)):
        return True

    if not isinstance(value, str):
        return False

    value = value.strip()
    if not value:
        return False

    formats = (
        "%d.%m.%Y",
        "%d.%m.%y",
        "%Y-%m-%d",
        "%d/%m/%Y",
        "%m/%d/%Y",
    )

    for fmt in formats:
        try:
            datetime.strptime(value, fmt)
            return True
        except ValueError:
            pass

    return False

headers = ['рао', 'рао збб та р', 'зас ураж', 'бпла', 'ппо', 'нсо', 'реб', 'овт та мсп', 'реч', 'інж', 'зв', 'рхбз', 'ас', 'прод', 'мед', 'пмм', 'гео', 'кес', 'елтех', 'пожежна', 'метрол']
sheets = ['БпЛА', 'ОВТ', 'ЗВ', 'ЗББ', 'ЗУ', 'РЕЧ', 'НСО (БТ)', 'Ел-тех', 'ІС', 'ГЕО', 'прод', 'пмм', 'СВТ (АС)', 'мед', 'КЕС( СІ-ІЗ)', 'Метрологія']

department_to_sheet = {
    None: None,
    "рао": None,
    "рао збб та р": "ЗББ",
    "зас ураж": "ЗУ",
    "бпла": "БпЛА",
    "ппо": None,
    "нсо": "НСО (БТ)",
    "реб": None,
    "овт та мсп": "ОВТ",
    "реч": "РЕЧ",
    "інж": "ІС",
    "зв": "ЗВ",
    "рхбз": None,
    "ас": "СВТ (АС)",
    "прод": "прод",
    "мед": "мед",
    "пмм": "пмм",
    "гео": "ГЕО",
    "кес": "КЕС( СІ-ІЗ)",
    "елтех": "Ел-тех",
    "пожежна": None,
    "метрол": "Метрологія",
}

def get_order_from_comment(s="(зміни в 2431)                               3719"):
    operation = str(s).strip()
    if operation.isdigit():
        order_id = int(operation)
    else:
        order_id = re.sub(r'\([^)]*\)', '', operation).strip()
        order_id = order_id.split()[-1]
    return order_id

def format_date_for_output(date_str):
    return datetime.strptime(date_str, "%Y-%m-%d").strftime("%d.%m.%Y")

if __name__=='__main__':
    print(format_date_for_output('2026-03-17'))
