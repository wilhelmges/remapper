from openpyxl.utils.datetime import from_excel

from datetime import datetime, date

def remove_parentheses(text):
    result = []
    depth = 0

    for char in text:
        if char == '(':
            depth += 1
        elif char == ')':
            if depth > 0:
                depth -= 1
        elif depth == 0:
            result.append(char)

    return ' '.join(''.join(result).split())

def get_parentheses_text(s: str) -> str | None:
    start = s.find("(")
    if start == -1:
        return None

    level = 0

    for i in range(start, len(s)):
        if s[i] == "(":
            level += 1
        elif s[i] == ")":
            level -= 1

            if level == 0:
                return s[start:i + 1]

    return None

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

def cell_to_sqlite_date(value) -> str | None:
    """
    Перетворює будь-яке значення в дату формату YYYY-MM-DD.

    Повертає:
        '2024-09-02' або None.
    """

    # порожня ячейка
    if value is None:
        return None
    # datetime
    if isinstance(value, datetime):
        return value.date().isoformat()
    # date
    if isinstance(value, date):
        return value.isoformat()

    # Excel serial date (число)
    if isinstance(value, (int, float)):
        try:
            return from_excel(value).date().isoformat()
        except Exception:
            return None

    # рядок
    if isinstance(value, str):

        value = value.strip()

        if not value:
            return None

        formats = (
            "%d.%m.%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            "%d/%m/%y",
        )

        for fmt in formats:
            try:
                return datetime.strptime(value, fmt).date().isoformat()
            except ValueError:
                pass

        return None

    return None

def value_to_sqlite_date(value) -> str | None:
    """
    Перетворює будь-яке значення в дату формату YYYY-MM-DD.

    Повертає:
        '2024-09-02' або None.
    """
    # порожня ячейка
    if value is None:
        return None
    # datetime
    if isinstance(value, datetime):
        return value.date().isoformat()
    # date
    if isinstance(value, date):
        return value.isoformat()

    # Excel serial date (число)
    if isinstance(value, (int, float)):
        try:
            return from_excel(value).date().isoformat()
        except Exception:
            return None

    # рядок
    if isinstance(value, str):

        value = value.strip()

        if not value:
            return None

        formats = (
            "%d.%m.%Y",
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
            '%d/%m/%y',
        )

        for fmt in formats:
            try:
                return datetime.strptime(value, fmt).date().isoformat()
            except ValueError:
                pass

        return None

    return None

def value_to_pydate(value)->date:
    return date.fromisoformat(value_to_sqlite_date(value))

def safe_int(value: str) -> int | None:
    if not isinstance(value, str):
        return None

    value = value.lstrip()

    digits = ""
    for char in value:
        if not char.isdigit():
            break
        digits += char

    return int(digits) if digits else None

def remove_duplicate_spaces(text: str) -> str:
    return " ".join(text.split())

#date oh sheet names 24/09/2026
sheets = ['ЗББ та Р', 'Зас ураж', 'НСО', 'БПЛА', 'ППО', 'ОВТ та МСП', 'РЕБ', 'Реч', 'Звяз', 'Прод', 'ПММ', 'Мед', 'Авто', 'КЕС', 'Елек', 'Інженерна', 'Пожежна', 'Засоби розвідки', 'Гео', 'Метр', 'РАО', 'РХБЗ', 'Інж']

