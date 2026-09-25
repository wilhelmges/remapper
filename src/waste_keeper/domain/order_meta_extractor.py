from openpyxl.worksheet.worksheet import Worksheet


class OrderMetaExtractor:
    def __init__(self, ws: Worksheet, row: int):
        self.ws = ws
        self.row = row

    @property
    def isorder(self):
        descr = self.ws.cell(row=self.row, column=2).value
        if descr is None:
            return False
        if is_valid_date(descr):
            return True
        return False