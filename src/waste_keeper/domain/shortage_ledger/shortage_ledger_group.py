
class Shortage_Ledger_group:
    def __init__(self, wb, sheet_name, num):
        self.wb = wb
        self.sheet_name = sheet_name
        self.num = num

    @property
    def total_amount(self):
        raise Exception('not implemented')



