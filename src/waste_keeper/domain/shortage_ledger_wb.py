from openpyxl import load_workbook

from waste_keeper.domain.shartage_ledger_group_status import StatusDetector
from waste_keeper.domain.shortage_ledger_group import Shortage_Ledger_group
from waste_keeper.waste_status import Waste_status

class ShortageLedger:
    def __init__(self,file_url):
        self.file_url = file_url

from waste_keeper.domain.core import sheets

if __name__=='__main__':
    from waste_keeper.config import tmp_property_loss_orders, tmp_shortage_ledger
    wb = load_workbook(tmp_shortage_ledger, data_only=True)

    sheets = ['ЗББ та Р', 'Зас ураж', 'БПЛА', 'ППО', 'ОВТ та МСП', 'РЕБ', 'РХБЗ', 'Авто', 'КЕС', 'Елек', 'Пожежна',
              'Засоби розвідки', 'Гео', ]
    perverted = ["НСО","Інж","Реч","Звяз", "Прод",'ПММ', 'Мед']
    sheets = ["Авто"]

    for sheet in sheets:
        print(sheet)
        ws = wb[sheet]
        num = 3
        while True:
            slg = Shortage_Ledger_group(ws, num)
            if slg.isgroup:
                break
            num += 1
        while True:
            print(f'in wb {num}')
            sd = StatusDetector(slg.description)
            if sd.waste_status == Waste_status.ETC:
                print(slg.ordernum,sd.waste_status.value, slg.description)
            num = slg.next_groupnum
            if not num:
                break
            slg = Shortage_Ledger_group(ws, num)
