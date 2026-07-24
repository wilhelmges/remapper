from Waste_row import Waste_row
from Raowaste_row import Raowaste_row

def row_factory(sheetname:str):
    if sheetname == 'РАО':
        return Raowaste_row
    return Waste_row
