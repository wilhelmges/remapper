import shutil

from waste_keeper.config import (property_loss_orders_readonly, tmp_property_loss_orders,
                                 shortage_ledger_readonly, tmp_shortage_ledger,
                                electronic_loss_book_readonly, tmp_electronic_loss_book)

def get_last_excel_files():
    shutil.copy2(property_loss_orders_readonly, tmp_property_loss_orders)
    shutil.copy2(shortage_ledger_readonly, tmp_shortage_ledger)
    shutil.copy2(electronic_loss_book_readonly, tmp_electronic_loss_book)

if __name__ == '__main__':
    get_last_excel_files()