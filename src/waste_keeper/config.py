import logging
from pathlib import Path

PROJECT_DIR = Path(__file__).resolve().parent.parent.parent
db_path = PROJECT_DIR /"data"/"wasted.db"

property_loss_orders_readonly = r'S:\Втрати майна\1._ Книга втрат та нестач А 4007\В_Ч А4007\накази_втрати майна  А4007.xlsx'
tmp_property_loss_orders = PROJECT_DIR /"data"/'tmp-накази_втрати майна  А4007.xlsx'

shortage_ledger_readonly = r'S:\Втрати майна\1._ Книга втрат та нестач А 4007\В_Ч А4007\Книга обліку нестач.xlsx'
tmp_shortage_ledger = PROJECT_DIR /"data"/'tmp-Книга обліку нестач.xlsx'

electronic_loss_book_readonly = r'S:\Втрати майна\1._ Книга втрат та нестач А 4007\В_Ч А4007\книга втрат електронний варіант 25 вересня 2026.xlsx'
tmp_electronic_loss_book = PROJECT_DIR /"data"/'tmp-книга втрат електронний варіант.xlsx'

# orders_wide = r"order_tracer\backups\накази_втрати майна  А4007.xlsx"
# sourcefile = "tmp-накази_втрати майна  А4007.xlsx"
# outputfile = "книга втрат електронний варіант.xlsx"
# wasted_network_url = r"S:\Втрати майна\Книга обліку нестач.xlsx"
#
# wasted_backup = r"C:\progs\remapper\order_tracer\backups\Книга обліку нестач.xlsx"
sample_xlsx = r"C:\progs\remapper\order_tracer\backups\sample.xlsx"

import logging
from logging.handlers import RotatingFileHandler

LOG_FILE = Path("backup.log")
logging.basicConfig(
    filename=LOG_FILE,
    encoding="utf-8",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)
