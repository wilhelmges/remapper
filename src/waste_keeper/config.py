import logging
from pathlib import Path

orders_wide = r"order_tracer\backups\накази_втрати майна  А4007.xlsx"
sourcefile = "tmp-накази_втрати майна  А4007.xlsx"
outputfile = "книга втрат електронний варіант.xlsx"
wasted_network_url = r"S:\Втрати майна\Книга обліку нестач.xlsx"

wasted_backup = r"C:\progs\remapper\order_tracer\backups\Книга обліку нестач.xlsx"
sample_xlsx = r"C:\progs\remapper\order_tracer\backups\sample.xlsx"

LOG_FILE = Path("backup.log")
logging.basicConfig(
    filename=LOG_FILE,
    encoding="utf-8",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)
