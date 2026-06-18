import logging
from pathlib import Path

orders_network_url = r"S:\Втрати майна\1._ Книга втрат та нестач А 4007\В_Ч А4007\накази_втрати майна  А4007.xlsx"
sourcefile = "tmp-накази_втрати майна  А4007.xlsx"
outputfile = "книга втрат електронний варіант.xlsx"

LOG_FILE = Path("backup.log")
logging.basicConfig(
    filename=LOG_FILE,
    encoding="utf-8",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)
