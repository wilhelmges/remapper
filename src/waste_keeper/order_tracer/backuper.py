import shutil
import sqlite3
from datetime import datetime
from pathlib import Path
from utils.core import calculate_md5
from config import orders_network_url
import logging
LOG_FILE = Path("backup.log")
logging.basicConfig(
    filename=LOG_FILE,
    encoding="utf-8",
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)
logger = logging.getLogger(__name__)

SOURCE_FILE = Path(orders_network_url)

BACKUP_DIR = Path("backups")
BACKUP_DIR.mkdir(exist_ok=True)
DB_FILE = Path("backup.db")

def init_db():
    conn = sqlite3.connect(DB_FILE)

    conn.execute("""
        CREATE TABLE IF NOT EXISTS backups (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            backup_datetime TEXT NOT NULL,
            signature TEXT NOT NULL UNIQUE
        )
    """)

    conn.commit()
    return

def get_db():
    conn = sqlite3.connect(DB_FILE)
    return conn

def get_last_signature(conn) -> str | None:
    cur = conn.execute("""
        SELECT signature
        FROM backups
        ORDER BY id DESC
        LIMIT 1
    """)
    row = cur.fetchone()
    return row[0] if row else None

def save_backup(file_path: Path, signature: str):
    short_sig = signature[:8]
    backup_name = (f"{file_path.stem}_{short_sig}.xlsx")
    backup_path = BACKUP_DIR / backup_name
    shutil.copy2(SOURCE_FILE, backup_path)

    return backup_path

def register_backup(conn, signature: str):
    conn.execute("""
        INSERT INTO backups (
            backup_datetime,
            signature
        )
        VALUES (?, ?)
    """, (
        datetime.now().isoformat(
            sep=" ",
            timespec="seconds"
        ),
        signature
    ))

    conn.commit()

def backuper():
    conn = get_db()
    if not SOURCE_FILE.exists():
        logger.error(
            f"Файл не знайдено: {SOURCE_FILE}"
        )
        exit(1)

    current_signature = calculate_md5(SOURCE_FILE)
    logger.info('backuped with hash '+current_signature)

    conn = get_db()
    try:
        last_signature = get_last_signature(conn)
        if current_signature == last_signature:
            logger.info("Змін не виявлено");print("Змін не виявлено")
            exit()

        backup_path = save_backup(SOURCE_FILE, current_signature)
        register_backup(conn, current_signature)
        logger.info(f"Створено бекап: {backup_path}")

    except  Exception as e:
        print(str(e))
    finally:
        conn.close()

if __name__ == "__main__":
    backuper()

