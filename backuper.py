import hashlib
import logging
import shutil
import sqlite3
import zipfile
from datetime import datetime
from pathlib import Path


SOURCE_FILE = Path(
    r"S:\Втрати майна\1._ Книга втрат та нестач А 4007\В_Ч А4007\накази_втрати майна  А4007.xlsx"
)

BACKUP_DIR = Path("backups")
DB_FILE = Path("backup.db")
LOG_FILE = Path("backup.log")


logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(message)s"
)



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


def calculate_md5(file_path: Path) -> str:
    md5 = hashlib.md5()

    with open(file_path, "rb") as f:
        while chunk := f.read(1024 * 1024):
            md5.update(chunk)

    return md5.hexdigest()


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
    BACKUP_DIR.mkdir(exist_ok=True)

    short_sig = signature[:8]

    backup_name = (
        f"{file_path.stem}_{short_sig}.zip"
    )

    backup_path = BACKUP_DIR / backup_name

    with zipfile.ZipFile(
        backup_path,
        "w",
        compression=zipfile.ZIP_DEFLATED,
        compresslevel=9
    ) as zf:
        zf.write(
            file_path,
            arcname=file_path.name
        )

    return backup_path


def register_backup(conn, signature: str):
    conn.execute("""
        INSERT INTO backups (
            backup_time,
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


def main():
    try:
        conn = get_db()

        if not SOURCE_FILE.exists():
            logging.error(
                f"Файл не знайдено: {SOURCE_FILE}"
            )
            return

        current_signature = calculate_md5(
            SOURCE_FILE
        )

        last_signature = get_last_signature(
            conn
        )

        if current_signature == last_signature:
            logging.info(
                "Змін не виявлено"
            )
            return

        backup_path = save_backup(
            SOURCE_FILE,
            current_signature
        )

        register_backup(
            conn,
            current_signature
        )

        logging.info(
            f"Створено бекап: {backup_path}"
        )

    except Exception:
        logging.exception(
            "Помилка під час виконання"
        )


if __name__ == "__main__":
    if not SOURCE_FILE.exists():
        logging.error(
            f"Файл не знайдено: {SOURCE_FILE}"
        )
        exit(1)

    current_signature = calculate_md5(SOURCE_FILE)
    print(current_signature)

    conn = get_db()
    try:
        last_signature = get_last_signature(conn)

        if current_signature == last_signature:
            logging.info(
                "Змін не виявлено"
            )
            exit()

    except  Exception as e:
        print(str(e))
    finally:
        conn.close()



