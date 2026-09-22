import sqlite3

conn = sqlite3.connect("../wasted.db")

def native_clear_table_rows(tables: list[str]):
    cursor = conn.cursor()
    for table in tables:
        cursor.execute(f"DELETE FROM {table}")
    conn.commit()
    conn.close()

if __name__ == "__main__":
    native_clear_table_rows(['orders_wasteaccountbook','titles_wastebook'])

