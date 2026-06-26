import sqlite3

conn = sqlite3.connect("ishchilar.db")
cur = conn.cursor()

# messages jadvalini yaratish (agar mavjud bo'lmasa)
cur.execute("""
CREATE TABLE IF NOT EXISTS messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ishchi_id INTEGER,
    passport TEXT
)
""")

# Zarur ustunlar ro'yxati va ularning DEFAULT qiymatlari
columns = {
    "sender": "'worker'",
    "type": "'text'",
    "content": "TEXT",
    "date": "TEXT",
    "is_read": "INTEGER DEFAULT 0"
}

for col, default in columns.items():
    try:
        cur.execute(f"ALTER TABLE messages ADD COLUMN {col} {default}")
        print(f"✅ {col} ustuni qo‘shildi!")
    except sqlite3.OperationalError:
        print(f"⚠️ {col} ustuni allaqachon mavjud.")

conn.commit()
conn.close()
print("✅ messages jadvali to‘liq tekshirildi va yangilandi!")
