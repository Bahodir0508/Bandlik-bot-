import sqlite3

conn = sqlite3.connect("ishchilar.db")
cur = conn.cursor()

try:
    cur.execute("ALTER TABLE ishchilar ADD COLUMN tg_id TEXT DEFAULT NULL")
    print("✅ tg_id ustuni qo‘shildi!")
except Exception as e:
    print("⚠️ Ehtimol ustun allaqachon mavjud:", e)

conn.commit()
conn.close()
