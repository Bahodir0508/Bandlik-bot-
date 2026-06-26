import sqlite3

conn = sqlite3.connect("ishchilar.db")
cur = conn.cursor()

# Ustunlar mavjudligini tekshiramiz va kerak bo‘lsa qo‘shamiz
try:
    cur.execute("ALTER TABLE messages ADD COLUMN sender TEXT DEFAULT 'worker'")
    print("✅ sender ustuni qo‘shildi!")
except Exception as e:
    print("⚠️ sender ustuni allaqachon mavjud yoki xato:", e)

try:
    cur.execute("ALTER TABLE messages ADD COLUMN type TEXT DEFAULT 'text'")
    print("✅ type ustuni qo‘shildi!")
except Exception as e:
    print("⚠️ type ustuni allaqachon mavjud yoki xato:", e)

conn.commit()
conn.close()
print("✅ messages jadvali yangilandi!")
