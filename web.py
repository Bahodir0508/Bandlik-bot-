from flask import Flask, render_template, request
import sqlite3
from datetime import datetime

app = Flask(__name__)

def db():
    return sqlite3.connect("ishchilar.db")

@app.route("/", methods=["GET", "POST"])
def dashboard():
    month = request.form.get("month")
    if not month:
        month = datetime.now().strftime("%Y-%m")

    conn = db()
    cur = conn.cursor()

    cur.execute("""
        SELECT i.fullname,
               COUNT(a.id) as days,
               i.daily_salary,
               COUNT(a.id) * i.daily_salary as total
        FROM ishchilar i
        LEFT JOIN attendance a
        ON i.passport = a.passport AND a.date LIKE ?
        GROUP BY i.fullname
    """, (f"{month}%",))

    data = cur.fetchall()
    return render_template("admin.html", data=data, month=month)

if __name__ == "__main__":
    app.run(debug=True)
