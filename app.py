# app.py — Flask Web App for Attendance Dashboard

from flask import Flask, render_template, jsonify, request
import mysql.connector
import datetime
import os
from config import DB_CONFIG, REPORTS_DIR

app = Flask(__name__)

def get_connection():
    return mysql.connector.connect(**DB_CONFIG)

# ── Helper: Get today's attendance ───────────────────────────
def get_today_attendance():
    conn   = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("""
        SELECT p.full_name, p.department, p.email,
               a.time_in, a.time_out, a.status, a.date
        FROM attendance_log a
        JOIN persons p ON p.person_id = a.person_id
        WHERE a.date = CURDATE()
        ORDER BY a.time_in
    """)
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    for r in rows:
        r["time_in"]  = str(r["time_in"])  if r["time_in"]  else "—"
        r["time_out"] = str(r["time_out"]) if r["time_out"] else "—"
        r["date"]     = str(r["date"])
    return rows

# ── Helper: Get all persons ───────────────────────────────────
def get_all_persons():
    conn   = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT * FROM persons ORDER BY full_name")
    rows = cursor.fetchall()
    cursor.close()
    conn.close()
    for r in rows:
        r["created_on"] = str(r["created_on"])
    return rows

# ── Helper: Get summary stats ─────────────────────────────────
def get_summary():
    conn   = get_connection()
    cursor = conn.cursor(dictionary=True)
    cursor.execute("SELECT COUNT(*) as total FROM persons")
    total = cursor.fetchone()["total"]
    cursor.execute("""
        SELECT COUNT(*) as present FROM attendance_log
        WHERE date = CURDATE()
    """)
    present = cursor.fetchone()["present"]
    cursor.execute("""
        SELECT a.date, COUNT(*) as count
        FROM attendance_log a
        WHERE a.date >= DATE_SUB(CURDATE(), INTERVAL 7 DAY)
        GROUP BY a.date ORDER BY a.date
    """)
    weekly = cursor.fetchall()
    cursor.close()
    conn.close()
    absent = total - present
    pct    = round((present / total * 100), 1) if total > 0 else 0
    for w in weekly:
        w["date"] = str(w["date"])
    return {
        "total"  : total,
        "present": present,
        "absent" : absent,
        "pct"    : pct,
        "weekly" : weekly
    }

# ── Routes ────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/api/summary")
def api_summary():
    return jsonify(get_summary())

@app.route("/api/attendance")
def api_attendance():
    return jsonify(get_today_attendance())

@app.route("/api/persons")
def api_persons():
    return jsonify(get_all_persons())

@app.route("/api/report")
def api_report():
    from excel_report import generate_excel_report
    path = generate_excel_report()
    return jsonify({"status": "success", "file": path})

if __name__ == "__main__":
    app.run(debug=True, port=5000)