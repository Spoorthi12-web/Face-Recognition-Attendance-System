# email_sender.py — Daily and Weekly Attendance Email Reports

import smtplib
import datetime
import mysql.connector
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
from config import DB_CONFIG, EMAIL_SENDER, EMAIL_PASSWORD, EMAIL_RECEIVER, REPORTS_DIR

def get_connection():
    return mysql.connector.connect(**DB_CONFIG)

# ── Get today's attendance data ───────────────────────────────
def get_today_data():
    conn   = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT p.full_name, p.department,
               a.time_in, a.time_out, a.status
        FROM attendance_log a
        JOIN persons p ON p.person_id = a.person_id
        WHERE a.date = CURDATE()
        ORDER BY a.time_in
    """)
    present = cursor.fetchall()

    cursor.execute("""
        SELECT full_name, department
        FROM persons
        WHERE person_id NOT IN (
            SELECT person_id FROM attendance_log
            WHERE date = CURDATE()
        )
    """)
    absent = cursor.fetchall()

    cursor.execute("SELECT COUNT(*) as total FROM persons")
    total = cursor.fetchone()["total"]

    cursor.close()
    conn.close()

    for r in present:
        r["time_in"]  = str(r["time_in"])  if r["time_in"]  else "—"
        r["time_out"] = str(r["time_out"]) if r["time_out"] else "—"

    return present, absent, total

# ── Get weekly data ───────────────────────────────────────────
def get_weekly_data():
    conn   = get_connection()
    cursor = conn.cursor(dictionary=True)

    cursor.execute("""
        SELECT a.date,
               COUNT(*) as present_count,
               (SELECT COUNT(*) FROM persons) - COUNT(*) as absent_count
        FROM attendance_log a
        WHERE a.date >= DATE_SUB(CURDATE(), INTERVAL 7 DAY)
        GROUP BY a.date
        ORDER BY a.date
    """)
    rows = cursor.fetchall()
    cursor.close()
    conn.close()

    for r in rows:
        r["date"] = str(r["date"])
    return rows

# ── Build Daily HTML Email ────────────────────────────────────
def build_daily_email(present, absent, total):
    today         = datetime.date.today().strftime("%d %B %Y")
    present_count = len(present)
    absent_count  = len(absent)
    pct           = round((present_count / total * 100), 1) if total > 0 else 0

    present_rows = ""
    for i, r in enumerate(present, start=1):
        bg = "#F1F8E9" if i % 2 == 0 else "#FFFFFF"
        present_rows += f"""
        <tr style="background:{bg}">
            <td style="padding:10px 14px;border-bottom:1px solid #E8F5E9">{i}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8F5E9;font-weight:500">{r['full_name']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8F5E9;color:#777">{r['department']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8F5E9;font-family:monospace">{r['time_in']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8F5E9;font-family:monospace">{r['time_out']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8F5E9">
                <span style="background:#E8F5E9;color:#2E7D32;padding:3px 10px;border-radius:10px;font-size:12px;font-weight:500">PRESENT</span>
            </td>
        </tr>"""

    absent_rows = ""
    for i, r in enumerate(absent, start=1):
        bg = "#FFF8F8" if i % 2 == 0 else "#FFFFFF"
        absent_rows += f"""
        <tr style="background:{bg}">
            <td style="padding:10px 14px;border-bottom:1px solid #FFEBEE">{i}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #FFEBEE;font-weight:500">{r['full_name']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #FFEBEE;color:#777">{r['department']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #FFEBEE">
                <span style="background:#FFEBEE;color:#C62828;padding:3px 10px;border-radius:10px;font-size:12px;font-weight:500">ABSENT</span>
            </td>
        </tr>"""

    if not absent_rows:
        absent_rows = """
        <tr>
            <td colspan="4" style="padding:20px;text-align:center;color:#2E7D32;font-weight:500">
                Full attendance today!
            </td>
        </tr>"""

    html = f"""
    <html>
    <body style="font-family:'Segoe UI',Arial,sans-serif;background:#F5F7FA;margin:0;padding:20px;">

    <div style="max-width:700px;margin:0 auto;">

        <!-- Header -->
        <div style="background:#0D47A1;border-radius:12px 12px 0 0;padding:28px 30px;">
            <h1 style="color:#fff;margin:0;font-size:22px;font-weight:500">Daily Attendance Report</h1>
            <p style="color:#90CAF9;margin:6px 0 0;font-size:14px">{today}</p>
        </div>

        <!-- Metric Cards -->
        <div style="background:#fff;padding:24px 30px;display:flex;gap:16px;border-bottom:1px solid #F0F0F0;">
            <div style="flex:1;background:#E3F2FD;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#0D47A1">{total}</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Total</div>
            </div>
            <div style="flex:1;background:#E8F5E9;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#2E7D32">{present_count}</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Present</div>
            </div>
            <div style="flex:1;background:#FFEBEE;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#C62828">{absent_count}</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Absent</div>
            </div>
            <div style="flex:1;background:#FFF3E0;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#E65100">{pct}%</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Attendance</div>
            </div>
        </div>

        <!-- Present Table -->
        <div style="background:#fff;padding:24px 30px;border-bottom:1px solid #F0F0F0;">
            <h2 style="font-size:15px;font-weight:500;color:#2E7D32;margin:0 0 16px">
                Present Employees ({present_count})
            </h2>
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
                <thead>
                    <tr style="background:#2E7D32;">
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">#</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Name</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Department</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Punch In</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Punch Out</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Status</th>
                    </tr>
                </thead>
                <tbody>{present_rows}</tbody>
            </table>
        </div>

        <!-- Absent Table -->
        <div style="background:#fff;padding:24px 30px;border-bottom:1px solid #F0F0F0;">
            <h2 style="font-size:15px;font-weight:500;color:#C62828;margin:0 0 16px">
                Absent Employees ({absent_count})
            </h2>
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
                <thead>
                    <tr style="background:#C62828;">
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">#</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Name</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Department</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Status</th>
                    </tr>
                </thead>
                <tbody>{absent_rows}</tbody>
            </table>
        </div>

        <!-- Footer -->
        <div style="background:#F5F7FA;border-radius:0 0 12px 12px;padding:16px 30px;text-align:center;">
            <p style="color:#aaa;font-size:12px;margin:0">
                This is an automated email from the Face Recognition Attendance System.
                Generated on {datetime.datetime.now().strftime("%d-%m-%Y at %H:%M:%S")}
            </p>
        </div>

    </div>
    </body>
    </html>
    """
    return html

# ── Build Weekly HTML Email ───────────────────────────────────
def build_weekly_email(weekly_data):
    today      = datetime.date.today().strftime("%d %B %Y")
    week_start = (datetime.date.today() - datetime.timedelta(days=7)).strftime("%d %B %Y")

    total_present = sum(r["present_count"] for r in weekly_data)
    total_absent  = sum(r["absent_count"]  for r in weekly_data)

    weekly_rows = ""
    for i, r in enumerate(weekly_data, start=1):
        bg    = "#F3F8FF" if i % 2 == 0 else "#FFFFFF"
        d     = datetime.datetime.strptime(r["date"], "%Y-%m-%d")
        day   = d.strftime("%A")
        date  = d.strftime("%d %b %Y")
        pct   = round(r["present_count"] / (r["present_count"] + r["absent_count"]) * 100, 1) if (r["present_count"] + r["absent_count"]) > 0 else 0
        bar_w = int(pct)
        weekly_rows += f"""
        <tr style="background:{bg}">
            <td style="padding:10px 14px;border-bottom:1px solid #E8EEF8;font-weight:500">{day}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8EEF8;color:#777">{date}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8EEF8;color:#2E7D32;font-weight:500">{r['present_count']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8EEF8;color:#C62828">{r['absent_count']}</td>
            <td style="padding:10px 14px;border-bottom:1px solid #E8EEF8">
                <div style="background:#E3F2FD;border-radius:4px;height:8px;width:100%">
                    <div style="background:#1565C0;border-radius:4px;height:8px;width:{bar_w}%"></div>
                </div>
                <span style="font-size:11px;color:#777">{pct}%</span>
            </td>
        </tr>"""

    html = f"""
    <html>
    <body style="font-family:'Segoe UI',Arial,sans-serif;background:#F5F7FA;margin:0;padding:20px;">

    <div style="max-width:700px;margin:0 auto;">

        <!-- Header -->
        <div style="background:#1A237E;border-radius:12px 12px 0 0;padding:28px 30px;">
            <h1 style="color:#fff;margin:0;font-size:22px;font-weight:500">Weekly Attendance Summary</h1>
            <p style="color:#9FA8DA;margin:6px 0 0;font-size:14px">{week_start} — {today}</p>
        </div>

        <!-- Summary Cards -->
        <div style="background:#fff;padding:24px 30px;display:flex;gap:16px;border-bottom:1px solid #F0F0F0;">
            <div style="flex:1;background:#E8EAF6;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#1A237E">{len(weekly_data)}</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Working Days</div>
            </div>
            <div style="flex:1;background:#E8F5E9;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#2E7D32">{total_present}</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Total Present</div>
            </div>
            <div style="flex:1;background:#FFEBEE;border-radius:10px;padding:16px;text-align:center;">
                <div style="font-size:28px;font-weight:600;color:#C62828">{total_absent}</div>
                <div style="font-size:12px;color:#777;margin-top:4px">Total Absent</div>
            </div>
        </div>

        <!-- Weekly Table -->
        <div style="background:#fff;padding:24px 30px;border-bottom:1px solid #F0F0F0;">
            <h2 style="font-size:15px;font-weight:500;color:#1A237E;margin:0 0 16px">
                Day-wise Breakdown
            </h2>
            <table style="width:100%;border-collapse:collapse;font-size:13px;">
                <thead>
                    <tr style="background:#1A237E;">
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Day</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Date</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Present</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Absent</th>
                        <th style="padding:10px 14px;color:#fff;text-align:left;font-weight:500">Attendance %</th>
                    </tr>
                </thead>
                <tbody>{weekly_rows}</tbody>
            </table>
        </div>

        <!-- Footer -->
        <div style="background:#F5F7FA;border-radius:0 0 12px 12px;padding:16px 30px;text-align:center;">
            <p style="color:#aaa;font-size:12px;margin:0">
                This is an automated weekly summary from the Face Recognition Attendance System.
                Generated on {datetime.datetime.now().strftime("%d-%m-%Y at %H:%M:%S")}
            </p>
        </div>

    </div>
    </body>
    </html>
    """
    return html

# ── Send Email Function ───────────────────────────────────────
def send_email(subject, html_body, attach_excel=False):
    try:
        msg = MIMEMultipart("alternative")
        msg["From"]    = EMAIL_SENDER
        msg["To"]      = EMAIL_RECEIVER
        msg["Subject"] = subject

        msg.attach(MIMEText(html_body, "html"))

        if attach_excel:
            from excel_report import generate_excel_report
            filepath = generate_excel_report()
            if os.path.exists(filepath):
                with open(filepath, "rb") as f:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition",
                                    f"attachment; filename={os.path.basename(filepath)}")
                    msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(EMAIL_SENDER, EMAIL_PASSWORD)
            server.sendmail(EMAIL_SENDER, EMAIL_RECEIVER, msg.as_string())

        print(f"Email sent successfully to {EMAIL_RECEIVER}")
        return True

    except Exception as e:
        print(f"Email error: {e}")
        return False

# ── Send Daily Report ─────────────────────────────────────────
def send_daily_report():
    print("Sending daily attendance report...")
    present, absent, total = get_today_data()
    html    = build_daily_email(present, absent, total)
    today   = datetime.date.today().strftime("%d %B %Y")
    subject = f"Daily Attendance Report — {today}"
    success = send_email(subject, html, attach_excel=True)
    if success:
        print("Daily report sent with Excel attachment!")
    return success

# ── Send Weekly Report ────────────────────────────────────────
def send_weekly_report():
    print("Sending weekly attendance summary...")
    weekly_data = get_weekly_data()
    html        = build_weekly_email(weekly_data)
    today       = datetime.date.today().strftime("%d %B %Y")
    subject     = f"Weekly Attendance Summary — {today}"
    success     = send_email(subject, html)
    if success:
        print("Weekly summary sent!")
    return success

# ── Auto Scheduler ────────────────────────────────────────────
def run_scheduler():
    import time
    print("Email scheduler started!")
    print("Daily report will be sent at 6:00 PM")
    print("Weekly report will be sent every Monday at 8:00 AM")
    print("Press Ctrl+C to stop")

    while True:
        now     = datetime.datetime.now()
        weekday = now.weekday()
        hour    = now.hour
        minute  = now.minute

        # Daily report at 6:00 PM every day
        if hour == 18 and minute == 0:
            send_daily_report()
            time.sleep(61)

        # Weekly report every Monday at 8:00 AM
        elif weekday == 0 and hour == 8 and minute == 0:
            send_weekly_report()
            time.sleep(61)

        time.sleep(30)

# ── Run directly to test ──────────────────────────────────────
if __name__ == "__main__":
    print("Choose what to send:")
    print("1 — Send daily report now (test)")
    print("2 — Send weekly report now (test)")
    print("3 — Start auto scheduler")
    choice = input("Enter 1, 2 or 3: ")

    if choice == "1":
        send_daily_report()
    elif choice == "2":
        send_weekly_report()
    elif choice == "3":
        run_scheduler()
