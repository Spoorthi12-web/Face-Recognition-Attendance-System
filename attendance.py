# attendance.py — Punch In / Punch Out Attendance System

import cv2
import pickle
import datetime
import mysql.connector
from deepface import DeepFace
from config import DB_CONFIG, KNOWN_FACES_DIR, CAMERA_INDEX

# Load saved face encodings
with open("encodings.pkl", "rb") as f:
    data = pickle.load(f)

known_names = data["names"]
print("Starting Attendance System...")
print(f"Loaded faces: {known_names}")

# Current detected person
current_name = "Unknown"

# ─── Database Functions ───────────────────────────────────────────

def get_connection():
    return mysql.connector.connect(**DB_CONFIG)

def punch_in(name):
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT person_id FROM persons WHERE full_name = %s", (name,))
        result = cursor.fetchone()

        if not result:
            print(f"⚠ {name} not found in database!")
            cursor.close(); conn.close()
            return False

        person_id = result[0]

        cursor.execute("""
            SELECT log_id FROM attendance_log 
            WHERE person_id = %s AND date = CURDATE() AND punch_type = 'IN'
        """, (person_id,))

        already = cursor.fetchone()
        if already:
            print(f"⚠ {name} already punched in today!")
            cursor.close(); conn.close()
            return False

        cursor.execute("""
            INSERT INTO attendance_log (person_id, date, time_in, status, punch_type)
            VALUES (%s, CURDATE(), CURTIME(), 'PRESENT', 'IN')
        """, (person_id,))

        conn.commit()
        print(f"✅ PUNCH IN: {name} at {datetime.datetime.now().strftime('%H:%M:%S')}")
        cursor.close(); conn.close()
        return True

    except Exception as e:
        print(f"Database error: {e}")
        return False

def punch_out(name):
    try:
        conn   = get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT person_id FROM persons WHERE full_name = %s", (name,))
        result = cursor.fetchone()

        if not result:
            print(f"⚠ {name} not found in database!")
            cursor.close(); conn.close()
            return False

        person_id = result[0]

        cursor.execute("""
            SELECT log_id FROM attendance_log 
            WHERE person_id = %s AND date = CURDATE() AND punch_type = 'IN'
        """, (person_id,))

        record = cursor.fetchone()
        if not record:
            print(f"⚠ {name} has not punched in today!")
            cursor.close(); conn.close()
            return False

        cursor.execute("""
            UPDATE attendance_log 
            SET time_out = CURTIME(), punch_type = 'OUT'
            WHERE person_id = %s AND date = CURDATE()
        """, (person_id,))

        conn.commit()
        print(f"✅ PUNCH OUT: {name} at {datetime.datetime.now().strftime('%H:%M:%S')}")
        cursor.close(); conn.close()
        return True

    except Exception as e:
        print(f"Database error: {e}")
        return False

# ─── Button Click Handler ─────────────────────────────────────────

status_message = ""
status_color   = (200, 200, 200)
status_time    = 0

def handle_button(action):
    global status_message, status_color, status_time

    if current_name == "Unknown":
        status_message = "No face detected!"
        status_color   = (0, 0, 200)
        status_time    = datetime.datetime.now()
        print("⚠ No face detected — show your face first!")
        return

    if action == "IN":
        success = punch_in(current_name)
        if success:
            status_message = f"PUNCH IN: {current_name}"
            status_color   = (0, 200, 0)
        else:
            status_message = "Already punched in!"
            status_color   = (0, 140, 255)

    elif action == "OUT":
        success = punch_out(current_name)
        if success:
            status_message = f"PUNCH OUT: {current_name}"
            status_color   = (0, 0, 200)
        else:
            status_message = "Not punched in yet!"
            status_color   = (0, 0, 200)

    status_time = datetime.datetime.now()

# ─── Mouse Click Detection ────────────────────────────────────────

def mouse_click(event, x, y, flags, param):
    if event == cv2.EVENT_LBUTTONDOWN:
        if 20 <= x <= 200 and 420 <= y <= 470:
            handle_button("IN")
        if 220 <= x <= 420 and 420 <= y <= 470:
            handle_button("OUT")

# ─── Main Program ─────────────────────────────────────────────────

cap = cv2.VideoCapture(CAMERA_INDEX)
if not cap.isOpened():
    print("Error: Could not open webcam")
    exit()

cv2.namedWindow("Attendance System")
cv2.setMouseCallback("Attendance System", mouse_click)

print("Webcam opened!")
print("Show your face → then click PUNCH IN or PUNCH OUT")
print("Press Q to quit")

while True:
    ret, frame = cap.read()
    if not ret:
        break

    frame = cv2.resize(frame, (640, 480))
    small = cv2.resize(frame, (0, 0), fx=0.5, fy=0.5)

    try:
        faces = DeepFace.extract_faces(
            img_path          = small,
            detector_backend  = "opencv",
            enforce_detection = False
        )

        detected = False
        for face in faces:
            region = face["facial_area"]
            x = region["x"] * 2
            y = region["y"] * 2
            w = region["w"] * 2
            h = region["h"] * 2
            face_crop = frame[y:y+h, x:x+w]

            if face_crop.size == 0:
                continue

            try:
                result = DeepFace.find(
                    img_path          = face_crop,
                    db_path           = KNOWN_FACES_DIR,
                    model_name        = "Facenet",
                    enforce_detection = False,
                    silent            = True
                )

                if len(result) > 0 and len(result[0]) > 0:
                    identity     = result[0].iloc[0]["identity"]

                    # ── FIXED LINE ──────────────────────────────
                    current_name = identity.replace("\\", "/").split("/")[-1]
                    current_name = current_name.replace("_", " ")\
                                               .replace(".jpeg", "")\
                                               .replace(".JPEG", "")\
                                               .replace(".jpg", "")\
                                               .replace(".JPG", "")\
                                               .replace(".png", "")\
                                               .replace(".PNG", "")\
                                               .strip().title()
                    # ───────────────────────────────────────────

                    color    = (0, 200, 0)
                    detected = True
                else:
                    current_name = "Unknown"
                    color        = (0, 0, 200)

            except Exception:
                current_name = "Unknown"
                color        = (0, 0, 200)

            cv2.rectangle(frame, (x, y), (x+w, y+h), color, 2)
            cv2.rectangle(frame, (x, y+h-30), (x+w, y+h), color, cv2.FILLED)
            cv2.putText(frame, current_name, (x+6, y+h-6),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.6, (255, 255, 255), 1)

        if not detected:
            current_name = "Unknown"

    except Exception:
        current_name = "Unknown"

    # ── Punch In Button ─────────────────────────────────────────
    cv2.rectangle(frame, (20, 420), (200, 470), (0, 180, 0), -1)
    cv2.rectangle(frame, (20, 420), (200, 470), (0, 255, 0), 2)
    cv2.putText(frame, "PUNCH IN", (45, 453),
                cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)

    # ── Punch Out Button ────────────────────────────────────────
    cv2.rectangle(frame, (220, 420), (420, 470), (0, 0, 180), -1)
    cv2.rectangle(frame, (220, 420), (420, 470), (0, 0, 255), 2)
    cv2.putText(frame, "PUNCH OUT", (232, 453),
                cv2.FONT_HERSHEY_SIMPLEX, 0.8, (255, 255, 255), 2)

    # ── Person Name ─────────────────────────────────────────────
    cv2.putText(frame, f"Person: {current_name}", (20, 30),
                cv2.FONT_HERSHEY_SIMPLEX, 0.7, (200, 200, 200), 2)

    # ── Timestamp ───────────────────────────────────────────────
    now = datetime.datetime.now().strftime("%Y-%m-%d  %H:%M:%S")
    cv2.putText(frame, now, (20, 60),
                cv2.FONT_HERSHEY_SIMPLEX, 0.6, (200, 200, 200), 1)

    # ── Status Message ───────────────────────────────────────────
    if status_message and status_time:
        diff = (datetime.datetime.now() - status_time).seconds
        if diff < 3:
            cv2.rectangle(frame, (10, 380), (630, 415), (50, 50, 50), -1)
            cv2.putText(frame, status_message, (20, 405),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.7, status_color, 2)

    cv2.imshow("Attendance System", frame)

    if cv2.waitKey(1) & 0xFF == ord('q'):
        break

cap.release()
cv2.destroyAllWindows()
print("Attendance system closed!")
