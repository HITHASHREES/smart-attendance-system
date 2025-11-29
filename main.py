import os
import csv
from datetime import datetime
import cv2
import face_recognition
import numpy as np
import pickle
import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
import sys

# ---------------------------------------------------------
# 📌 Read subject argument from Admin Dashboard
# ---------------------------------------------------------
# Example: python main.py auto Java
SUBJECT = "General"
if len(sys.argv) >= 3:
    SUBJECT = sys.argv[2]     # Java / DBMS / Python etc.

print(f"📘 Selected Subject → {SUBJECT}")

# ---------------------------------------------------------
# 📁 File paths
# ---------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STUDENT_IMAGES_DIR = os.path.join(BASE_DIR, "student_images")
ENCODINGS_FILE = os.path.join(BASE_DIR, "face_encodings.pkl")

# CSV file for THIS SUBJECT only
ATTENDANCE_FILE = os.path.join(BASE_DIR, f"Attendance_{SUBJECT}.csv")

os.makedirs(STUDENT_IMAGES_DIR, exist_ok=True)

print("🚀 Starting Smart Attendance System...")


# ---------------------------------------------------------
# ☁ Google Sheets Setup (SEPARATE worksheet per subject)
# ---------------------------------------------------------
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "credentials.json")
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

sheet = None
worksheet = None

try:
    creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    client = gspread.authorize(creds)

    # whole sheet group
    sheet = client.open("Attendance Report")

    # create worksheet per subject
    try:
        worksheet = sheet.worksheet(SUBJECT)
    except gspread.WorksheetNotFound:
        print(f"📝 Creating Google Sheet Tab: {SUBJECT}")
        worksheet = sheet.add_worksheet(title=SUBJECT, rows=2000, cols=10)
        worksheet.append_row(["USN", "Name", "Date", "Time", "Status"])

    print("✔ Connected to Google Sheet!")
except Exception as e:
    print("⚠ Could not connect to Google Sheets:", e)
    sheet = None


# ---------------------------------------------------------
# ☁ Mark attendance in Google sheet (per subject)
# ---------------------------------------------------------
def mark_google(usn, name):
    if worksheet is None:
        return

    today = datetime.now().strftime("%Y-%m-%d")
    time_now = datetime.now().strftime("%H:%M:%S")

    try:
        records = worksheet.get_all_records()
        for r in records:
            if str(r.get("USN")) == usn and str(r.get("Date")) == today:
                return
        worksheet.append_row([usn, name, today, time_now, "PRESENT"])
        print(f"☁ Synced → {usn} ({name}) to GoogleSheet → {SUBJECT}")
    except Exception as e:
        print("⚠ Google Sheet sync failed:", e)


# ---------------------------------------------------------
# 📝 Local CSV attendance (per subject)
# ---------------------------------------------------------
def mark_local_csv(usn, name):
    today = datetime.now().strftime("%Y-%m-%d")
    time_now = datetime.now().strftime("%H:%M:%S")

    # Create file if not exists
    if not os.path.exists(ATTENDANCE_FILE):
        with open(ATTENDANCE_FILE, "w", encoding="utf-8") as f:
            f.write("USN,Name,Date,Time,Status\n")

    df = pd.read_csv(ATTENDANCE_FILE)

    # Do not mark twice same day for same subject
    if not df[(df["USN"] == usn) & (df["Date"] == today)].empty:
        return

    with open(ATTENDANCE_FILE, "a", encoding="utf-8") as f:
        f.write(f"{usn},{name},{today},{time_now},PRESENT\n")

    print(f"📁 Saved locally → {usn} ({name}) for {SUBJECT}")


# ---------------------------------------------------------
# 📸 Face Recognition Attendance
# ---------------------------------------------------------
def take_attendance():

    if not os.path.exists(ENCODINGS_FILE):
        print("⚠ No encodings found. Register students first!")
        return

    with open(ENCODINGS_FILE, "rb") as f:
        known_encodings, known_names = pickle.load(f)

    cap = cv2.VideoCapture(0)
    print("📷 Camera started (press q to exit)")

    today = datetime.now().strftime("%Y-%m-%d")
    cooldown = {}       # to avoid spam "already marked"
    cooldown_seconds = 5

    # Load already marked students for this subject
    already_present = set()
    if os.path.exists(ATTENDANCE_FILE):
        df = pd.read_csv(ATTENDANCE_FILE)
        df_today = df[df["Date"] == today]
        already_present = set(df_today["USN"].astype(str))

    while True:
        ret, frame = cap.read()
        if not ret:
            break

        rgb = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        faces = face_recognition.face_locations(rgb)
        encodings = face_recognition.face_encodings(rgb, faces)

        for enc, loc in zip(encodings, faces):
            matches = face_recognition.compare_faces(known_encodings, enc, tolerance=0.47)
            dist = face_recognition.face_distance(known_encodings, enc)
            best = np.argmin(dist)

            if matches[best]:
                full_name = known_names[best]
                usn, student_name = full_name.split("_", 1)

                now_ts = datetime.now().timestamp()

                # already marked?
                if usn in already_present:
                    if usn not in cooldown or (now_ts - cooldown[usn]) > cooldown_seconds:
                        print(f"🙂 Already marked today → {usn} ({student_name}) [{SUBJECT}]")
                        cooldown[usn] = now_ts
                else:
                    # NEW attendance
                    mark_local_csv(usn, student_name)
                    mark_google(usn, student_name)
                    already_present.add(usn)
                    print(f"✔ Marked PRESENT → {usn} ({student_name}) [{SUBJECT}]")

                # Draw green box
                top, right, bottom, left = loc
                cv2.rectangle(frame, (left, top), (right, bottom), (0, 255, 0), 2)
                cv2.putText(frame, student_name, (left, top - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)

            else:
                # unknown person
                top, right, bottom, left = loc
                cv2.rectangle(frame, (left, top), (right, bottom), (0, 0, 255), 2)
                cv2.putText(frame, "Unknown", (left, top - 10),
                            cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 0, 255), 2)

        cv2.imshow("Face Recognition - Smart Attendance", frame)
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    cap.release()
    cv2.destroyAllWindows()


# ---------------------------------------------------------
# 🚪 Entry Point
# ---------------------------------------------------------
if __name__ == "__main__":
    if len(sys.argv) >= 2 and sys.argv[1] == "auto":
        print(f"🎯 Attendance mode started → SUBJECT = {SUBJECT}")
        take_attendance()
    else:
        print("Usage: python main.py auto <SUBJECT>")
