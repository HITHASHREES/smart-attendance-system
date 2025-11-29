import pandas as pd
import os
import gspread
from google.oauth2.service_account import Credentials

# -----------------------------
# 📁 File paths
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ATTENDANCE_FILE = os.path.join(BASE_DIR, "attendance.csv")
CLEANED_FILE = os.path.join(BASE_DIR, "attendance_cleaned.csv")
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "service_account.json")  # 🔑 your credentials file

print("🧹 Cleaning attendance data...")

# -----------------------------
# 🧾 Read and clean attendance data
# -----------------------------
df = pd.read_csv(ATTENDANCE_FILE)
df.drop_duplicates(subset=["USN", "Date"], keep="last", inplace=True)
df["Status"] = df["Status"].fillna("Present")
df.to_csv(CLEANED_FILE, index=False)

print("✅ Attendance cleaned successfully!")
print(f"Cleaned file saved as: {CLEANED_FILE}")

# -----------------------------
# ☁️ Upload to Google Sheets
# -----------------------------
print("📤 Uploading to Google Sheets...")

# Define Google Sheets scope and credentials
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scope)
client = gspread.authorize(creds)

# Open your Google Sheet (replace with your sheet name)
sheet = client.open("SmartAttendanceSheet").sheet1  # 📄 change sheet name here if needed

# Clear old data
sheet.clear()

# Write new header + data
sheet.update([df.columns.values.tolist()] + df.values.tolist())

print("✅ Google Sheet updated successfully!")
