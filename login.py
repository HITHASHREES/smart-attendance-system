import tkinter as tk
from tkinter import messagebox
import gspread
from google.oauth2.service_account import Credentials
import subprocess
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# Google Sheets Setup
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

try:
    creds = Credentials.from_service_account_file(
        os.path.join(BASE_DIR, "service_account.json"),
        scopes=SCOPES
    )
    client = gspread.authorize(creds)
except:
    client = None

SHEET_NAME = "Attendance Report"


# STUDENT LOGIN
def login_user():
    usn = usn_entry.get().strip()

    try:
        sheet = client.open(SHEET_NAME).sheet1
        usn_list = sheet.col_values(1)

        if usn in usn_list:
            messagebox.showinfo("Success", f"Welcome {usn}!")
            root.destroy()
            subprocess.Popen([
                os.path.join(BASE_DIR, "venv", "Scripts", "python.exe"),
                os.path.join(BASE_DIR, "dashboard.py"),
                usn
            ])
        else:
            messagebox.showerror("Invalid", "USN not found")

    except Exception as e:
        messagebox.showerror("Error", f"{e}")


# ADMIN LOGIN
def login_admin():
    pwd = admin_pass_entry.get().strip()

    if pwd == "admin123":
        messagebox.showinfo("Admin", "Welcome Admin!")
        root.destroy()
        subprocess.Popen([
            os.path.join(BASE_DIR, "venv", "Scripts", "python.exe"),
            os.path.join(BASE_DIR, "admin_dashboard.py")
        ])
    else:
        messagebox.showerror("Invalid", "Wrong password")


# TKINTER UI
root = tk.Tk()
root.title("Login Page")
root.geometry("450x450")  # ⭐ INCREASED HEIGHT

title_lbl = tk.Label(root, text="Attendance Login", font=("Arial", 20, "bold"))
title_lbl.pack(pady=10)

# USER SECTION
user_frame = tk.LabelFrame(root, text="USER LOGIN", padx=20, pady=20)
user_frame.pack(fill="x", padx=20, pady=10)

tk.Label(user_frame, text="Enter USN:", font=("Arial", 12)).pack()
usn_entry = tk.Entry(user_frame, font=("Arial", 12), width=30)
usn_entry.pack(pady=5)
usn_entry.focus_set()

tk.Button(
    user_frame, text="Login as User",
    font=("Arial", 12, "bold"), bg="#4CAF50", fg="white",
    command=login_user
).pack(pady=8)

# ADMIN SECTION
admin_frame = tk.LabelFrame(root, text="ADMIN LOGIN", padx=20, pady=20)
admin_frame.pack(fill="x", padx=20, pady=10)

tk.Label(admin_frame, text="Admin Password:", font=("Arial", 12)).pack()
admin_pass_entry = tk.Entry(admin_frame, font=("Arial", 12), width=30, show="*")
admin_pass_entry.pack(pady=5)

# ⭐ BUTTON WAS HIDDEN BEFORE — NOW FIXED ⭐
admin_login_btn = tk.Button(
    admin_frame, text="Login as Admin",
    font=("Arial", 12, "bold"), bg="#2196F3", fg="white",
    command=login_admin
)
admin_login_btn.pack(pady=8)  # ⭐ Visible now

root.mainloop()
