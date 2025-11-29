import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import os
import sys
import matplotlib
matplotlib.use("Agg")  # Prevent backend errors
from datetime import datetime

# -----------------------------
# Subjects & Paths
# -----------------------------
SUBJECTS = ["DBMS", "Java", "Python", "OS", "CN"]  # Fixed list of subjects

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STUDENTS_FILE = os.path.join(BASE_DIR, "students.csv")

def get_csv_path_for_subject(subject):
    """
    Return CSV path for a given subject.
    Example: Attendance_DBMS.csv, Attendance_Java.csv, etc.
    """
    filename = f"Attendance_{subject}.csv"
    return os.path.join(BASE_DIR, filename)

# -----------------------------
# Receive USN from login.py
# -----------------------------
try:
    USN = sys.argv[1].upper()
except:
    messagebox.showerror("Error", "USN not received from login.")
    sys.exit()

# -----------------------------
# Fetch Student Name
# -----------------------------
def get_student_name():
    if not os.path.exists(STUDENTS_FILE):
        return USN

    df = pd.read_csv(STUDENTS_FILE)
    df["USN"] = df["USN"].str.upper()

    match = df[df["USN"] == USN]
    if not match.empty:
        return match.iloc[0]["Name"]
    return USN

STUDENT_NAME = get_student_name()

# -----------------------------
# Load Attendance for a subject
# -----------------------------
def load_attendance_for_subject(subject):
    """
    Load attendance for a specific subject from CSV:
    Attendance_<subject>.csv
    """
    csv_path = get_csv_path_for_subject(subject)
    if not os.path.exists(csv_path):
        # Return empty DataFrame with expected columns
        return pd.DataFrame(columns=["Date", "Time", "USN", "Name", "Status"])
    df = pd.read_csv(csv_path)
    if "USN" in df.columns:
        df["USN"] = df["USN"].astype(str).str.upper()
    return df

# -----------------------------
# Student Dashboard Class
# -----------------------------
class StudentDashboard:

    def __init__(self, root):
        self.root = root
        self.root.title("Student Dashboard")
        self.root.geometry("900x650")
        self.root.config(bg="white")

        # Currently selected subject
        self.selected_subject = tk.StringVar()
        self.selected_subject.set(SUBJECTS[0])  # default first subject

        # Title
        tk.Label(
            root,
            text=f"Welcome {STUDENT_NAME} ({USN}) 👋",
            font=("Arial", 22, "bold"),
            bg="#4CAF50",
            fg="white",
            pady=10
        ).pack(fill="x")

        # Subject Selection Frame
        subject_frame = tk.Frame(root, bg="white")
        subject_frame.pack(pady=10)

        tk.Label(
            subject_frame,
            text="Select Subject:",
            font=("Arial", 12, "bold"),
            bg="white",
            fg="black"
        ).grid(row=0, column=0, padx=5)

        self.subject_combo = ttk.Combobox(
            subject_frame,
            textvariable=self.selected_subject,
            values=SUBJECTS,
            state="readonly",
            font=("Arial", 11),
            width=18
        )
        self.subject_combo.grid(row=0, column=1, padx=5)
        self.subject_combo.current(0)

        # Button Frame
        btn_frame = tk.Frame(root, bg="white")
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame, text="📅 View Today's Attendance",
            font=("Arial", 14), bg="#2196F3", fg="white",
            command=self.show_today, width=25
        ).grid(row=0, column=0, pady=8, padx=5)

        tk.Button(
            btn_frame, text="📘 View Full Attendance",
            font=("Arial", 14), bg="#9C27B0", fg="white",
            command=self.show_full, width=25
        ).grid(row=1, column=0, pady=8, padx=5)

        tk.Button(
            btn_frame, text="📈 View Attendance Trend",
            font=("Arial", 14), bg="#F57C00", fg="white",
            command=self.view_trend, width=25
        ).grid(row=2, column=0, pady=8, padx=5)

        tk.Button(
            btn_frame, text="🚪 Logout",
            font=("Arial", 14), bg="#E53935", fg="white",
            command=self.logout, width=25
        ).grid(row=3, column=0, pady=8, padx=5)

        # Table Frame
        self.table_frame = tk.Frame(root)
        self.table_frame.pack(fill="both", expand=True, padx=20, pady=20)

        self.tree = ttk.Treeview(
            self.table_frame,
            columns=("Date", "Time", "Status"),
            show="headings"
        )

        for col in ("Date", "Time", "Status"):
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center", width=200)

        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=vsb.set)
        vsb.pack(side="right", fill="y")
        self.tree.pack(fill="both", expand=True)

    # -----------------------------
    # Helper: Get current subject
    # -----------------------------
    def get_current_subject(self):
        subject = self.selected_subject.get()
        if subject not in SUBJECTS:
            messagebox.showwarning("Subject", "Please select a valid subject.")
            return None
        return subject

    # -----------------------------
    # Show Today's Attendance
    # -----------------------------
    def show_today(self):
        subject = self.get_current_subject()
        if not subject:
            return

        df = load_attendance_for_subject(subject)
        today = datetime.now().strftime("%Y-%m-%d")
        today_att = df[(df["USN"] == USN) & (df["Date"] == today)]

        self.tree.delete(*self.tree.get_children())

        if today_att.empty:
            messagebox.showinfo(
                "Today",
                f"You are ABSENT today for {subject}."
            )
        else:
            for _, r in today_att.iterrows():
                self.tree.insert("", "end", values=[r["Date"], r["Time"], r["Status"]])

    # -----------------------------
    # Show Full Attendance
    # -----------------------------
    def show_full(self):
        subject = self.get_current_subject()
        if not subject:
            return

        df = load_attendance_for_subject(subject)
        my_att = df[df["USN"] == USN]

        self.tree.delete(*self.tree.get_children())

        if my_att.empty:
            messagebox.showinfo(
                "No Records",
                f"No attendance records found for {subject}."
            )
            return

        for _, r in my_att.iterrows():
            self.tree.insert("", "end", values=[r["Date"], r["Time"], r["Status"]])

    # -----------------------------
    # Attendance Trend Graph
    # -----------------------------
    def view_trend(self):
        subject = self.get_current_subject()
        if not subject:
            return

        df = load_attendance_for_subject(subject)
        my_att = df[df["USN"] == USN]

        if my_att.empty:
            messagebox.showinfo(
                "Info",
                f"No attendance records found for {subject}."
            )
            return

        try:
            from matplotlib.figure import Figure
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
        except:
            messagebox.showerror("Error", "Install matplotlib:\n pip install matplotlib")
            return

        win = tk.Toplevel(self.root)
        win.title(f"Attendance Trend - {subject}")
        win.geometry("900x600")

        fig = Figure(figsize=(8, 5))
        ax = fig.add_subplot(111)

        dates = sorted(my_att["Date"].unique())
        # Just mark presence as 1 for each date
        values = [1] * len(dates)

        ax.plot(dates, values, marker="o")
        ax.set_ylim(0, 2)
        ax.set_yticks([1])
        ax.set_yticklabels(["Present"])
        ax.set_title(f"Attendance Trend - {STUDENT_NAME} ({USN}) - {subject}")
        ax.tick_params(axis="x", rotation=45)

        canvas = FigureCanvasTkAgg(fig, win)
        canvas.draw()
        canvas.get_tk_widget().pack(fill="both", expand=True)

    # -----------------------------
    # Logout
    # -----------------------------
    def logout(self):
        self.root.destroy()
        os.system("python login.py")

# -----------------------------
# Run App
# -----------------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = StudentDashboard(root)
    root.mainloop()
