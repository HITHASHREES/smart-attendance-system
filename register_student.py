import os
import cv2
import face_recognition
import numpy as np
import pandas as pd
import pickle
import threading
import tkinter as tk
from tkinter import messagebox

# -----------------------------
# Paths
# -----------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
STUDENT_IMAGES_DIR = os.path.join(BASE_DIR, "student_images")
ENCODINGS_FILE = os.path.join(BASE_DIR, "face_encodings.pkl")
STUDENTS_FILE = os.path.join(BASE_DIR, "students.csv")

os.makedirs(STUDENT_IMAGES_DIR, exist_ok=True)


# -----------------------------
# 🧠 Build encodings (used in thread)
# -----------------------------
def generate_face_encodings_thread(callback):
    """
    Heavy encoding function - runs in background thread.
    Calls callback(success_flag) after completion.
    """
    try:
        known_encodings = []
        known_names = []

        for folder in os.listdir(STUDENT_IMAGES_DIR):
            path = os.path.join(STUDENT_IMAGES_DIR, folder)
            if not os.path.isdir(path):
                continue

            try:
                usn, name = folder.split("_", 1)
                full_name = f"{usn}_{name}"
            except:
                print("Skipping wrongly formatted folder:", folder)
                continue

            for file in os.listdir(path):
                if not file.lower().endswith((".jpg", ".jpeg", ".png")):
                    continue

                img_path = os.path.join(path, file)
                img = cv2.imread(img_path)
                if img is None:
                    continue

                rgb = cv2.cvtColor(img, cv2.COLOR_BGR2RGB)
                faces = face_recognition.face_locations(rgb)
                enc = face_recognition.face_encodings(rgb, faces)

                if enc:
                    known_encodings.append(enc[0])
                    known_names.append(full_name)

        if len(known_encodings) == 0:
            callback(False)
            return

        with open(ENCODINGS_FILE, "wb") as f:
            pickle.dump((known_encodings, known_names), f)

        callback(True)

    except Exception as e:
        print("Encoding error:", e)
        callback(False)


# -----------------------------
# Ensure student in CSV
# -----------------------------
def ensure_student_in_csv(usn, name):
    if not os.path.exists(STUDENTS_FILE):
        pd.DataFrame(columns=["USN", "Name"]).to_csv(STUDENTS_FILE, index=False)

    df = pd.read_csv(STUDENTS_FILE)
    usn_up = usn.upper()

    if not (df["USN"].astype(str).str.upper() == usn_up).any():
        df.loc[len(df)] = [usn_up, name]
        df.to_csv(STUDENTS_FILE, index=False)


# -----------------------------
# Registration Window
# -----------------------------
class StudentRegistrationWindow:
    def __init__(self, parent=None):
        self.parent = parent
        self.window = tk.Toplevel(parent) if parent else tk.Tk()
        self.window.title("➕ Register New Student")
        self.window.geometry("420x270")
        self.window.resizable(False, False)

        tk.Label(self.window, text="Register New Student",
                 font=("Helvetica", 16, "bold")).pack(pady=10)

        form = tk.Frame(self.window)
        form.pack(pady=10)

        tk.Label(form, text="USN:", font=("Helvetica", 12)).grid(row=0, column=0, sticky="e", padx=5, pady=5)
        tk.Label(form, text="Name:", font=("Helvetica", 12)).grid(row=1, column=0, sticky="e", padx=5, pady=5)

        self.usn_entry = tk.Entry(form, font=("Helvetica", 12), width=26)
        self.name_entry = tk.Entry(form, font=("Helvetica", 12), width=26)

        self.usn_entry.grid(row=0, column=1)
        self.name_entry.grid(row=1, column=1)

        btns = tk.Frame(self.window)
        btns.pack(pady=15)

        tk.Button(btns, text="📸 Capture Photos",
                  font=("Helvetica", 11),
                  command=self.capture_photos).grid(row=0, column=0, padx=10)

        tk.Button(btns, text="✨ Generate Encodings & Save",
                  font=("Helvetica", 11),
                  command=self.generate_and_save).grid(row=0, column=1, padx=10)

    # -----------------------------
    # Capture Photos
    # -----------------------------
    def capture_photos(self):
        usn = self.usn_entry.get().strip().upper()
        name = self.name_entry.get().strip()

        if not usn or not name:
            messagebox.showerror("Error", "Enter both USN and Name.")
            return

        save_dir = os.path.join(STUDENT_IMAGES_DIR, f"{usn}_{name}")
        os.makedirs(save_dir, exist_ok=True)

        cap = cv2.VideoCapture(0)
        count = 0
        target = 20

        messagebox.showinfo("Instructions",
                            "Press 'c' to capture image\nPress 'q' to stop early")

        while True:
            ok, frame = cap.read()
            if not ok:
                break

            cv2.putText(frame, f"Captured {count}/{target}", (10, 30),
                        cv2.FONT_HERSHEY_SIMPLEX, 0.8, (0, 255, 0), 2)

            cv2.imshow("Photo Capture", frame)
            key = cv2.waitKey(1) & 0xFF

            if key == ord('c'):
                cv2.imwrite(os.path.join(save_dir, f"{count}.jpg"), frame)
                count += 1
                if count >= target:
                    break
            elif key == ord('q'):
                break

        cap.release()
        cv2.destroyAllWindows()

        if count > 0:
            messagebox.showinfo("Done", f"Captured {count} photos.")
        else:
            messagebox.showinfo("Info", "No photos captured.")

    # -----------------------------
    # Generate Encodings (Threaded)
    # -----------------------------
    def generate_and_save(self):
        usn = self.usn_entry.get().strip().upper()
        name = self.name_entry.get().strip()

        if not usn or not name:
            messagebox.showerror("Error", "Enter both USN and Name.")
            return

        folder = os.path.join(STUDENT_IMAGES_DIR, f"{usn}_{name}")
        if not os.path.exists(folder):
            messagebox.showerror("Error", "No photos found.\nCapture photos first.")
            return

        # Loading popup
        loading = tk.Toplevel(self.window)
        loading.title("Processing...")
        loading.geometry("250x100")
        tk.Label(loading, text="Generating Encodings...\nPlease wait...",
                 font=("Helvetica", 12)).pack(pady=20)
        loading.grab_set()

        def callback(success):
            loading.destroy()
            if success:
                ensure_student_in_csv(usn, name)
                messagebox.showinfo("Success", "Student registered successfully!")
                self.window.destroy()
            else:
                messagebox.showerror("Error", "Failed to generate encodings.\nCheck images and retry.")

        # Run encoding in background thread
        threading.Thread(target=generate_face_encodings_thread, args=(callback,), daemon=True).start()


def open_registration_window(parent=None):
    StudentRegistrationWindow(parent=parent)


if __name__ == "__main__":
    root = tk.Tk()
    StudentRegistrationWindow(root)
    root.mainloop()
