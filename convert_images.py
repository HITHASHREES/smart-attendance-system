import os
from PIL import Image

# 📁 Path where your student images are stored
base_dir = r"C:\Users\inbox\OneDrive\Desktop\SmartAttendanceSystem\student_images"

# Loop through all images and convert them properly
for root, dirs, files in os.walk(base_dir):
    for file in files:
        if file.lower().endswith((".jpg", ".jpeg", ".png")):
            path = os.path.join(root, file)
            try:
                img = Image.open(path).convert("RGB")   # force 8-bit RGB
                img.save(path, format="JPEG", quality=95)
                print(f"✅ Converted: {path}")
            except Exception as e:
                print(f"❌ Failed: {path} — {e}")

print("🎯 All image conversions completed successfully!")
