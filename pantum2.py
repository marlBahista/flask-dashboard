import pytesseract
import pyautogui
import re
import time
from PIL import Image, ImageEnhance, ImageFilter
import tkinter as tk
from tkinter import messagebox

# Set Tesseract-OCR path (Update if needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def capture_print_dialog():
    """Captures the screenshot of the print dialog when the user clicks the button."""
    time.sleep(2)  # Allow time for user to switch to the print dialog
    screenshot_path = "print_screenshot.png"
    pyautogui.screenshot(screenshot_path)
    messagebox.showinfo("Screenshot Captured", "Print dialog screenshot saved successfully!")
    process_screenshot(screenshot_path)

def preprocess_image(image_path):
    """Enhance image for better OCR detection."""
    try:
        img = Image.open(image_path)
        img = img.convert("L")  # Grayscale
        img = ImageEnhance.Contrast(img).enhance(2.0)  # Increase contrast
        img = img.filter(ImageFilter.SHARPEN)  # Sharpen image
        return img
    except FileNotFoundError:
        print(f"❌ Error: Screenshot '{image_path}' not found.")
        return None

def extract_copies_from_text(text):
    """Extracts the number of copies using regex."""
    match = re.search(r"Copies[:=|]?\s*\[?\s*(\d+)", text, re.IGNORECASE)
    if match:
        try:
            copies = int(match.group(1))
            print(f"✅ Extracted Copies: {copies}")
            messagebox.showinfo("Detected Copies", f"Number of copies detected: {copies}")
            return copies
        except ValueError:
            print("❌ Error: Could not convert extracted copies to an integer.")

    print("⚠️ No copies detected. Defaulting to 1.")
    return 1

def process_screenshot(image_path):
    """Detects the number of copies from a saved screenshot using OCR."""
    img = preprocess_image(image_path)
    if img is None:
        return

    print(f"📸 Processing image: {image_path}...")
    extracted_text = pytesseract.image_to_string(img, config="--psm 11").strip()
    print("🔍 Extracted Text:\n", extracted_text)
    extract_copies_from_text(extracted_text)

# Create the Tkinter UI
root = tk.Tk()
root.title("Print Dialog Screenshot")

capture_button = tk.Button(root, text="Capture Print Dialog", command=capture_print_dialog, padx=20, pady=10)
capture_button.pack(pady=20)

root.mainloop()
