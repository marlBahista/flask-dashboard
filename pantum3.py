import pytesseract
import pyautogui
import re
import time
from PIL import Image, ImageEnhance, ImageFilter
import tkinter as tk
from tkinter import messagebox

# Set Tesseract-OCR path (Update if needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

# Define Print Button Area (Adjust based on screen resolution)
PRINT_BUTTON_REGION = (800, 600, 200, 100)  # Example (x, y, width, height)

def is_mouse_hovering_print():
    """Check if the mouse is hovering over the print button."""
    x, y = pyautogui.position()
    px, py, pw, ph = PRINT_BUTTON_REGION
    return px <= x <= px + pw and py <= y <= py + ph

def auto_capture_on_hover():
    """Captures print dialog when the mouse hovers over the print button."""
    while True:
        if is_mouse_hovering_print():
            print("🖱️ Hover detected on Print button! Capturing dialog...")
            time.sleep(0.5)  # Small delay to avoid multiple captures
            capture_print_dialog()
            break  # Stop monitoring after capturing

def capture_print_dialog():
    """Captures a screenshot of the print dialog."""
    time.sleep(1)  # Small delay to allow focus switch
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
root.title("Auto Capture on Print Button Hover")

capture_button = tk.Button(root, text="Start Hover Detection", command=auto_capture_on_hover, padx=20, pady=10)
capture_button.pack(pady=20)

root.mainloop()
