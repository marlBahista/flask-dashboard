import pyautogui
import time
import pytesseract
from PIL import ImageGrab

# Set the path to Tesseract-OCR (Update this based on your installation)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def get_copies_from_print_window():
    """Detects the number of copies from the Print dialog."""
    
    print("📸 Capturing print window...")

    # Give time for user to open the Print dialog
    time.sleep(10)

    # Define region (adjust these coordinates based on your screen resolution)
    copies_region = (800, 500, 900, 550)  # Adjust for your system
    
    # Capture the region where "Copies" field appears
    screenshot = ImageGrab.grab(bbox=copies_region)

    # Convert Image to String using OCR
    copies_text = pytesseract.image_to_string(screenshot, config="--psm 6").strip()

    # Extract numbers (default to 1 copy if detection fails)
    copies = int(copies_text) if copies_text.isdigit() else 1

    print(f"🔍 Detected Copies: {copies}")
    return copies

# Example Usage:
print("🚀 Open the Print Dialog (Ctrl + P) within 2 seconds...")
copies_detected = get_copies_from_print_window()
print(f"✅ Number of Copies Detected: {copies_detected}")
