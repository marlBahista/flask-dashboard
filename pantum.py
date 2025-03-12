import pytesseract
from PIL import Image, ImageEnhance, ImageFilter
import re

# Set Tesseract-OCR path (Update if needed)
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

def preprocess_image(image_path):
    """Enhance image for better OCR detection."""
    try:
        img = Image.open(image_path)

        # Convert to grayscale and enhance contrast
        img = img.convert("L")  # Grayscale
        img = ImageEnhance.Contrast(img).enhance(2.0)  # Increase contrast
        img = img.filter(ImageFilter.SHARPEN)  # Sharpen image
        
        return img
    except FileNotFoundError:
        print(f"❌ Error: Screenshot '{image_path}' not found.")
        return None

def extract_copies_from_text(text):
    """Extracts the number of copies using a refined regex pattern that captures multi-digit numbers."""
    # Improved regex to handle multiple spaces, brackets, and symbols
    match = re.search(r"Copies[:=|]?\s*\[?\s*(\d+)", text, re.IGNORECASE)

    if match:
        try:
            copies = int(match.group(1))
            print(f"✅ Extracted Copies: {copies}")  # Debugging
            return copies
        except ValueError:
            print("❌ Error: Could not convert extracted copies to an integer.")

    print("⚠️ No copies detected. Defaulting to 1.")
    return 1  # Default to 1 if not found


def get_copies_from_screenshot(image_path):
    """Detects the number of copies from a saved screenshot using OCR."""
    img = preprocess_image(image_path)
    if img is None:
        return 1
    
    print(f"📸 Processing image: {image_path}...")

    # Run OCR with improved settings
    extracted_text = pytesseract.image_to_string(img, config="--psm 11").strip()

    print("🔍 Extracted Text:\n", extracted_text)  # Debugging: Show extracted text

    # Extract copies field
    copies = extract_copies_from_text(extracted_text)
    print(f"🖨️ Detected Copies: {copies}")
    return copies

if __name__ == "__main__":
    image_path = "print_screenshot.png"  # Ensure this screenshot exists
    detected_copies = get_copies_from_screenshot(image_path)
    print(f"✅ Final Detected Copies: {detected_copies}")
