import pyautogui
import time
import os
import requests

def open_ink_levels():
    print("🔎 Searching for 'Printers & scanners'...")
    pyautogui.hotkey("win", "s")
    time.sleep(1)
    
    pyautogui.write("Printers & scanners")
    time.sleep(1)
    pyautogui.press("enter")
    time.sleep(3)  # Wait for the window to open

    print("🖨️ Selecting 'EPSON L4160 Series'...")
    printer_button = pyautogui.locateCenterOnScreen("epson_printer_button.png", confidence=0.8)
    if printer_button:
        pyautogui.click(printer_button)
    else:
        print("❌ Error: 'EPSON L4160 Series' button not found.")
        return
    time.sleep(2)

    print("⚙️ Opening 'Printing Preferences'...")
    preferences_button = pyautogui.locateCenterOnScreen("printing_preferences_button.png", confidence=0.8)
    if preferences_button:
        pyautogui.click(preferences_button)
    else:
        print("❌ Error: 'Printing Preferences' button not found.")
        return
    time.sleep(3)

    print("🎨 Clicking 'Ink Levels' button...")
    ink_levels_button = pyautogui.locateCenterOnScreen("ink_levels_button.png", confidence=0.8)
    if ink_levels_button:
        pyautogui.click(ink_levels_button)
        print("✅ Ink Levels window opened successfully!")
        time.sleep(3)  # Wait for the window to load
        take_screenshot_of_window()
    else:
        print("❌ Error: 'Ink Levels' button not found.")

def take_screenshot_of_window():

    print("📸 Capturing only the 'EPSON L4160 Series' window...")

    # Locate the window on screen
    window_location = pyautogui.locateOnScreen("epson_window_title.png", confidence=0.8)

    if window_location:
        x, y, width, height = window_location
        print(f"✅ Window found at ({x}, {y}, {width}, {height})")

        # Convert to integers to avoid assertion errors
        x, y, width, height = int(x), int(y), int(width), int(height)

        screenshot = pyautogui.screenshot(region=(x, y, width, height))
        screenshot_path = os.path.join(os.getcwd(), "capture.png")
        screenshot.save(screenshot_path)
        print(f"✅ Screenshot saved as {screenshot_path}")
    else:
        print("❌ Error: 'EPSON L4160 Series' window not found. Taking full screenshot as fallback.")
        pyautogui.screenshot("capture.png")

    print("❌ Closing all open windows...")
    pyautogui.hotkey("alt", "f4")  # Close the Ink Levels window
    time.sleep(1)
    pyautogui.hotkey("alt", "f4")  # Close Printing Preferences
    time.sleep(1)
    pyautogui.hotkey("alt", "f4")  # Close Printers & Scanners
    print("✅ All windows closed.")

import requests

def send_screenshot():
    screenshot_path = os.path.join(os.getcwd(), "static/capture.png")
    url = "http://127.0.0.1:5000/upload_screenshot"
    
    with open(screenshot_path, "rb") as img_file:
        files = {"file": img_file}
        response = requests.post(url, files=files)
    
    if response.status_code == 200:
        print("✅ Screenshot uploaded successfully!")
    else:
        print("❌ Error uploading screenshot.")

# Call after screenshot is taken
send_screenshot()    

import requests

def send_screenshot():
    screenshot_path = os.path.join(os.getcwd(), "capture.png")
    url = "http://127.0.0.1:5000/upload_screenshot"
    
    with open(screenshot_path, "rb") as img_file:
        files = {"file": img_file}
        response = requests.post(url, files=files)
    
    if response.status_code == 200:
        print("✅ Screenshot uploaded successfully!")
    else:
        print("❌ Error uploading screenshot.")

# Call this after taking the screenshot
send_screenshot()


# Run the function
open_ink_levels()
