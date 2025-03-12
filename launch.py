import pyautogui
import time

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
    else:
        print("❌ Error: 'Ink Levels' button not found.")

# Run the function
open_ink_levels()


# 1. Press Win + s
# 2. Type "Printers & scanners"
# 3. Press enter
# 4. Click button with label "EPSON L4160 Series"
# 5. Click button with label "Printing Preferences"
# 6. Click button with label "Ink Levels"