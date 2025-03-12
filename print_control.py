import os
import ctypes
import serial
import pymsgbox
import time
import win32print

# Define cost per page
BW_COST_PER_PAGE = 5  # ₱5 per black-and-white page
COLOR_COST_PER_PAGE = 10  # ₱10 per colored page
balance = 0
PRINTER_NAME = "EPSON L360 Series"  # Update with your printer name
ESP32_PORT = "COM4"  # Update with actual ESP32 port

# Ensure script runs as Administrator
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", "python", __file__, None, 1)
    exit()

# Connect to ESP32
try:
    ser = serial.Serial(ESP32_PORT, 115200, timeout=1, dsrdtr=False)
    print(f"✅ ESP32 Connected on {ESP32_PORT}")
except serial.SerialException:
    pymsgbox.alert("Error: Could not connect to ESP32. Check your COM port.", "Serial Connection Error")
    exit()

# **Detect print jobs and get page count & color mode**
def get_print_job_details():
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)  # Get all print jobs
    
    if jobs:
        last_job = jobs[-1]  # Get the latest print job
        job_id = last_job.get("JobId", "Unknown")  # Default to "Unknown" if missing
        document_name = last_job.get("Document", "Unknown")  # Default to "Unknown"
        total_pages = last_job.get("TotalPages", 1)  # Default to 1 if missing
        color_mode = "Color" if last_job.get("Color", False) else "BlackAndWhite"  # Default to B&W if missing
        win32print.ClosePrinter(printer_handle)
        return job_id, document_name, total_pages, color_mode
    
    win32print.ClosePrinter(printer_handle)
    return None, None, 0, "Unknown"  # No active print jobs

# **Send command to ESP32**
def send_esp32_command(command):
    ser.write((command + "\n").encode())
    print(f"🔹 Sent to ESP32: {command}")

try:
    print("🔍 Monitoring print jobs...")

    while True:
        job_id, doc_name, pages, color_mode = get_print_job_details()
        
        if pages > 0:  # If a print job is detected
            total_cost = (pages * COLOR_COST_PER_PAGE) if color_mode == "Color" else (pages * BW_COST_PER_PAGE)

            print(f"🖨 Print job detected: {doc_name} ({pages} pages, {color_mode}). Total cost: ₱{total_cost}")
            send_esp32_command("OFF")  # Keep the printer OFF

            # **Prompt the user for payment**
            while balance < total_cost:
                pymsgbox.alert(f"Insert ₱5 coins until balance reaches ₱{total_cost}. Current balance: ₱{balance}", "Payment Required")

                if ser.in_waiting > 0:
                 coin = ser.readline().strip()
    
                try:
                    coinInString = coin.decode('utf-8', errors='ignore')  # Ignore invalid bytes
                    print(f"🔹 Received from ESP32: {coinInString}")

                    if "COIN_INSERTED:5" in coinInString:
                        balance += 5
                        print(f"💰 Balance: ₱{balance} / ₱{total_cost}")
                except UnicodeDecodeError:
                    print("⚠️ Skipped invalid serial data")


            # **Payment complete → Turn ON Printer**
            pymsgbox.alert("✅ Payment complete! Printing now...", "Payment Status")
            send_esp32_command("ON")
            time.sleep(5)  # Give time for printer to start

            # **After Printing, Turn Off Printer**
            time.sleep(10)  # Allow enough time for print job to finish
            send_esp32_command("OFF")

            balance = 0  # Reset balance

        time.sleep(1)  # Avoid overloading CPU

except Exception as e:
    print(f"❌ ERROR: {e}")
    input("Press Enter to exit...")
