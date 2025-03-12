import sys
import ctypes
import serial
import time
import wmi
import pymsgbox
import win32print

# Check if the script is running as Admin
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# Restart script as Admin if not already
if not is_admin():
    print("🔴 Restarting with Admin privileges...")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# Pricing constants
BW_COST_PER_PAGE = 5  # ₱5 per black-and-white page
COLOR_COST_PER_PAGE = 10  # ₱10 per colored page
PRINTER_NAME = "EPSON L360 Series"
ESP32_PORT = "COM4"

# Connect to ESP32
try:
    ser = serial.Serial(ESP32_PORT, 115200, timeout=1)
    print(f"✅ ESP32 Connected on {ESP32_PORT}")
except serial.SerialException:
    pymsgbox.alert("Error: Could not connect to ESP32.")
    exit()

def set_printer_offline():
    """Sets the printer to offline mode using WMI."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = True
            printer.Put_()
            print(f"🚫 {PRINTER_NAME} is now OFFLINE.")

def set_printer_online():
    """Sets the printer to online mode using WMI."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = False
            printer.Put_()
            print(f"✅ {PRINTER_NAME} is now ONLINE.")

# Set printer to offline at startup
print("🔧 Setting printer offline on startup...")
set_printer_offline()

def compute_cost():
    """Computes printing cost only if there are pending print jobs."""
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)

    if not jobs:  # No pending print jobs
        return 0

    total_pages = 0
    color_pages = 0
    for job in jobs:
        total_pages += job.get('TotalPages', 0)
        if job.get('Document') and "color" in job['Document'].lower():
            color_pages += job['TotalPages']

    bw_pages = total_pages - color_pages
    total_cost = (bw_pages * BW_COST_PER_PAGE) + (color_pages * COLOR_COST_PER_PAGE)
    return total_cost

def wait_for_payment(amount):
    """Waits for ESP32 to confirm payment before resuming printing."""
    ser.write(f"{amount}\n".encode())  # Send cost to ESP32
    while True:
        if ser.in_waiting > 0:
            response = ser.readline().decode(errors='ignore').strip()
            print(f"ESP32 Response: {response}")
            if response == "PAID":
                set_printer_online()  # Bring printer online to resume printing
                return True

def wait_for_print_completion():
    """Waits for all print jobs to complete before setting the printer offline again."""
    print("📄 Waiting for print job to complete...")
    while True:
        time.sleep(2)  # Check every 2 seconds
        printer_handle = win32print.OpenPrinter(PRINTER_NAME)
        jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
        win32print.ClosePrinter(printer_handle)
        if not jobs:  # If there are no more print jobs, printing is done
            print("✅ Printing complete. Setting printer offline.")
            set_printer_offline()
            return

def main():
    while True:
        total_cost = compute_cost()
        if total_cost > 0:
            set_printer_offline()  # Ensure printer is offline before transaction
            pymsgbox.alert(f"Printing cost: ₱{total_cost}. Insert coins to continue.")
            if wait_for_payment(total_cost):
                print("✅ Payment received. Printing now...")
                wait_for_print_completion()  # Wait until printing is done and set printer offline
            else:
                print("❌ Payment failed. Printer remains offline.")
        time.sleep(2)

if __name__ == "__main__":
    main()
