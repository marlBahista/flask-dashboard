import sys
import ctypes
import serial
import time
import wmi
import win32print
import tkinter as tk

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
    esp_status = f"✅ ESP32 Connected on {ESP32_PORT}"
except serial.SerialException:
    esp_status = "❌ Error: Could not connect to ESP32."

# Initialize UI
root = tk.Tk()
root.title("Printer Vendo System")
root.geometry("400x400")
status_label = tk.Label(root, text=esp_status, fg="green")
status_label.pack(pady=10)
cost_label = tk.Label(root, text="Waiting for print job...")
cost_label.pack(pady=10)
coin_label = tk.Label(root, text="Coins Inserted: ₱0")
coin_label.pack(pady=10)
log_box = tk.Text(root, height=10, width=50)
log_box.pack(pady=10)

inserted_coins = 0  # Global variable to track inserted coins
total_cost = 0  # Global variable to store the printing cost

def log_message(message):
    log_box.insert(tk.END, message + "\n")
    log_box.see(tk.END)

def set_printer_offline():
    """Sets the printer to offline mode using WMI."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = True
            printer.Put_()
            log_message(f"🚫 {PRINTER_NAME} is now OFFLINE.")

def set_printer_online():
    """Sets the printer to online mode using WMI."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = False
            printer.Put_()
            log_message(f"✅ {PRINTER_NAME} is now ONLINE.")

# Set printer to offline at startup
set_printer_offline()

def compute_cost():
    """Computes printing cost only if there are pending print jobs."""
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)
    if not jobs:
        return 0
    total_pages = sum(job.get('TotalPages', 0) for job in jobs)
    return total_pages * BW_COST_PER_PAGE  # Only count number of pages

def update_coin_display():
    """Updates UI with real-time coin insertion amount."""
    root.after(0, lambda: coin_label.config(text=f"Coins Inserted: ₱{inserted_coins}"))

def request_coin_data():
    """Sends signal to ESP32 to start sending coin insertion data."""
    ser.write(b"REQUEST_COIN_DATA\n")

def listen_for_coins():
    """Listens for coin insertion from ESP32 in real time."""
    global inserted_coins, total_cost
    while ser.in_waiting > 0:
        response = ser.readline().decode(errors='ignore').strip()
        log_message(f"ESP32 Response: {response}")
        if response.isdigit():
            inserted_coins = int(response)
            update_coin_display()
            if inserted_coins >= total_cost and total_cost > 0:
                log_message("✅ Payment complete. Printing now...")
                update_coin_display()
                root.after(1000, set_printer_online)  # Add delay before executing set_printer_online
                root.after(1500, wait_for_print_completion)  # Ensures print completion starts after UI update
                ser.write(b"RESET\n")  # Reset coin count in ESP32
                return  # Exit loop after successful transaction
    root.after(200, listen_for_coins)  # Check every 200ms for faster response

def wait_for_print_completion():
    """Waits for all print jobs to complete before setting the printer offline again."""
    log_message("📄 Waiting for print job to complete...")
    while True:
        time.sleep(2)
        printer_handle = win32print.OpenPrinter(PRINTER_NAME)
        jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
        win32print.ClosePrinter(printer_handle)
        if not jobs:
            log_message("✅ Printing complete. Setting printer offline.")
            set_printer_offline()
            global inserted_coins, total_cost
            inserted_coins = 0  # Reset coin count for the next transaction
            total_cost = 0  # Reset printing cost
            cost_label.config(text="Waiting for print job...")
            update_coin_display()
            return

def main_loop():
    global total_cost, inserted_coins
    total_cost = compute_cost()
    if total_cost > 0 and inserted_coins < total_cost:
        set_printer_offline()
        cost_label.config(text=f"Printing cost: ₱{total_cost}")
        log_message(f"Printing cost: ₱{total_cost}. Insert coins to continue.")
        request_coin_data()  # Now tells ESP32 to start sending coin data
        root.after(500, listen_for_coins)  # Ensure it starts listening for coins immediately
    root.after(2000, main_loop)

root.after(2000, main_loop)
root.mainloop()
