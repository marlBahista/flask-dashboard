import sys
import ctypes
import serial
import time
import wmi
import pymsgbox
import win32print
import tkinter as tk
from tkinter import messagebox

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
    total_cost = total_pages * BW_COST_PER_PAGE  # Only count number of pages
    return total_cost

def update_coin_display(amount):
    root.after(0, lambda: coin_label.config(text=f"Coins Inserted: ₱{amount}"))
    """Updates UI with real-time coin insertion amount."""
    coin_label.config(text=f"Coins Inserted: ₱{amount}")

def wait_for_payment(amount):
    """Waits for ESP32 to confirm payment and updates UI."""
    ser.write(f"{amount}\n".encode())
    inserted_coins = 0
    while inserted_coins < amount:
        if ser.in_waiting > 0:
            response = ser.readline().decode(errors='ignore').strip()
            log_message(f"ESP32 Response: {response}")
            if response.isdigit():
                inserted_coins += int(response)
                update_coin_display(inserted_coins)
            if inserted_coins >= amount:
                log_message("✅ Payment complete. Printing now...")
                set_printer_online()
                return True
    return False

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
            update_coin_display(0)
            return

def main_loop():
    total_cost = compute_cost()
    if total_cost > 0:
        set_printer_offline()
        cost_label.config(text=f"Printing cost: ₱{total_cost}")
        log_message(f"Printing cost: ₱{total_cost}. Insert coins to continue.")
        messagebox.showinfo("Payment Required", f"Printing cost: ₱{total_cost}. Insert coins to continue.")
        if wait_for_payment(total_cost):
            wait_for_print_completion()
        else:
            log_message("❌ Payment failed. Printer remains offline.")
    root.after(2000, main_loop)

root.after(2000, main_loop)
root.mainloop()
