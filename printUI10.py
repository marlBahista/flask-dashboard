#  Integrated Fixes:
# ✅ Added 3-second delay before checking print queue (compute_cost()).
# ✅ Checks if the printer is active using WMI (is_printer_active()).
# ✅ Forces print queue refresh after printing (clear_print_queue()).
# ✅ Tracks paper count and warns owner when low (update_paper_display()).
# ✅ Allows manual paper refill (refill_paper() with UI button).


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

if not is_admin():
    print("🔴 Restarting with Admin privileges...")
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

# Pricing and Paper Tracking Constants
BW_COST_PER_PAGE = 5  # ₱5 per black-and-white page
COLOR_COST_PER_PAGE = 10  # ₱10 per colored page
PRINTER_NAME = "EPSON L360 Series"
ESP32_PORT = "COM4"
PAPER_FEED_CAPACITY = 7  # Default paper loaded initially
LOW_PAPER_THRESHOLD = 2  # Warning when paper is low

# Global Variables
inserted_coins = 0
total_cost = 0
remaining_paper = PAPER_FEED_CAPACITY  # Tracks paper availability

# Connect to ESP32
try:
    ser = serial.Serial(ESP32_PORT, 115200, timeout=1)
    esp_status = f"✅ ESP32 Connected on {ESP32_PORT}"
except serial.SerialException:
    esp_status = "❌ Error: Could not connect to ESP32."

# Initialize UI
root = tk.Tk()
root.title("Printer Vendo System")
root.geometry("400x500")
status_label = tk.Label(root, text=esp_status, fg="green")
status_label.pack(pady=10)
cost_label = tk.Label(root, text="Waiting for print job...")
cost_label.pack(pady=10)
coin_label = tk.Label(root, text="Coins Inserted: ₱0")
coin_label.pack(pady=10)
paper_status_label = tk.Label(root, text=f"📄 Paper Remaining: {remaining_paper}", fg="blue")
paper_status_label.pack(pady=10)
log_box = tk.Text(root, height=10, width=50)
log_box.pack(pady=10)

def log_message(message):
    log_box.insert(tk.END, message + "\n")
    log_box.see(tk.END)

def set_printer_offline():
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = True
            printer.Put_()
            log_message(f"🚫 {PRINTER_NAME} is now OFFLINE.")

def set_printer_online():
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = False
            printer.Put_()
            log_message(f"✅ {PRINTER_NAME} is now ONLINE.")

set_printer_offline()

def is_printer_active():
    """Checks if the printer is currently active (even if the queue looks empty)."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME and printer.PrinterStatus in [3, 4]:  # 3 = Printing, 4 = Warm-up
            return True
    return False

def compute_cost():
    """Computes printing cost only if there are pending print jobs."""
    time.sleep(3)  # 🕒 Add delay to allow Windows to update the print queue
    
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)

    if not jobs and not is_printer_active():
        return 0  # No active print job detected
    total_pages = sum(job.get('TotalPages', 0) for job in jobs)
    return total_pages * BW_COST_PER_PAGE  

def update_coin_display():
    root.after(0, lambda: coin_label.config(text=f"Coins Inserted: ₱{inserted_coins}"))

def request_coin_data():
    ser.write(b"REQUEST_COIN_DATA\n")

def listen_for_coins():
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
                root.after(1000, set_printer_online)
                root.after(1500, process_print_job)
                ser.write(b"RESET\n")
                return  
    root.after(200, listen_for_coins)

def clear_print_queue():
    """Clears all pending print jobs to prevent false detection."""
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    win32print.SetPrinter(printer_handle, 2, None, 3)  # 3 = Refresh printer
    win32print.ClosePrinter(printer_handle)
    log_message("🗑️ Print queue cleared to avoid false detections.")

def process_print_job():
    """Handles print completion and paper tracking."""
    global remaining_paper
    log_message("📄 Printing in progress...")

    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)

    total_pages = sum(job.get('TotalPages', 0) for job in jobs)
    remaining_paper -= total_pages  # Subtract pages used
    if remaining_paper < 0:
        remaining_paper = 0  # Prevent negative values

    log_message(f"✅ Printing complete. {remaining_paper} pages left.")
    
    clear_print_queue()  # 🔥 Clears queue to avoid false detections

    set_printer_offline()
    update_paper_display()
    
    global inserted_coins, total_cost
    inserted_coins = 0
    total_cost = 0
    cost_label.config(text="Waiting for print job...")
    update_coin_display()

def update_paper_display():
    """Updates the UI with remaining paper count and warns if low."""
    if remaining_paper <= LOW_PAPER_THRESHOLD:
        paper_status_label.config(text=f"⚠️ Low Paper: {remaining_paper} sheets left!", fg="red")
        log_message("⚠️ Warning: Paper is running low!")
    else:
        paper_status_label.config(text=f"📄 Paper Remaining: {remaining_paper}", fg="blue")

def refill_paper():
    """Allows the owner to manually update the paper count."""
    global remaining_paper
    new_paper = tk.simpledialog.askinteger("Refill Paper", "Enter number of sheets added:", minvalue=1)
    if new_paper:
        remaining_paper += new_paper
        if remaining_paper > PAPER_FEED_CAPACITY:
            remaining_paper = PAPER_FEED_CAPACITY  # Cap at max tray capacity
        log_message(f"📄 Paper refilled. New count: {remaining_paper} sheets.")
        update_paper_display()

# Add a button to refill paper manually
refill_button = tk.Button(root, text="Refill Paper", command=refill_paper)
refill_button.pack(pady=10)

def main_loop():
    global total_cost, inserted_coins
    total_cost = compute_cost()
    
    update_paper_display()
    
    if total_cost > 0 and inserted_coins < total_cost:
        set_printer_offline()
        cost_label.config(text=f"Printing cost: ₱{total_cost}")
        log_message(f"Printing cost: ₱{total_cost}. Insert coins to continue.")
        request_coin_data()
        root.after(500, listen_for_coins)
    
    root.after(5000, main_loop)

root.after(2000, main_loop)
root.mainloop()
