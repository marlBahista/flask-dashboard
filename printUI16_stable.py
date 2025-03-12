import sys
import ctypes
import serial
import time
import wmi
import win32print
import tkinter as tk
import requests

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

# Constants
BW_COST_PER_PAGE = 5  
PRINTER_NAME = "EPSON L4160 Series"
ESP32_PORT = "COM4"

# Paper tracking
PAPER_FEED_CAPACITY = 18  
LOW_PAPER_THRESHOLD = 5  
remaining_paper = PAPER_FEED_CAPACITY  

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

inserted_coins = 0  
total_cost = 0  

def log_message(message):
    log_box.insert(tk.END, message + "\n")
    log_box.see(tk.END)

def update_paper_status():
    """Update paper count in UI."""
    paper_status_label.config(text=f"📄 Paper Remaining: {remaining_paper}")

def clear_print_queue():
    """Clears all print jobs from the queue."""
    log_message("🗑️ Clearing all remaining print jobs...")
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)

    for job in jobs:
        try:
            win32print.SetJob(printer_handle, job["JobId"], 0, None, win32print.JOB_CONTROL_DELETE)
            log_message(f"🗑️ Deleted job {job['JobId']}.")
        except Exception as e:
            log_message(f"⚠️ Error deleting job {job['JobId']}: {e}")

    win32print.ClosePrinter(printer_handle)
    log_message("✅ Print queue cleared.")

def set_printer_offline():
    """Sets the printer to offline mode."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = True
            printer.Put_()
            log_message(f"🚫 {PRINTER_NAME} is now OFFLINE.")

def set_printer_online():
    """Sets the printer to online mode."""
    c = wmi.WMI()
    for printer in c.Win32_Printer():
        if printer.Name == PRINTER_NAME:
            printer.WorkOffline = False
            printer.Put_()
            log_message(f"✅ {PRINTER_NAME} is now ONLINE.")

# Clear queue and set offline at startup
clear_print_queue()
set_printer_offline()

def compute_cost():
    """Calculate the total printing cost."""
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)
    if not jobs:
        return 0
    total_pages = sum(job.get('TotalPages', 0) for job in jobs)
    return total_pages * BW_COST_PER_PAGE  

def request_coin_data():
    """Enable coin insertion detection from ESP32."""
    ser.write(b"REQUEST_COIN_DATA\n")
    log_message("💰 Coin insertion enabled.")

def listen_for_coins():
    """Continuously listens for coin insertions from ESP32."""
    global inserted_coins, total_cost

    if ser.in_waiting > 0:
        response = ser.readline().decode(errors='ignore').strip()
        log_message(f"ESP32 Response: {response}")
        if response.isdigit():
            inserted_coins = int(response)
            coin_label.config(text=f"Coins Inserted: ₱{inserted_coins}")

            if inserted_coins >= total_cost and total_cost > 0:
                log_message("✅ Payment complete. Printing now...")
                root.after(1000, set_printer_online)
                root.after(1500, monitor_printing)
                ser.write(b"RESET\n")
                return  

    root.after(200, listen_for_coins)  

def monitor_printing(timeout=15):
    """Monitors printing status, ensures job completion, and clears stuck jobs."""
    global remaining_paper
    start_time = time.time()  # Track how long we've been checking

    while True:
        printer_handle = win32print.OpenPrinter(PRINTER_NAME)
        jobs = win32print.EnumJobs(printer_handle, 0, -1, 2)  # Get detailed job list
        win32print.ClosePrinter(printer_handle)

        # Check for active printing jobs
        active_jobs = [
            job for job in jobs if job.get("Status", 0) & win32print.JOB_STATUS_PRINTING
        ]

        if active_jobs:
            log_message(f"⏳ Printing in progress... {len(active_jobs)} job(s) still printing.")
            time.sleep(2)
            continue  # Keep monitoring

        # If jobs exist but none are actively printing, check timeout
        if jobs:
            elapsed = time.time() - start_time
            if elapsed >= timeout:
                log_message("⚠️ Printing seems done but jobs are stuck. Force clearing queue.")
                break  # Proceed to clear queue after timeout
            log_message("⏳ Jobs still in queue but not printing. Checking again...")
            time.sleep(2)
            continue  # Keep checking

        # If no jobs remain, break out of the loop
        break

    # At this point, all jobs should be done or timeout is reached
    log_message("✅ Printing completed. Clearing queue and setting offline...")
    clear_print_queue()

    # Reduce paper count
    pages_used = total_cost // BW_COST_PER_PAGE  
    remaining_paper -= pages_used
    if remaining_paper < 0:
        remaining_paper = 0  

    requests.post("http://127.0.0.1:5000/add_transaction", json={"date": time.strftime("%Y-%m-%d %H:%M:%S"), "pages": pages_used, "cost": total_cost})
    requests.post("http://127.0.0.1:5000/update_paper", json={"remaining_paper": remaining_paper})

    log_message(f"✅ Printed {pages_used} pages. Cost: ₱{total_cost}. Paper left: {remaining_paper} sheets.")
    set_printer_offline()
    update_paper_status()
    root.after(1000, reset_transaction)



def reset_transaction():
    """Resets all variables and UI for a new transaction."""
    global inserted_coins, total_cost
    inserted_coins = 0  
    total_cost = 0  

    cost_label.config(text="Waiting for print job...")
    coin_label.config(text="Coins Inserted: ₱0")

    log_message("🔄 Ready for a new transaction.")
    set_printer_offline()

def main_loop():
    global total_cost, inserted_coins
    total_cost = compute_cost()
    
    if total_cost > 0 and inserted_coins < total_cost:
        set_printer_offline()
        cost_label.config(text=f"Printing cost: ₱{total_cost}")
        log_message(f"Printing cost: ₱{total_cost}. Insert coins to continue.")
        request_coin_data()  
        root.after(500, listen_for_coins)

    root.after(2000, main_loop)

root.after(2000, main_loop)
root.mainloop()
