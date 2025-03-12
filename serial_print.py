# Python Code:
import serial
import time
import win32print
import pymsgbox

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

def get_print_job_details():
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)
    printer_info = win32print.GetPrinter(printer_handle, 2)
    jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
    win32print.ClosePrinter(printer_handle)
    
    total_pages = 0
    color_pages = 0
    
    if jobs:
        for job in jobs:
            total_pages += job.get('TotalPages', 0)  # Use get() to avoid KeyError
            if job.get('Document') and "color" in job['Document'].lower():  # Check if key exists
                color_pages += job['TotalPages']
    
    return total_pages, color_pages

def compute_cost():
    total_pages, color_pages = get_print_job_details()
    bw_pages = total_pages - color_pages
    total_cost = (bw_pages * BW_COST_PER_PAGE) + (color_pages * COLOR_COST_PER_PAGE)
    return total_cost

def wait_for_payment(amount):
    ser.write(f"{amount}\n".encode())  # Send cost to ESP32
    while True:
        if ser.in_waiting > 0:
            response = ser.readline().decode(errors='ignore').strip()  # Ignore decoding errors
            print(f"ESP32 Response: {response}")  # Debug print
            if response == "PAID":
                ser.write(b"TRIGGER_PRINTER\n")  # Command ESP32 to turn on SSR
                return True

def wait_for_print_completion():
    print("📄 Waiting for print job to complete...")
    while True:
        time.sleep(2)  # Check print jobs every 2 seconds
        printer_handle = win32print.OpenPrinter(PRINTER_NAME)
        jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
        win32print.ClosePrinter(printer_handle)
        if not jobs:  # If there are no more print jobs, printing is done
            print("✅ Printing complete. Turning off printer.")
            ser.write(b"TURN_OFF_PRINTER\n")  # Command ESP32 to turn off SSR
            return

def main():
    while True:
        total_cost = compute_cost()
        if total_cost > 0:
            pymsgbox.alert(f"Printing cost: ₱{total_cost}. Insert coins to continue.")
            if wait_for_payment(total_cost):
                print("✅ Payment received. Printing now...")
                wait_for_print_completion()  # Wait until printing is done and turn off printer
            else:
                print("❌ Payment failed. Printing canceled.")
        time.sleep(2)

if __name__ == "__main__":
    main()
