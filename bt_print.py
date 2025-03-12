import serial
import time
import win32print
import pymsgbox

PRINTER_NAME = "EPSON L360 Series"
BT_PORT = "COM5"  # Replace with your actual Bluetooth COM port
BW_COST = 5       # Cost per black & white page
COLOR_COST = 10   # Cost per colored page

def get_print_job_info():
    hprinter = win32print.OpenPrinter(PRINTER_NAME)
    jobs = win32print.EnumJobs(hprinter, 0, -1, 1)
    win32print.ClosePrinter(hprinter)

    if not jobs:
        return None

    job = jobs[0]
    return {
        "total_pages": job["TotalPages"],
        "color_mode": job["pDatatype"]  # Adjust based on actual color detection logic
    }

def compute_cost(pages, color):
    return pages * (COLOR_COST if color else BW_COST)

def main():
    try:
        ser = serial.Serial(BT_PORT, 115200, timeout=1)
        print("✅ Connected to ESP32 via Bluetooth!")
    except serial.SerialException:
        print("❌ Failed to connect to ESP32. Check Bluetooth pairing and COM port.")
        return

    while True:
        job_info = get_print_job_info()
        if job_info:
            pages = job_info["total_pages"]
            color = job_info["color_mode"] == "RAW"  # Adjust if needed
            cost = compute_cost(pages, color)
            
            pymsgbox.alert(f"Printing cost: ₱{cost}. Insert coins to continue.", "PisoPrinter")
            ser.write(f"PRINT_COST:{cost}\n".encode())

            while True:
                response = ser.readline().decode().strip()
                if response == "PRINT_OK":
                    print("✅ Printing started!")
                    break
                elif response == "NOT_ENOUGH_CREDIT":
                    pymsgbox.alert("Not enough credit. Please insert more coins.", "PisoPrinter")
                    ser.write(f"PRINT_COST:{cost}\n".encode())
                time.sleep(1)

if __name__ == "__main__":
    main()
