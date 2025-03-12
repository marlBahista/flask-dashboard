import os
import time

SPOOL_DIR = r"C:\Windows\System32\spool\PRINTERS"

def monitor_spool():
    print("📄 Monitoring print spooler for new print jobs...")
    existing_files = set(os.listdir(SPOOL_DIR))

    while True:
        time.sleep(1)
        new_files = set(os.listdir(SPOOL_DIR))
        added_files = new_files - existing_files

        for file in added_files:
            print(f"📑 New print job detected: {file}")
            if file.endswith(".SHD") or file.endswith(".SPL"):
                print(f"🔍 {file} needs to be analyzed!")
        
        existing_files = new_files

monitor_spool()
