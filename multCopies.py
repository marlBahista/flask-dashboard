import time
import win32print

PRINTER_NAME = "EPSON L4160 Series"  # Change this to your printer

def monitor_print_queue():
    printer_handle = win32print.OpenPrinter(PRINTER_NAME)

    print(f"🖨️ Monitoring print queue for {PRINTER_NAME}...\n")
    
    active_jobs = {}  # Dictionary to track job IDs and statuses

    try:
        while True:
            jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)  # Get all jobs in the queue
            
            if jobs:
                for job in jobs:
                    job_id = job["JobId"]
                    status = job.get("Status", "Unknown")
                    total_pages = job.get("TotalPages", "Unknown")
                    document_name = job.get("Document", "Unknown")

                    if job_id not in active_jobs:
                        print(f"📌 New Print Job Detected: {document_name} (ID: {job_id})")
                        print(f"   📄 Pages: {total_pages} | Status: {status}")
                        active_jobs[job_id] = status  # Track job status

                    else:
                        if active_jobs[job_id] != status:
                            print(f"🔄 Job {job_id} Status Updated: {status}")
                            active_jobs[job_id] = status

            else:
                print("✅ No active print jobs.")

            time.sleep(2)  # Check every 2 seconds

    except KeyboardInterrupt:
        print("\n🛑 Monitoring stopped.")

    finally:
        win32print.ClosePrinter(printer_handle)

if __name__ == "__main__":
    monitor_print_queue()
