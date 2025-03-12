import win32print

def get_print_job_copies(printer_name):
    # Open the printer
    printer_handle = win32print.OpenPrinter(printer_name)
    
    try:
        # Get all print jobs in the queue
        jobs = win32print.EnumJobs(printer_handle, 0, -1, 2)  # Level 2 for detailed job info
        
        if not jobs:
            return "No active print jobs found."
        
        for job in jobs:
            job_id = job["JobId"]
            job_info = win32print.GetJob(printer_handle, job_id, 2)  # Get detailed job info (JOB_INFO_2)
            dev_mode = job_info.get("pDevMode", None)  # Extract DEVMODE
            
            if dev_mode:
                copies = dev_mode.dmCopies
                print(f"Job ID: {job_id}, Copies: {copies}")
            else:
                print(f"Job ID: {job_id}, DEVMODE not available")
    
    finally:
        win32print.ClosePrinter(printer_handle)

# Example Usage
printer_name = "EPSON L4160 Series"
get_print_job_copies(printer_name)
