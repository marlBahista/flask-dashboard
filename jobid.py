import win32print

def list_print_jobs(printer_name):
    # Open the printer
    printer_handle = win32print.OpenPrinter(printer_name)
    try:
        # Enumerate print jobs (level 1 -> JOB_INFO_1)
        jobs = win32print.EnumJobs(printer_handle, 0, -1, 1)
        if not jobs:
            print("No active print jobs found.")
            return None
        
        # Display job details
        for job in jobs:
            print(f"Job ID: {job['JobId']}, Document: {job['pDocument']}, Status: {job['Status']}")
            return job["JobId"]  # Return the first job ID
        
        return None  # No jobs found
    finally:
        win32print.ClosePrinter(printer_handle)

# Example usage
printer_name = "EPSON L4160 Series"  # Replace with actual printer name
job_id = list_print_jobs(printer_name)

if job_id:
    print(f"First active job ID: {job_id}")
else:
    print("No active jobs found.")
