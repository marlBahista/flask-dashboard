import os
import win32print
import win32api
import time

def log_print_job(document_path, copies):
    log_file = os.path.join(os.environ["USERPROFILE"], "Desktop", "PrintLog.txt")
    with open(log_file, "a") as file:
        file.write(f"Date: {time.ctime()}\n")
        file.write(f"Document: {document_path}\n")
        file.write(f"Copies: {copies}\n")
        file.write(f"Printer: {win32print.GetDefaultPrinter()}\n")
        file.write("-" * 30 + "\n")

def print_document(document_path, copies):
    log_print_job(document_path, copies)
    for _ in range(copies):
        win32api.ShellExecute(0, "print", document_path, None, ".", 0)

# Example Usage: Manually set file & copies
file_path = r"C:\Users\marlo\Documents\testPrint.docx"  # Change this to your actual file
num_copies = 1  # Change this based on user input

print_document(file_path, num_copies)
