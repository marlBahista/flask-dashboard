import os
import tkinter as tk
from tkinter import filedialog, messagebox
import win32print
import win32api
from PIL import Image, ImageTk
import time  # ✅ Add this line

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[
        ("Documents", "*.docx;*.pdf;*.txt"),
        ("Images", "*.png;*.jpg;*.jpeg"),
        ("All Files", "*.*")
    ])
    if file_path:
        entry_file.delete(0, tk.END)
        entry_file.insert(0, file_path)
        preview_file(file_path)

def preview_file(file_path):
    # Only preview images for now
    if file_path.lower().endswith((".png", ".jpg", ".jpeg")):
        img = Image.open(file_path)
        img = img.resize((200, 200))
        img_tk = ImageTk.PhotoImage(img)
        label_preview.config(image=img_tk)
        label_preview.image = img_tk
    else:
        label_preview.config(text="Preview not available")

def print_document():
    file_path = entry_file.get()
    copies = int(entry_copies.get())

    if not file_path:
        messagebox.showerror("Error", "No file selected!")
        return
    
    # Log the print job
    log_file = os.path.join(os.environ["USERPROFILE"], "Desktop", "PrintLog.txt")
    with open(log_file, "a") as file:
        file.write(f"Date: {time.ctime()}\n")
        file.write(f"Document: {file_path}\n")
        file.write(f"Copies: {copies}\n")
        file.write(f"Printer: {win32print.GetDefaultPrinter()}\n")
        file.write("-" * 30 + "\n")

    # Send print command
    for _ in range(copies):
        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
    
    messagebox.showinfo("Success", "Printing started!")

# GUI Setup
root = tk.Tk()
root.title("Custom Printing App")
root.geometry("400x350")

tk.Label(root, text="Select File:").pack()
entry_file = tk.Entry(root, width=40)
entry_file.pack()
tk.Button(root, text="Browse", command=browse_file).pack()

tk.Label(root, text="Number of Copies:").pack()
entry_copies = tk.Entry(root, width=5)
entry_copies.insert(0, "1")
entry_copies.pack()

label_preview = tk.Label(root, text="No Preview")
label_preview.pack()

tk.Button(root, text="Print", command=print_document).pack()

root.mainloop()
