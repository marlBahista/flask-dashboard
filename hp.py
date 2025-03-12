import win32print

printer_name = "HP Universal Printing PCL 6"  # Change to the installed printer name
printer_handle = win32print.OpenPrinter(printer_name)
printer_info = win32print.GetPrinter(printer_handle, 2)
devmode = printer_info["pDevMode"]

print("Detected Copies Field:", devmode.dmCopies)  # Check if dmCopies is exposed

win32print.ClosePrinter(printer_handle)
