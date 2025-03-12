from pywinauto import Application
import time

def get_word_print_copies():
    try:
        # Connect to the active Word print dialog
        app = Application().connect(title_re="Print", timeout=10)
        dlg = app.window(title_re="Print")

        # Locate the 'Copies' input field
        copies_box = dlg.child_window(control_type="Edit", found_index=1)  # Adjust index if needed
        copies = copies_box.get_value()

        print(f"Number of copies selected: {copies}")
        return int(copies)
    except Exception as e:
        print("Error detecting Word print dialog:", e)
        return None

# Wait for the user to open the Print dialog
print("Open the Print Dialog in MS Word now...")
time.sleep(5)  # Give time to open the Print window

copies = get_word_print_copies()
if copies:
    print(f"Detected {copies} copies in MS Word print.")
