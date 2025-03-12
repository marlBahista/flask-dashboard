import serial

ser = serial.Serial(port='COM4', baudrate=9600, timeout=1)  # Added timeout

ser.flushInput()  # Clears any noise from the buffer before reading

while True:
    value = ser.readline().strip()  # Strip whitespace/newline
    try:
        valueInString = value.decode('utf-8', errors='ignore')  # Ignore undecodable bytes
        print(valueInString)
    except UnicodeDecodeError as e:
        print(f"Decoding error: {e}")  # Print error but continue running
