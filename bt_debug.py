import serial
import serial.tools.list_ports

# List available COM ports
ports = list(serial.tools.list_ports.comports())
print("Available COM ports:")
for port in ports:
    print(f" - {port.device}: {port.description}")

# Change to the correct COM port
BT_PORT = "COM5"  # Update this based on available ports

try:
    ser = serial.Serial(BT_PORT, 115200, timeout=1)
    print(f"✅ Connected to ESP32 on {BT_PORT}!")
except serial.SerialException as e:
    print(f"❌ Failed to connect: {e}")
