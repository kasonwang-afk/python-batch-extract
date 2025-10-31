import wmi
from datetime import datetime
import csv
import time

# Connect to OpenHardwareMonitor WMI namespace
c = wmi.WMI(namespace="root\\OpenHardwareMonitor")

csv_file = "thermal_report.csv"
with open(csv_file, "w", newline="") as f:
    writer = csv.writer(f)
    writer.writerow(["Timestamp", "Sensor", "Value (°C)"])

print("🔍 Reading temperatures from OpenHardwareMonitor... Press Ctrl+C to stop.")

try:
    while True:
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for sensor in c.Sensor():
            if sensor.SensorType == "Temperature":
                print(f"[{timestamp}] {sensor.Name}: {sensor.Value:.1f} °C")
                with open(csv_file, "a", newline="") as f:
                    writer = csv.writer(f)
                    writer.writerow([timestamp, sensor.Name, sensor.Value])
        time.sleep(5)

except KeyboardInterrupt:
    print(f"\n✅ Report saved to {csv_file}")
