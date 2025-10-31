import os
import subprocess
import webbrowser
from pathlib import Path
import glob

# Save somewhere you have full permission
output_path = Path(os.environ["USERPROFILE"]) / "Documents" / "battery_report.html"

# Run powercfg and capture its real output
res = subprocess.run(
    ["powercfg", "/batteryreport", "/output", str(output_path)],
    capture_output=True, text=True, shell=True
)

print(res.stdout.strip())
if res.stderr.strip():
    print("STDERR:", res.stderr.strip())

# Verify and open
if output_path.exists():
    print(f"✅ Battery report saved to: {output_path}")
    webbrowser.open(str(output_path))
else:
    print("⚠️ Report not found at expected path. Searching common locations...")
    candidates = []
    homes = [
        Path(os.environ["USERPROFILE"]),
        Path(os.environ["USERPROFILE"]) / "Documents",
        Path(os.environ["USERPROFILE"]) / "Desktop",
        Path("C:/"),
    ]
    for h in homes:
        candidates += glob.glob(str(h / "battery*.html"))
    if candidates:
        print("Found these possible reports:")
        for p in candidates:
            print(" -", p)
    else:
        print("No report files found. Try running your terminal as Administrator.")
