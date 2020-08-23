# hardware-info-exporter
Prometheus Windows Hardware Info Exporter

This requires `OpenHardwareMonitorLib.dll` to interface with Windows.
It can be downloaded from here: https://openhardwaremonitor.org/

# Usage

Make sure you run it as `ADMIN` for acessing internal sensors.

1. Ensure `OpenHardwareMonitorLib.dll` is in the same folder as `exporter.py`
2. Update variables `PORT` and `SCRAPE_INTERVAL` as needed.
3. Run `exporter.py`

# Installation

`exporter.py` can be ran in Windows as you wish (recommended Windows Service). Consider using pyinstaller to generate standalone binary.
