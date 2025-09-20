# System Health Monitor

A comprehensive Python script to monitor and report on the health and performance of a computer system. This tool gathers detailed information about your hardware, operating system, and network, and provides a health score to help you identify potential issues.

## Features

- **General System Info**: Displays information about your OS, CPU, and host.
- **CPU & RAM Monitoring**: Provides real-time usage, frequency, and temperature (if available) for your CPU, and shows detailed RAM usage.
- **Disk Information**: Reports on disk partitions, total space, and usage percentage.
- **Battery Status**: Shows charge level and power source for laptops.
- **GPU & Advanced Hardware Info**: Detects and reports on graphics cards (NVIDIA, AMD, and others via WMI) and provides details on motherboard and BIOS.
- **Windows Update & Upgrade Compatibility**: Checks for the last Windows update and assesses if your system is compatible with a newer Windows version (e.g., from Windows 10 to 11).
- **Health Score**: Calculates a health score based on system performance indicators.
- **Logging**: Saves a detailed report to a log file for future reference.

## Requirements

The script requires the following Python libraries. It will automatically attempt to install them upon execution.

- `psutil`
- `GPUtil`
- `colorama`
- `wmi` (Windows only)

## How to Use

1.  Make sure you have Python installed on your system.
2.  Save the code from `system_health_monitor.py` to a file with the same name.
3.  Open a terminal or command prompt.
4.  Run the script using the following command:

    ```bash
    python system_health_monitor.py
    ```

5.  The script will print the system health report to the console and save a detailed log file in the same directory.