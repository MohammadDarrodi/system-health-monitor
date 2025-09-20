import subprocess
import sys
import importlib
import platform
import shutil
import socket
import logging
from datetime import datetime
import os
import wmi
import psutil
import GPUtil
from colorama import Fore, Style

# Author Information
AUTHOR_INFO = "Created by Mohammad Darudi"

# Install required libraries if they are not already present
def install_and_import(lib_name):
    """
    Installs and imports a Python library if it's not already installed.
    """
    try:
        importlib.import_module(lib_name)
    except ImportError:
        print(f"Installing {lib_name}...")
        try:
            subprocess.check_call([sys.executable, "-m", "pip", "install", lib_name])
        except subprocess.CalledProcessError as e:
            print(f"Error installing {lib_name}: {e}")
            sys.exit(1)

# We'll use these libraries for system information
REQUIRED_LIBRARIES = ['psutil', 'GPUtil', 'colorama', 'wmi', 'platform']

try:
    for lib in REQUIRED_LIBRARIES:
        install_and_import(lib)
    import psutil
    import GPUtil
    from colorama import Fore, Style
    import wmi
    # Add a fallback for GPUtil on non-Windows systems
    if platform.system() != 'Windows' and 'wmi' in sys.modules:
        del sys.modules['wmi']
except ImportError as e:
    print(f"❌ Error importing a required library: {e}. Please ensure you have pip and an internet connection.")
    sys.exit(1)
except Exception as e:
    print(f"An unexpected error occurred during library setup: {e}")
    sys.exit(1)

# Configure logging to save the report to a file
log_file = f"system_health_report_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.log"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

def print_section(title):
    """
    Prints a formatted section title to the console and log file.
    """
    section_title = f"\n{Fore.CYAN}{'='*10} {title.upper()} {'='*10}{Style.RESET_ALL}"
    print(section_title)
    logging.info(f"--- {title} ---")

class SystemHealthMonitor:
    """
    A class to monitor and report on the health of a computer system.
    """
    def __init__(self):
        self.health_score = 100
        print(f"{Fore.MAGENTA}🩺 Starting system health check...{Style.RESET_ALL}")
        logging.info("Starting system health check...")
        print(f"{Fore.YELLOW}{AUTHOR_INFO}{Style.RESET_ALL}")

    def get_system_info(self):
        """
        Gathers and prints general system information.
        """
        print_section("System Information")
        info = {
            "Operating System": f"{platform.system()} {platform.release()}",
            "Architecture": platform.machine(),
            "Processor": platform.processor(),
            "Hostname": platform.node(),
            "Current Time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        for key, value in info.items():
            print(f"{key}: {value}")
            logging.info(f"{key}: {value}")

    def get_last_windows_update(self):
        """
        Finds and prints the most recent Windows update information.
        This feature is exclusive to Windows.
        """
        if platform.system() != "Windows":
            print("❌ This feature is only available for Windows.")
            return

        print_section("Last Windows Update")
        try:
            result = subprocess.check_output(
                'powershell "Get-HotFix | Sort-Object InstalledOn -Descending | Select-Object -First 1"',
                shell=True,
                text=True,
                stderr=subprocess.STDOUT
            )
            update_info = result.strip()
            if update_info:
                print(update_info)
                logging.info(f"Last update info:\n{update_info}")
            else:
                print("✅ No update information found. Your system may be up-to-date.")
                logging.info("No Windows update information found.")
        except subprocess.CalledProcessError as e:
            print(f"❌ Failed to get update info. Error: {e.output.strip()}")
            logging.error(f"Failed to get Windows update info: {e.output.strip()}")
        except Exception as e:
            print(f"❌ An unexpected error occurred while getting update info: {e}")
            logging.error(f"Failed to get Windows update info: {e}")

    def get_cpu_info(self):
        """
        Gathers and prints detailed CPU information, including usage and temperature (if available).
        """
        print_section("Processor (CPU) Information")
        try:
            usage = psutil.cpu_percent(interval=1)
            freq = psutil.cpu_freq()
            print(f"Core Count: {psutil.cpu_count(logical=True)}")
            print(f"Max Frequency: {freq.max:.2f} MHz")
            print(f"Current Frequency: {freq.current:.2f} MHz")
            print(f"CPU Usage: {usage}%")

            # Check for CPU temperature
            if hasattr(psutil, 'sensors_temperatures'):
                temps = psutil.sensors_temperatures()
                if 'coretemp' in temps:
                    cpu_temp = temps['coretemp'][0].current
                    print(f"CPU Temperature: {cpu_temp}°C")
                    if cpu_temp > 85:
                        self.health_score -= 15
                        print(f"{Fore.RED}⚠️ WARNING: High CPU temperature detected ({cpu_temp}°C).{Style.RESET_ALL}")
                        logging.warning(f"High CPU temperature: {cpu_temp}°C")
                else:
                    print("❕ CPU temperature information not available for this platform.")
            else:
                print("❌ psutil.sensors_temperatures() is not available on this system.")

            if usage > 85:
                self.health_score -= 10
                print(f"{Fore.RED}⚠️ WARNING: High CPU usage detected ({usage}%).{Style.RESET_ALL}")
                logging.warning(f"High CPU usage: {usage}%")
        except Exception as e:
            print(f"❌ Failed to get CPU information: {e}")
            logging.error(f"Failed to get CPU info: {e}")

    def get_ram_info(self):
        """
        Gathers and prints detailed RAM information.
        """
        print_section("Memory (RAM) Information")
        try:
            mem = psutil.virtual_memory()
            print(f"Total Memory: {mem.total / (1024**3):.2f} GB")
            print(f"Used Memory: {mem.used / (1024**3):.2f} GB")
            print(f"Memory Usage: {mem.percent}%")
            if mem.percent > 85:
                self.health_score -= 10
                print(f"{Fore.RED}⚠️ WARNING: High RAM usage detected ({mem.percent}%).{Style.RESET_ALL}")
                logging.warning(f"High RAM usage: {mem.percent}%")
        except Exception as e:
            print(f"❌ Failed to get RAM information: {e}")
            logging.error(f"Failed to get RAM info: {e}")

    def get_disk_info(self):
        """
        Gathers and prints detailed disk partition information.
        """
        print_section("Disk Information")
        for partition in psutil.disk_partitions():
            try:
                usage = psutil.disk_usage(partition.mountpoint)
                print(f"🔹 Drive: {partition.device}")
                print(f"  Total Space: {usage.total / (1024**3):.2f} GB")
                print(f"  Used Space: {usage.used / (1024**3):.2f} GB")
                print(f"  Free Space: {usage.free / (1024**3):.2f} GB")
                print(f"  Usage Percentage: {usage.percent}%")
                if usage.percent > 90:
                    self.health_score -= 10
                    print(f"{Fore.RED}⚠️ WARNING: Drive is nearly full ({usage.percent}%).{Style.RESET_ALL}")
                    logging.warning(f"High disk usage on {partition.device}: {usage.percent}%")
            except PermissionError:
                print(f"  ❌ Access denied for drive {partition.device}.")
            except Exception as e:
                print(f"  ❌ Failed to get disk information for {partition.device}: {e}")
                logging.error(f"Failed to get disk info for {partition.device}: {e}")

    def get_battery_info(self):
        """
        Gathers and prints battery information (if available).
        """
        print_section("Battery Information")
        battery = psutil.sensors_battery()
        if battery:
            print(f"Charge Level: {battery.percent}%")
            print(f"Status: {'Plugged In' if battery.power_plugged else 'On Battery'}")
            if battery.percent < 30 and not battery.power_plugged:
                self.health_score -= 10
                print(f"{Fore.RED}⚠️ WARNING: Low battery level and not plugged in.{Style.RESET_ALL}")
                logging.warning("Low battery level and not plugged in.")
        else:
            print("❕ Battery information not available.")

    def get_gpu_info(self):
        """
        Gathers and prints GPU information using GPUtil and WMI (as a fallback).
        """
        print_section("Graphics Card (GPU) Information")
        try:
            gpus = GPUtil.getGPUs()
            if gpus:
                for gpu in gpus:
                    print(f"🔹 NVIDIA GPU: {gpu.name}")
                    print(f"  Total Memory: {gpu.memoryTotal} MB")
                    print(f"  Used Memory: {gpu.memoryUsed} MB")
                    print(f"  Temperature: {gpu.temperature} °C")
                    print(f"  Load: {gpu.load * 100:.1f}%")
                    if gpu.load * 100 > 90 or gpu.temperature > 85:
                        self.health_score -= 10
                        print(f"{Fore.RED}⚠️ WARNING: High GPU temperature or load.{Style.RESET_ALL}")
                        logging.warning("High GPU temperature or load.")
            else:
                print("❕ No NVIDIA GPU found using GPUtil.")
                self._fallback_gpu_detection()
        except Exception as e:
            print(f"❌ GPUtil failed. Error: {e}")
            logging.error(f"GPUtil failed: {e}")
            self._fallback_gpu_detection()

    def _fallback_gpu_detection(self):
        """
        A fallback method to detect GPUs and display info using WMI on Windows.
        """
        try:
            if platform.system() == "Windows":
                c = wmi.WMI()
                gpu_list = c.Win32_VideoController()
                display_list = c.Win32_DesktopMonitor()

                if gpu_list:
                    print("🔍 Identified Graphics Cards:")
                    for gpu in gpu_list:
                        print(f"  - {gpu.Name}")
                        print(f"    Driver Version: {gpu.DriverVersion}")
                        
                if display_list:
                    print("🔍 Connected Displays:")
                    for monitor in display_list:
                        print(f"  - {monitor.Name}")
                        print(f"    Resolution: {monitor.ScreenWidth}x{monitor.ScreenHeight}")
                
                if not gpu_list and not display_list:
                    print("🔍 No graphics cards or displays identified via WMI.")
            else:
                print("🔍 WMI is only available on Windows. No fallback detection available.")
        except Exception as e:
            print(f"❌ GPU or display detection via WMI failed. Error: {e}")

    def get_network_info(self):
        """
        Gathers and prints network interface information.
        """
        print_section("Network Information")
        try:
            net = psutil.net_if_addrs()
            for interface, addresses in net.items():
                print(f"🔹 Interface: {interface}")
                for addr in addresses:
                    if addr.family == -1:
                        print(f"  MAC Address: {addr.address}")
                    elif addr.family == socket.AF_INET:
                        print(f"  IP Address: {addr.address}")
        except Exception as e:
            print(f"❌ Failed to get network information: {e}")
            logging.error(f"Failed to get network info: {e}")

    def get_advanced_hardware_info(self):
        """
        Gathers and prints detailed motherboard and BIOS information.
        """
        if platform.system() != "Windows":
            print_section("Advanced Hardware Information")
            print("❌ This feature is only available for Windows.")
            return
            
        print_section("Advanced Hardware Information")
        try:
            c = wmi.WMI()
            
            # Get Motherboard info
            motherboard_info = c.Win32_BaseBoard()[0]
            print(f"Motherboard Manufacturer: {motherboard_info.Manufacturer}")
            print(f"Motherboard Model: {motherboard_info.Product}")
            
            # Get BIOS info
            bios_info = c.Win32_BIOS()[0]
            print(f"BIOS Manufacturer: {bios_info.Manufacturer}")
            print(f"BIOS Version: {bios_info.SMBIOSBIOSVersion}")
            
        except Exception as e:
            print(f"❌ Failed to get advanced hardware information: {e}")
            logging.error(f"Failed to get advanced hardware info: {e}")

    def check_upgrade_compatibility(self):
        """
        Checks if the system meets the minimum requirements to upgrade to a newer Windows version.
        """
        if platform.system() != "Windows":
            print_section("Windows Upgrade Compatibility")
            print("❌ This check is only applicable for Windows operating systems.")
            return

        current_version = platform.release()
        current_version_major = int(current_version.split('.')[0])
        upgrade_target = None
        upgrade_requirements = {}

        if current_version_major == 7:
            upgrade_target = "Windows 8.1"
            upgrade_requirements = {
                "CPU_CORES": 1,
                "CPU_FREQ": 1000, # MHz
                "RAM_GB": 2,
                "DISK_GB": 40
            }
        elif current_version_major == 8:
            upgrade_target = "Windows 10"
            upgrade_requirements = {
                "CPU_CORES": 1,
                "CPU_FREQ": 1000,
                "RAM_GB": 2,
                "DISK_GB": 40
            }
        elif current_version_major == 10: # This covers both Windows 10 and 11
            # Check if it's Windows 10 or 11 based on build number
            build_number = int(platform.version().split('.')[2])
            if build_number >= 22000:
                print_section("Windows Upgrade Compatibility")
                print(f"{Fore.GREEN}✅ Your system is already running Windows 11 or a newer version. No upgrade check needed.{Style.RESET_ALL}")
                return
            
            upgrade_target = "Windows 11"
            upgrade_requirements = {
                "CPU_CORES": 2,
                "CPU_FREQ": 1000,
                "RAM_GB": 4,
                "DISK_GB": 64,
                "TPM_2": True,
                "SECURE_BOOT": True
            }
        else:
            print_section("Windows Upgrade Compatibility")
            print(f"❕ No specific upgrade path is defined for your current OS version: {current_version}.")
            return

        print_section(f"Upgrade Compatibility Check for {upgrade_target}")
        compatible = True

        # Check CPU
        try:
            cpu_cores = psutil.cpu_count(logical=True)
            cpu_freq = psutil.cpu_freq().current
            print(f"CPU Cores: {cpu_cores} (Required: {upgrade_requirements['CPU_CORES']}+)")
            print(f"CPU Frequency: {cpu_freq:.2f} MHz (Required: {upgrade_requirements['CPU_FREQ']}+ MHz)")
            if cpu_cores < upgrade_requirements['CPU_CORES'] or cpu_freq < upgrade_requirements['CPU_FREQ']:
                print(f"{Fore.RED}❌ FAILED: Processor does not meet minimum requirements.{Style.RESET_ALL}")
                compatible = False
            else:
                print(f"{Fore.GREEN}✅ PASSED: Processor meets minimum requirements.{Style.RESET_ALL}")
        except Exception:
            print(f"{Fore.RED}❌ FAILED: Could not check processor info.{Style.RESET_ALL}")
            compatible = False

        # Check RAM
        try:
            ram_gb = psutil.virtual_memory().total / (1024**3)
            print(f"Total RAM: {ram_gb:.2f} GB (Required: {upgrade_requirements['RAM_GB']}+ GB)")
            if ram_gb < upgrade_requirements['RAM_GB']:
                print(f"{Fore.RED}❌ FAILED: Insufficient RAM.{Style.RESET_ALL}")
                compatible = False
            else:
                print(f"{Fore.GREEN}✅ PASSED: RAM meets minimum requirements.{Style.RESET_ALL}")
        except Exception:
            print(f"{Fore.RED}❌ FAILED: Could not check RAM info.{Style.RESET_ALL}")
            compatible = False

        # Check Storage
        try:
            disk_usage = psutil.disk_usage('C:\\')
            disk_gb = disk_usage.total / (1024**3)
            print(f"System Drive Storage: {disk_gb:.2f} GB (Required: {upgrade_requirements['DISK_GB']}+ GB)")
            if disk_gb < upgrade_requirements['DISK_GB']:
                print(f"{Fore.RED}❌ FAILED: Insufficient storage on system drive.{Style.RESET_ALL}")
                compatible = False
            else:
                print(f"{Fore.GREEN}✅ PASSED: Storage meets minimum requirements.{Style.RESET_ALL}")
        except Exception:
            print(f"{Fore.RED}❌ FAILED: Could not check storage info.{Style.RESET_ALL}")
            compatible = False
            
        # Check TPM and Secure Boot for Windows 11 only
        if upgrade_target == "Windows 11":
            try:
                c = wmi.WMI()
                # Check Secure Boot
                secure_boot = c.Win32_OperatingSystem()[0].SecureBoot
                print(f"Secure Boot: {'Enabled' if secure_boot else 'Disabled'} (Required: Enabled)")
                if not secure_boot:
                    print(f"{Fore.RED}❌ FAILED: Secure Boot is not enabled.{Style.RESET_ALL}")
                    compatible = False
                else:
                    print(f"{Fore.GREEN}✅ PASSED: Secure Boot is enabled.{Style.RESET_ALL}")

                # Check TPM 2.0
                tpm_status = c.Win32_Tpm()[0]
                tpm_version = tpm_status.SpecVersion
                is_tpm_2_0 = '2.0' in tpm_version if tpm_version else False
                
                print(f"TPM Version: {tpm_version} (Required: 2.0)")
                if not is_tpm_2_0:
                    print(f"{Fore.RED}❌ FAILED: TPM 2.0 is not detected or enabled.{Style.RESET_ALL}")
                    compatible = False
                else:
                    print(f"{Fore.GREEN}✅ PASSED: TPM 2.0 is detected.{Style.RESET_ALL}")
            except Exception as e:
                print(f"{Fore.RED}❌ FAILED: An error occurred while checking TPM/Secure Boot: {e}{Style.RESET_ALL}")
                compatible = False
                
        # Final result
        print("\n" + "="*30)
        if compatible:
            print(f"{Fore.GREEN}✅ Your PC can be upgraded to {upgrade_target}.{Style.RESET_ALL}")
        else:
            print(f"{Fore.RED}❌ Your PC does not meet the minimum requirements for {upgrade_target}.{Style.RESET_ALL}")
        print("="*30)
        logging.info(f"Upgrade compatibility for {upgrade_target} Result: {'Compatible' if compatible else 'Not Compatible'}")
    
    def run_all_checks(self):
        """
        Runs all system health checks sequentially.
        """
        self.get_system_info()
        self.get_last_windows_update()
        self.get_cpu_info()
        self.get_ram_info()
        self.get_disk_info()
        self.get_battery_info()
        self.get_gpu_info()
        self.get_network_info()
        self.get_advanced_hardware_info()
        self.check_upgrade_compatibility()
        self.show_health_report()
        logging.info("System health check finished.")

    def show_health_report(self):
        """
        Prints the final system health report and score.
        """
        print_section("System Health Report")
        color = Fore.GREEN if self.health_score >= 80 else Fore.YELLOW if self.health_score >= 50 else Fore.RED
        message = (
            "✅ The system is in good condition." if self.health_score >= 80
            else "⚠️ Some areas need attention." if self.health_score >= 50
            else "❌ The system is in poor condition. Professional inspection is recommended."
        )
        report_line = f"{color}🩺 System Health Score: {self.health_score}%{Style.RESET_ALL}"
        print(report_line)
        print(message)
        logging.info(f"Final health score: {self.health_score}% - {message}")

if __name__ == "__main__":
    monitor = SystemHealthMonitor()
    monitor.run_all_checks()