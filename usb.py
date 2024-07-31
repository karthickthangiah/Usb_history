import subprocess
import sys
import winreg
import datetime

# Function to install the pywin32 module if not installed
def install_pywin32():
    try:
        import winreg
        return winreg
    except ImportError:
        print("pywin32 module not found. Installing...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pywin32"])
        import winreg
        return winreg

# Ensure winreg is installed and import it
winreg = install_pywin32()

def get_usb_registry_info():
    usb_devices = []
    
    registry_path = r"SYSTEM\CurrentControlSet\Enum\USBSTOR"
    try:
        reg_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_path)
        for i in range(0, winreg.QueryInfoKey(reg_key)[0]):
            sub_key_name = winreg.EnumKey(reg_key, i)
            sub_key_path = registry_path + "\\" + sub_key_name
            sub_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, sub_key_path)
            
            for j in range(0, winreg.QueryInfoKey(sub_key)[0]):
                device_key_name = winreg.EnumKey(sub_key, j)
                device_key_path = sub_key_path + "\\" + device_key_name
                device_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, device_key_path)
                
                try:
                    device_info = {
                        "Device": sub_key_name,
                        "SerialNumber": device_key_name,
                    }
                    for k in range(0, winreg.QueryInfoKey(device_key)[1]):
                        value = winreg.EnumValue(device_key, k)
                        if value[0] == "FriendlyName":
                            device_info["FriendlyName"] = value[1]
                        elif value[0] == "DeviceDesc":
                            device_info["DeviceDesc"] = value[1]
                        elif value[0] == "InstallDate":
                            device_info["InstallDate"] = value[1]
                        elif value[0] == "LastWriteTime":
                            device_info["LastWriteTime"] = value[1]
                    
                    usb_devices.append(device_info)
                except Exception as e:
                    print(f"Error reading device key: {e}")
                finally:
                    winreg.CloseKey(device_key)
    except Exception as e:
        print(f"Error accessing registry: {e}")
    finally:
        winreg.CloseKey(reg_key)
    
    return usb_devices

def print_usb_registry_info(devices):
    print("\nUSB History from Registry:")
    print("-" * 60)
    for device in devices:
        print(f"Device: {device['Device']}")
        print(f"Serial Number: {device['SerialNumber']}")
        print(f"Friendly Name: {device.get('FriendlyName', 'N/A')}")
        print(f"Device Description: {device.get('DeviceDesc', 'N/A')}")
        last_write_time = device.get('LastWriteTime')
        if last_write_time:
            # Convert Windows FILETIME to datetime
            try:
                filetime = last_write_time / 10**7 - 11644473600
                print(f"Last Write Time: {datetime.datetime.fromtimestamp(filetime)}")
            except Exception as e:
                print(f"Error converting Last Write Time: {e}")
        else:
            print("Last Write Time: N/A")
        print("-" * 60)

if __name__ == "__main__":
    print("USB Forensic Analysis via Registry")
    print("=" * 60)
    
    usb_devices = get_usb_registry_info()
    print_usb_registry_info(usb_devices)
    
    print(f"\nTotal USB devices found in registry: {len(usb_devices)}")
