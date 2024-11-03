import win32com.client

def list_video_devices():
    """
    List all video capture devices (camera devices) on the system and return their device instance paths.
    """
    wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
    service = wmi.ConnectServer(".", "root\\cimv2")
    query = "SELECT * FROM Win32_PnPEntity WHERE Description LIKE '%camera%' OR Description LIKE '%video%'"
    devices = service.ExecQuery(query)

    device_paths = []
    for i, device in enumerate(devices):
        if device.PNPDeviceID:
            device_paths.append(device.PNPDeviceID)
            print(f"[{i}] {device.Description} - Device Path: {device.PNPDeviceID}")
    
    return device_paths

def choose_device(device_paths):
    """
    Prompt the user to choose a device from the list.
    """
    while True:
        try:
            choice = int(input("Enter the number of the camera device you want to use: "))
            if 0 <= choice < len(device_paths):
                return device_paths[choice]
            else:
                print("Invalid choice. Please select a number from the list.")
        except ValueError:
            print("Invalid input. Please enter a number.")

def main():
    print("Listing all available camera devices...")
    device_paths = list_video_devices()

    if not device_paths:
        print("No camera devices found.")
        return

    print("\nSelect a device from the list above.")
    chosen_device = choose_device(device_paths)

    print(f"\nYou selected device path: {chosen_device}")
    # You can pass `chosen_device` to your CameraCapture script for testing
    # For example:
    # CameraCapture.CaptureImage(chosen_device, "TestImage", "")

if __name__ == "__main__":
    main()
