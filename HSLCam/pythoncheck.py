import clr  # pythonnet library for .NET interop
import sys
from pathlib import Path
from System import String
from System.Reflection import Assembly, BindingFlags
import win32com.client
from System.Text import StringBuilder  # Using StringBuilder for mutable text

# Path to the DLL
dll_path = r"C:\Users\milot\Documents\Github\CommandCam\HSLCam\HSLCam\bin\Debug\net48\HSLCam.dll"

# Add the directory to the system path (required for loading the DLL)
sys.path.append(str(Path(dll_path).parent))

# Load the DLL
clr.AddReference("HSLCam")

# Load the assembly
assembly = Assembly.LoadFile(dll_path)

# Get the CameraCapture type from the HSLCam namespace
camera_capture_type = assembly.GetType("HSLCam.CameraCapture")

if camera_capture_type is None:
    print("Failed to find CameraCapture class in HSLCam namespace.")
    sys.exit(1)

# Verify the CaptureImage method
method_info = camera_capture_type.GetMethod("CaptureImage", BindingFlags.Public | BindingFlags.Static)
if method_info:
    print("Method found with signature:")
    for parameter in method_info.GetParameters():
        print(f"{parameter.Name}: {parameter.ParameterType}")
else:
    print("CaptureImage method not found.")
    sys.exit(1)

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

def capture_image(device_path, filename_prefix="TestImage"):
    """
    Calls the CaptureImage method from the HSLCam.CameraCapture class with the selected device path.
    """
    # Convert Python strings to System.String
    device_path = String(device_path)
    filename_prefix = String(filename_prefix)

    # Prepare an error message as a mutable container using StringBuilder
    error_message = StringBuilder()  # Initialize with an empty mutable StringBuilder

    # Call the CaptureImage method directly without ParameterModifier
    try:
        result = method_info.Invoke(
            None,  # Static method, so no instance required
            [device_path, filename_prefix, error_message]  # Arguments for the method
        )
        
        # Display results
        print(f"Result Code: {result}")
        print(f"Error Message: {error_message.ToString()}")  # Convert StringBuilder to a string

        # Check if the output file was created based on the prefix
        save_path = Path(r"C:\Program Files (x86)\HAMILTON\LogFiles\HSLCamera")
        file_exists = any(save_path.glob(f"{filename_prefix}_*.bmp"))
        print("Output file created:", file_exists)

    except Exception as ex:
        print("An error occurred while invoking CaptureImage:", ex)

def main():
    print("Listing all available camera devices...")
    device_paths = list_video_devices()

    if not device_paths:
        print("No camera devices found.")
        return

    print("\nSelect a device from the list above.")
    chosen_device = choose_device(device_paths)

    print(f"\nYou selected device path: {chosen_device}")
    
    # Pass the selected device path to the CaptureImage method in the DLL
    capture_image(chosen_device)

if __name__ == "__main__":
    main()
