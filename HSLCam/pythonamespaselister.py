import clr  # pythonnet library for .NET interop
import sys
from pathlib import Path
from System.Reflection import Assembly

# Path to the DLL
dll_path = r"C:\Users\milot\Documents\Github\CommandCam\HSLCam\HSLCam\bin\Debug\net48\HSLCam.dll"

# Add the directory to the system path (required for loading the DLL)
sys.path.append(str(Path(dll_path).parent))

# Load the DLL
clr.AddReference("HSLCam")

# Load the assembly
assembly = Assembly.LoadFile(dll_path)

# Print all namespaces, types, and methods within the DLL
print(f"Assembly: {assembly.FullName}")
for type in assembly.GetTypes():
    print(f"  Namespace: {type.Namespace}")
    print(f"  Type: {type.Name}")
    
    # List all methods in the type
    for method in type.GetMethods():
        print(f"    Method: {method.Name}")
