using System;
using HSLCam;

namespace CameraTest
{
    class Program
    {
        static void Main(string[] args)
        {
            string deviceInstancePath = "USB\\VID_13D3&PID_5405&MI_00\\6&140C2090&1&0000"; // Update with the correct path
            string filenamePrefix = "TestImage";
            string errorMessage;

            int result = CameraCapture.CaptureImage(deviceInstancePath, filenamePrefix, out errorMessage);

            Console.WriteLine($"Result Code: {result}");
            Console.WriteLine($"Error Message: {errorMessage}");
        }
    }
}