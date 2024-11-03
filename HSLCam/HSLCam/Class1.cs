using System;
using System.IO;
using System.Runtime.InteropServices;
using DirectShowLib;
using System.Runtime.InteropServices.ComTypes;

[StructLayout(LayoutKind.Sequential, Pack = 1)]
public struct BitmapInfoHeader
{
    public uint biSize;
    public int biWidth;
    public int biHeight;
    public ushort biPlanes;
    public ushort biBitCount;
    public uint biCompression;
    public uint biSizeImage;
    public int biXPelsPerMeter;
    public int biYPelsPerMeter;
    public uint biClrUsed;
    public uint biClrImportant;
}

namespace HSLCam
{
    public class CameraCapture
    {
        [DllImport("ole32.dll")]
        private static extern int CoInitializeEx(IntPtr pvReserved, uint dwCoInit);

        [DllImport("ole32.dll")]
        private static extern void CoUninitialize();

        private const uint COINIT_MULTITHREADED = 0x0;
        private const int S_OK = 0;
        private const int S_FALSE = 1;

        public static int CaptureImage(string deviceInstancePath, string filenamePrefix, out string errorMessage)
        {
            errorMessage = "Success";

            // Step 1: Initialize COM library with detailed exit code for failure
            int hr = CoInitializeEx(IntPtr.Zero, COINIT_MULTITHREADED);
            if (hr != 0)
            {
                errorMessage = $"COM initialization failed with HRESULT: 0x{hr:X8}";
                return 101; // Exit code 101: COM initialization failed
            }

            IGraphBuilder graphBuilder = null;
            ICaptureGraphBuilder2 captureGraphBuilder = null;
            ICreateDevEnum devEnum = null;
            IEnumMoniker enumMoniker = null;
            ISampleGrabber sampleGrabber = null;
            IBaseFilter videoDevice = null;
            IBaseFilter nullRenderer = null;
            IMediaControl mediaControl = null;

            try
            {
                // Step 2: Create filter graph
                graphBuilder = (IGraphBuilder)new FilterGraph();
                captureGraphBuilder = (ICaptureGraphBuilder2)new CaptureGraphBuilder2();
                captureGraphBuilder.SetFiltergraph(graphBuilder);
                if (graphBuilder == null || captureGraphBuilder == null)
                {
                    errorMessage = "Failed to create filter graph.";
                    return 102; // Exit code 102: Filter graph creation failed
                }

                // Step 3: Create device enumerator
                devEnum = (ICreateDevEnum)new CreateDevEnum();
                hr = devEnum.CreateClassEnumerator(FilterCategory.VideoInputDevice, out enumMoniker, 0);
                if (hr != 0 || enumMoniker == null)
                {
                    errorMessage = $"No video devices found or failed to create enumerator, HRESULT: 0x{hr:X8}";
                    return 103; // Exit code 103: Device enumerator creation failed
                }

                // Step 4: Locate specified video device
                IMoniker[] monikers = new IMoniker[1];
                bool deviceFound = false;

                while (enumMoniker.Next(1, monikers, IntPtr.Zero) == 0)
                {
                    IPropertyBag propertyBag = null;
                    try
                    {
                        monikers[0].BindToStorage(null, null, typeof(IPropertyBag).GUID, out object bagObj);
                        propertyBag = (IPropertyBag)bagObj;

                        propertyBag.Read("DevicePath", out object devPathObj, null);
                        string currentDevicePath = devPathObj as string;

                        if (currentDevicePath == deviceInstancePath)
                        {
                            monikers[0].BindToObject(null, null, typeof(IBaseFilter).GUID, out object filterObj);
                            videoDevice = (IBaseFilter)filterObj;
                            deviceFound = true;
                            break;
                        }
                    }
                    finally
                    {
                        if (propertyBag != null)
                            Marshal.ReleaseComObject(propertyBag);
                        if (monikers[0] != null)
                            Marshal.ReleaseComObject(monikers[0]);
                    }
                }

                if (!deviceFound)
                {
                    errorMessage = "Specified device instance path not found.";
                    return 104; // Exit code 104: Specified device path not found
                }

                // Step 5: Add video device to graph
                hr = graphBuilder.AddFilter(videoDevice, "Video Capture");
                if (hr != 0)
                {
                    errorMessage = $"Could not add video device to graph, HRESULT: 0x{hr:X8}";
                    return 105; // Exit code 105: Adding video device to graph failed
                }

                // Step 6: Set up sample grabber
                sampleGrabber = (ISampleGrabber)new SampleGrabber();
                AMMediaType mediaType = new AMMediaType
                {
                    majorType = MediaType.Video,
                    subType = MediaSubType.RGB24
                };
                hr = sampleGrabber.SetMediaType(mediaType);
                if (hr != 0)
                {
                    errorMessage = $"Could not set media type in sample grabber, HRESULT: 0x{hr:X8}";
                    return 106; // Exit code 106: Setting media type failed
                }

                hr = sampleGrabber.SetBufferSamples(true);
                if (hr != 0)
                {
                    errorMessage = $"Could not enable sample buffering in sample grabber, HRESULT: 0x{hr:X8}";
                    return 107; // Exit code 107: Enabling sample buffering failed
                }

                // Step 7: Add sample grabber to filter graph
                hr = graphBuilder.AddFilter((IBaseFilter)sampleGrabber, "Sample Grabber");
                if (hr != 0)
                {
                    errorMessage = $"Could not add sample grabber to filter graph, HRESULT: 0x{hr:X8}";
                    return 108; // Exit code 108: Adding sample grabber to graph failed
                }

                // Step 8: Add null renderer
                nullRenderer = (IBaseFilter)new NullRenderer();
                hr = graphBuilder.AddFilter(nullRenderer, "Null Renderer");
                if (hr != 0)
                {
                    errorMessage = $"Could not add null renderer to filter graph, HRESULT: 0x{hr:X8}";
                    return 109; // Exit code 109: Adding null renderer failed
                }

                // Step 9: Render video stream
                hr = captureGraphBuilder.RenderStream(PinCategory.Capture, MediaType.Video, videoDevice, sampleGrabber as IBaseFilter, nullRenderer);
                if (hr != 0)
                {
                    errorMessage = $"Could not render capture video stream, HRESULT: 0x{hr:X8}";
                    return 110; // Exit code 110: Rendering video stream failed
                }

                // Step 10: Run media control
                mediaControl = (IMediaControl)graphBuilder;
                hr = mediaControl.Run();
                if (hr != 0)
                {
                    errorMessage = $"Could not run the graph, HRESULT: 0x{hr:X8}";
                    return 111; // Exit code 111: Running media control failed
                }

                System.Threading.Thread.Sleep(100); // Allow buffer to fill
                return 0; // Success
            }
            finally
            {
                // Release COM objects
                if (enumMoniker != null)
                    Marshal.ReleaseComObject(enumMoniker);
                if (devEnum != null)
                    Marshal.ReleaseComObject(devEnum);
                if (captureGraphBuilder != null)
                    Marshal.ReleaseComObject(captureGraphBuilder);
                if (graphBuilder != null)
                    Marshal.ReleaseComObject(graphBuilder);
                if (sampleGrabber != null)
                    Marshal.ReleaseComObject(sampleGrabber);
                if (nullRenderer != null)
                    Marshal.ReleaseComObject(nullRenderer);
                if (videoDevice != null)
                    Marshal.ReleaseComObject(videoDevice);
                if (mediaControl != null)
                    Marshal.ReleaseComObject(mediaControl);

                CoUninitialize();
            }
        }
}
}