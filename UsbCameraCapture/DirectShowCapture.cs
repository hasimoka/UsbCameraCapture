using DirectShowLib;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.Versioning;

namespace UsbCameraCapture
{

    public class DirectShowCapture : ISampleGrabberCB, IDisposable
    {

        #region Member variables

        private const double DPI = 96.0d;

        /// <summary> graph builder interface. </summary>
        private IFilterGraph2 _filterGraph;

        // Used to snap picture on Still pin
        private IAMVideoControl _videoControl;
        private IPin _pinStill;

        // <summary> buffer for bitmap data.Always release by caller</summary>
        private IntPtr _ipBuffer;

        private int _width;

        private int _height;

        private int _stride;

        #endregion

        #region APIs

        [DllImport("Kernel32.dll", EntryPoint = "RtlMoveMemory")]
        private static extern void CopyMemory(IntPtr Destination, IntPtr Source, [MarshalAs(UnmanagedType.U4)] int Length);

        #endregion

        private FixedSizedQueue<Tuple<DateTime, Bitmap>> _captureFrames;
        //private Bitmap _processingFrame;
        private Tuple<DateTime, Bitmap> _thumbnailFrame;

        public DirectShowCapture()
        {
            _filterGraph = null;

            _videoControl = null;
            _pinStill = null;

            _ipBuffer = IntPtr.Zero;
            _width = 0;
            _height = 0;
            _stride = 0;

            _captureFrames = new FixedSizedQueue<Tuple<DateTime, Bitmap>>();
            //_processingFrame = null;
            _thumbnailFrame = null;

            IsRunning = false;
        }

        public bool IsRunning { get; set; }

        public static List<DsDevice> GetCaptureDevices()
        {
            DsDevice[] capDevices;

            // Get the collection of video devices
            capDevices = DsDevice.GetDevicesOfCat(FilterCategory.VideoInputDevice);

            return capDevices.ToList<DsDevice>();
        }

        [SupportedOSPlatform("windows")]
        public static List<List<long>> GetVideoInfos(string devicePath, string name)
        {
            List<List<long>> results = new List<List<long>>();

            int hr;

            IBaseFilter captureFilter = null;
            IPin pinStill = null;

            // Get the graphbuilder object
            var filterGraph = new FilterGraph() as IFilterGraph2;

            try
            {
                IMoniker moniker = null;
                var devices = DirectShowCapture.GetCaptureDevices();
                foreach (var device in devices)
                {
                    if (device.DevicePath == devicePath)
                    {
                        moniker = device.Mon;
                    }
                }

                // add the video input device
                hr = filterGraph.AddSourceFilterForMoniker(moniker, null, name, out captureFilter);
                DsError.ThrowExceptionForHR(hr);

                // Find the still pin
                pinStill = DsFindPin.ByCategory(captureFilter, PinCategory.Still, 0);

                // Didn't find one.  Is there a preview pin?
                if (pinStill == null)
                {
                    pinStill = DsFindPin.ByCategory(captureFilter, PinCategory.Preview, 0);
                }

                // Still haven't found one.  Need to put a splitter in so we have
                // one stream to capture the bitmap from, and one to display.  Ok, we
                // don't *have* to do it that way, but we are going to anyway.
                if (pinStill == null)
                {
                    IPin pinRaw = null;

                    // Add a splitter
                    var smartTee = (IBaseFilter)new SmartTee();

                    try
                    {
                        hr = filterGraph.AddFilter(smartTee, "SmartTee");
                        DsError.ThrowExceptionForHR(hr);

                        // Find the find the capture pin from the video device and the
                        // input pin for the splitter, and connnect them
                        pinRaw = DsFindPin.ByCategory(captureFilter, PinCategory.Capture, 0);
                        IAMStreamConfig videoStreamConfig = pinRaw as IAMStreamConfig;

                        results = DirectShowCapture.GetStreamCapabilities(videoStreamConfig);
                    }
                    finally
                    {
                        if (pinRaw != null)
                        {
                            Marshal.ReleaseComObject(pinRaw);
                        }
                        if (pinRaw != smartTee)
                        {
                            Marshal.ReleaseComObject(smartTee);
                        }
                    }
                }
                else
                {
                    IAMStreamConfig videoStreamConfig = pinStill as IAMStreamConfig;

                    results = DirectShowCapture.GetStreamCapabilities(videoStreamConfig);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
            finally
            {
                try
                {
                    if (filterGraph != null)
                    {
                        IMediaControl mediaCtrl = filterGraph as IMediaControl;

                        // Stop the graph
                        hr = mediaCtrl.Stop();
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex);
                }

                if (filterGraph != null)
                {
                    Marshal.ReleaseComObject(filterGraph);
                    filterGraph = null;
                }

                if (pinStill != null)
                {
                    Marshal.ReleaseComObject(pinStill);
                    pinStill = null;
                }

            }

            return results;
        }

        private static List<List<long>> GetStreamCapabilities(IAMStreamConfig videoStreamConfig)
        {
            var results = new List<List<long>>();

            int hr;

            int capabilityCount;
            int capabilitySize;
            hr = videoStreamConfig.GetNumberOfCapabilities(out capabilityCount, out capabilitySize);
            DsError.ThrowExceptionForHR(hr);

            IntPtr taskMemPointer = IntPtr.Zero;
            AMMediaType mediaType = null;

            try
            {
                taskMemPointer = Marshal.AllocCoTaskMem(capabilitySize);

                for (int i = 0; i < capabilityCount; i++)
                {
                    hr = videoStreamConfig.GetStreamCaps(i, out mediaType, taskMemPointer);
                    DsError.ThrowExceptionForHR(hr);

                    var videoInfo = (VideoInfoHeader)Marshal.PtrToStructure(mediaType.formatPtr, typeof(VideoInfoHeader));
                    if (videoInfo.BmiHeader.Width != 0 && videoInfo.BmiHeader.Height != 0 && videoInfo.BmiHeader.BitCount != 0 && videoInfo.AvgTimePerFrame != 0)
                    {
                        results.Add(new List<long> { videoInfo.BmiHeader.Width, videoInfo.BmiHeader.Height, videoInfo.BmiHeader.BitCount, videoInfo.AvgTimePerFrame });
                    }
                }
            }
            finally
            {
                if (taskMemPointer != IntPtr.Zero)
                {
                    Marshal.FreeCoTaskMem(taskMemPointer);
                }
                if (mediaType != null)
                {
                    DsUtils.FreeAMMediaType(mediaType);
                }
            }

            return results;
        }

        private static byte[] BitmapToByteArray(Bitmap bmp)
        {
            Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
            BitmapData bmpData =
                bmp.LockBits(rect, ImageLockMode.ReadWrite,
                PixelFormat.Format32bppArgb);

            // Bitmapの先頭アドレスを取得
            IntPtr ptr = bmpData.Scan0;

            // 32bppArgbフォーマットで値を格納
            int bytes = bmp.Width * bmp.Height * 4;
            byte[] rgbValues = new byte[bytes];

            // Bitmapをbyte[]へコピー
            Marshal.Copy(ptr, rgbValues, 0, bytes);

            bmp.UnlockBits(bmpData);

            return rgbValues;
        }

        [SupportedOSPlatform("windows")]
        public int SampleCB(double sampleTime, IMediaSample mediaSample)
        {
            Marshal.ReleaseComObject(mediaSample);
            return 0;
        }

        public int BufferCB(double sampleTime, IntPtr buffer, int bufferLength)
        {
            // Save the buffer
            CopyMemory(_ipBuffer, buffer, bufferLength);

            var bitmap = new Bitmap(_width, _height, _stride, PixelFormat.Format24bppRgb, _ipBuffer);

            // 画像を180度回転する
            bitmap.RotateFlip(RotateFlipType.RotateNoneFlipY);

            var currentTime = DateTime.Now;
            _captureFrames.Enqueue(new Tuple<DateTime, Bitmap>(currentTime, bitmap));

            // サムネイル画像の保存
            if ((_thumbnailFrame == null) || ((currentTime - _thumbnailFrame.Item1).TotalSeconds >= 1.0))
            {
                _thumbnailFrame = new Tuple<DateTime, Bitmap>(currentTime, bitmap);
            }

            return 0;
        }

        [SupportedOSPlatform("windows")]
        public void Dispose()
        {
            CloseInterfaces();
        }

        [SupportedOSPlatform("windows")]
        public bool Start(string devicePath, int width, int height, short bitCount, long avgTimePerFrame)
        {
            if (IsRunning)
                return false;

            var result = false;
            try
            {
                var devices = DirectShowCapture.GetCaptureDevices();
                foreach (var device in devices)
                {
                    if (device.DevicePath == devicePath)
                    {
                        // Set up the capture graph
                        SetupGraph(device.Mon, device.Name, width, height, bitCount, avgTimePerFrame);

                        IsRunning = true;
                        result = true;
                    }
                }
            }
            catch
            {
                CloseInterfaces();
                throw;
            }

            return result;
        }

        [SupportedOSPlatform("windows")]
        public void Stop()
        {
            CloseInterfaces();

            IsRunning = false;

            _captureFrames.Clear();
            //_processingFrame = null;
            _thumbnailFrame = null;
        }

        public bool GetFrame(out string timestamp, out byte[] image, out int? height, out int? width)
        {
            bool result = false;

            timestamp = null;
            image = null;
            height = null;
            width = null;

            try
            {
                // キャプチャ画像を取得する
                var frameTuple = _captureFrames.Dequeue();

                //_processingFrame = frameTuple.Item2;

                timestamp = frameTuple.Item1.ToString("yyyy/MM/dd HH:mm:ss.ffffff");
                image = BitmapToByteArray(frameTuple.Item2);
                height = frameTuple.Item2.Height;
                width = frameTuple.Item2.Width;

                result = true;
            }
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine(ex);
            }

            return result;
        }

        public bool GetThumbnail(out string timestamp, out byte[] image, out int? height, out int? width)
        {
            bool result = false;

            timestamp = null;
            image = null;
            height = null;
            width = null;

            try
            {
                // サムネイル画像を取得する
                if (_thumbnailFrame != null)
                {
                    var frameTuple = _thumbnailFrame;

                    timestamp = frameTuple.Item1.ToString("yyyy/MM/dd HH:mm:ss.fff");
                    image = BitmapToByteArray(frameTuple.Item2);
                    height = frameTuple.Item2.Height;
                    width = frameTuple.Item2.Width;

                    result = true;
                }
            }
            catch (InvalidOperationException ex)
            {
                Debug.WriteLine(ex);
            }

            return result;
        }

        /// <summary> build the capture graph for grabber. </summary>
        [SupportedOSPlatform("windows")]
        private void SetupGraph(IMoniker moniker, string deviceName, int width, int height, short bitCount, long avgTimePerFrame)
        {
            int hr;

            ISampleGrabber sampleGrabber = null;
            IBaseFilter captureFilter = null;
            //IPin pinCaptureOut = null;
            IPin pinSampleIn = null;
            //IPin pinRenderIn = null;

            // Get the graphbuilder object
            _filterGraph = new FilterGraph() as IFilterGraph2;

            try
            {
                // add the video input device
                hr = _filterGraph.AddSourceFilterForMoniker(moniker, null, deviceName, out captureFilter);
                DsError.ThrowExceptionForHR(hr);

                // Find the still pin
                _pinStill = DsFindPin.ByCategory(captureFilter, PinCategory.Still, 0);

                // Didn't find one.  Is there a preview pin?
                if (_pinStill == null)
                {
                    _pinStill = DsFindPin.ByCategory(captureFilter, PinCategory.Preview, 0);
                }

                // Still haven't found one.  Need to put a splitter in so we have
                // one stream to capture the bitmap from, and one to display.  Ok, we
                // don't *have* to do it that way, but we are going to anyway.
                if (_pinStill == null)
                {
                    IPin pinCapture = null;
                    IPin pinSmartTee = null;

                    // There is no still pin
                    _videoControl = null;

                    // Add a splitter
                    IBaseFilter filterSmartTee = (IBaseFilter)new SmartTee();

                    try
                    {
                        hr = _filterGraph.AddFilter(filterSmartTee, "SmartTee");
                        DsError.ThrowExceptionForHR(hr);

                        // Find the find the capture pin from the video device and the
                        // input pin for the splitter, and connnect them
                        pinCapture = DsFindPin.ByCategory(captureFilter, PinCategory.Capture, 0);
                        pinSmartTee = DsFindPin.ByDirection(filterSmartTee, PinDirection.Input, 0);

                        hr = _filterGraph.Connect(pinCapture, pinSmartTee);
                        DsError.ThrowExceptionForHR(hr);

                        // Now set the capture and still pins (from the splitter)
                        _pinStill = DsFindPin.ByName(filterSmartTee, "Capture");
                        //pinCaptureOut = DsFindPin.ByName(filterSmartTee, "Preview");

                        // If any of the default config items are set, perform the config
                        // on the actual video device (rather than the splitter)
                        if (height + width + bitCount > 0)
                        {
                            SetConfigParms(pinCapture, width, height, bitCount, avgTimePerFrame);
                        }
                    }
                    finally
                    {
                        if (pinCapture != null)
                        {
                            Marshal.ReleaseComObject(pinCapture);
                        }
                        if (pinCapture != pinSmartTee)
                        {
                            Marshal.ReleaseComObject(pinSmartTee);
                        }
                        if (pinCapture != filterSmartTee)
                        {
                            Marshal.ReleaseComObject(filterSmartTee);
                        }
                    }
                }
                else
                {
                    // Get a control pointer (used in Click())
                    _videoControl = captureFilter as IAMVideoControl;

                    // If any of the default config items are set
                    if (height + width + bitCount > 0)
                    {
                        SetConfigParms(_pinStill, width, height, bitCount, avgTimePerFrame);
                    }
                }

                // Get the SampleGrabber interface
                sampleGrabber = new SampleGrabber() as ISampleGrabber;

                // Configure the sample grabber
                IBaseFilter baseGrabFlt = sampleGrabber as IBaseFilter;
                ConfigureSampleGrabber(sampleGrabber);
                pinSampleIn = DsFindPin.ByDirection(baseGrabFlt, PinDirection.Input, 0);

                // Add the sample grabber to the graph
                hr = _filterGraph.AddFilter(baseGrabFlt, "Ds.NET Grabber");
                DsError.ThrowExceptionForHR(hr);

                if (_videoControl == null)
                {
                    // Connect the Still pin to the sample grabber
                    hr = _filterGraph.Connect(_pinStill, pinSampleIn);
                    DsError.ThrowExceptionForHR(hr);
                }
                else
                {
                    // Connect the Still pin to the sample grabber
                    hr = _filterGraph.Connect(_pinStill, pinSampleIn);
                    DsError.ThrowExceptionForHR(hr);
                }

                // Learn the video properties
                SaveSizeInfo(sampleGrabber);

                // Start the graph
                IMediaControl mediaCtrl = _filterGraph as IMediaControl;
                hr = mediaCtrl.Run();
                DsError.ThrowExceptionForHR(hr);
            }
            finally
            {
                if (sampleGrabber != null)
                {
                    Marshal.ReleaseComObject(sampleGrabber);
                    sampleGrabber = null;
                }
                if (pinSampleIn != null)
                {
                    Marshal.ReleaseComObject(pinSampleIn);
                    pinSampleIn = null;
                }
            }
        }

        private void SaveSizeInfo(ISampleGrabber sampGrabber)
        {
            int hr;

            // Get the media type from the SampleGrabber
            AMMediaType media = new AMMediaType();

            hr = sampGrabber.GetConnectedMediaType(media);
            DsError.ThrowExceptionForHR(hr);

            if ((media.formatType != FormatType.VideoInfo) || (media.formatPtr == IntPtr.Zero))
            {
                throw new NotSupportedException("Unknown Grabber Media Format");
            }

            // Grab the size info
            VideoInfoHeader videoInfoHeader = (VideoInfoHeader)Marshal.PtrToStructure(media.formatPtr, typeof(VideoInfoHeader));
            _width = videoInfoHeader.BmiHeader.Width;
            _height = videoInfoHeader.BmiHeader.Height;
            _stride = _width * (videoInfoHeader.BmiHeader.BitCount / 8);

            _ipBuffer = Marshal.AllocCoTaskMem(Math.Abs(_stride) * _height);

            DsUtils.FreeAMMediaType(media);
            media = null;
        }

        /// <summary>
        /// フィルタ内を通るサンプルをバッファにコピーする設定をおこなう。
        /// サンプル時のコールバック設定をおこなう。
        /// </summary>
        /// <param name="sampGrabber"></param>
        private void ConfigureSampleGrabber(ISampleGrabber sampGrabber)
        {
            int hr;
            AMMediaType media = new AMMediaType();

            // Set the media type to Video/RBG24
            media.majorType = MediaType.Video;
            media.subType = MediaSubType.RGB24;
            media.formatType = FormatType.VideoInfo;
            hr = sampGrabber.SetMediaType(media);
            DsError.ThrowExceptionForHR(hr);

            DsUtils.FreeAMMediaType(media);
            media = null;

            // Configure the samplegrabber
            hr = sampGrabber.SetCallback(this, 1);
            DsError.ThrowExceptionForHR(hr);
        }

        /// <summary>
        /// Set the Framerate, and video size.
        /// </summary>
        /// <param name="pinStill"></param>
        /// <param name="width"></param>
        /// <param name="height"></param>
        /// <param name="bitCount"></param>
        private void SetConfigParms(IPin pinStill, int width, int height, short bitCount, long avgTimePerFrame)
        {
            int hr;
            AMMediaType media;
            VideoInfoHeader v;

            IAMStreamConfig videoStreamConfig = pinStill as IAMStreamConfig;

            // Get the existing format block
            hr = videoStreamConfig.GetFormat(out media);
            DsError.ThrowExceptionForHR(hr);

            try
            {
                // copy out the videoinfoheader
                v = new VideoInfoHeader();
                Marshal.PtrToStructure(media.formatPtr, v);

                // if overriding the width, set the width
                if (width > 0)
                {
                    v.BmiHeader.Width = width;
                }

                // if overriding the Height, set the Height
                if (height > 0)
                {
                    v.BmiHeader.Height = height;
                }

                // if overriding the bits per pixel
                if (bitCount > 0)
                {
                    v.BmiHeader.BitCount = bitCount;
                }

                // if overriding the average time per frame
                if (avgTimePerFrame > 0)
                {
                    v.AvgTimePerFrame = avgTimePerFrame;
                }

                // Copy the media structure back
                Marshal.StructureToPtr(v, media.formatPtr, false);

                // Set the new format
                hr = videoStreamConfig.SetFormat(media);
                DsError.ThrowExceptionForHR(hr);
            }
            finally
            {
                DsUtils.FreeAMMediaType(media);
                media = null;
            }
        }

        /// <summary> Shut down capture </summary>
        [SupportedOSPlatform("windows")]
        private void CloseInterfaces()
        {
            int hr;

            try
            {
                if (_filterGraph != null)
                {
                    IMediaControl mediaCtrl = _filterGraph as IMediaControl;

                    // Stop the graph
                    hr = mediaCtrl.Stop();
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex);
            }

            if (_filterGraph != null)
            {
                Marshal.ReleaseComObject(_filterGraph);
                _filterGraph = null;
            }

            if (_videoControl != null)
            {
                Marshal.ReleaseComObject(_videoControl);
                _videoControl = null;
            }

            if (_pinStill != null)
            {
                Marshal.ReleaseComObject(_pinStill);
                _pinStill = null;
            }

            if (_ipBuffer != IntPtr.Zero)
            {
                Marshal.FreeCoTaskMem(_ipBuffer);
                _ipBuffer = IntPtr.Zero;
            }
        }
    }
}
