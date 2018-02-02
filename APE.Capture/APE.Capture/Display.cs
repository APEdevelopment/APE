//
//Copyright 2016 David Beales
//
//Licensed under the Apache License, Version 2.0 (the "License");
//you may not use this file except in compliance with the License.
//You may obtain a copy of the License at
//
//    http://www.apache.org/licenses/LICENSE-2.0
//
//Unless required by applicable law or agreed to in writing, software
//distributed under the License is distributed on an "AS IS" BASIS,
//WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
//See the License for the specific language governing permissions and
//limitations under the License.
//
using System;
using System.Diagnostics;
using System.Windows.Forms;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;
using System.Drawing.Imaging;
using APE.Native;
using NM = APE.Native.NativeMethods;
                                //http://nquant.codeplex.com/license
using nQuant;                   //Install-Package nQuant
using System.IO;
using System.Collections.Generic;

namespace APE.Capture
{
    /// <summary>
    /// Provides methods related to the display such as to capture screenshots and video
    /// </summary>
    public sealed class Display
    {
        private static bool m_IsCapturingVideo = false; // Only one video is allowed to be recored at a time...
        private static Thread VideoThread = null;
        private static string m_VideoFilename = "";
        private static IntPtr m_VideoCaptureWindow = IntPtr.Zero;
        private static string m_VideoCodec = "SCPR";
        private static int m_VideoFrameRate = 25;
        private static PixelFormat m_VideoPixelFormat = PixelFormat.Format16bppRgb555;
        private static int m_VideoQuality = 100;
        private static int m_VideoKeyFrameEvery = 500;
        private static PixelFormat m_ScreenPixelFormat = PixelFormat.Format8bppIndexed;

        /// <summary>
        /// Private constructor so no one can create a instance of this static class
        /// </summary>
        private Display()
        {
        }

        /// <summary>
        /// Returns a list of the four character codes for all video compression codecs installed
        /// </summary>
        public static List<string> Codecs
        {
            get
            {
                return VfW.Codecs();
            }
        }

        /// <summary>
        /// The pixel format to to use when capturing screenshots.  Default PixelFormat.Format8bppIndexed
        /// </summary>
        public static PixelFormat ScreenPixelFormat
        {
            set
            {
                m_ScreenPixelFormat = value;
            }
        }

        /// <summary>
        /// The handle of the window to capture video of, zero means the whole primary monitor.  Default 0
        /// </summary>
        public static IntPtr VideoCaptureWindow
        {
            set
            {
                CheckIfCapturingVideo();
                m_VideoCaptureWindow = value;
            }
        }

        /// <summary>
        /// The four character code of the codec to use for compression of video.  Default SCPR
        /// </summary>
        public static string VideoCodec
        {
            set
            {
                CheckIfCapturingVideo();
                m_VideoCodec = value;
            }
        }

        /// <summary>
        /// The frame rate to use to when capturing video.  Default 25
        /// </summary>
        public static int VideoFrameRate
        {
            set
            {
                CheckIfCapturingVideo();
                m_VideoFrameRate = value;
            }
        }

        /// <summary>
        /// The pixel format to to use when capturing video.  Default PixelFormat.Format16bppRgb555
        /// </summary>
        public static PixelFormat VideoPixelFormat
        {
            set
            {
                CheckIfCapturingVideo();
                m_VideoPixelFormat = value;
            }
        }

        /// <summary>
        /// The quality to use in the compression of the video.  Default 100
        /// </summary>
        public static int VideoQuality
        {
            set
            {
                CheckIfCapturingVideo();
                m_VideoQuality = value;
            }
        }

        /// <summary>
        /// How often a key fram should be written to the video.  Default 500
        /// </summary>
        public static int VideoKeyFrameEvery
        {
            set
            {
                CheckIfCapturingVideo();
                m_VideoKeyFrameEvery = value;
            }
        }

        /// <summary>
        /// Captures a screenshot of the whole desktop and returns it as a byte array
        /// </summary>
        /// <param name="format">The format to save the image as</param>
        /// <returns>A byte array containing the desktop image</returns>
        public static byte[] ScreenCapture(ImageFormat format)
        {
            Image bitmap = ScreenCapture(IntPtr.Zero);
            using (MemoryStream memoryStream = new MemoryStream())
            {
                bitmap.Save(memoryStream, format);
                return memoryStream.ToArray();
            }
        }

        /// <summary>
        /// Captures a screenshot of the whole desktop and saves it to the specified file
        /// </summary>
        /// <param name="fileName">The file name to save the image as</param>
        public static void ScreenCapture(string fileName)
        {
            ScreenCapture(fileName, IntPtr.Zero);
        }

        /// <summary>
        /// Gets the rectangle representing the window in device independant pixels
        /// </summary>
        /// <param name="window">The handle of the widown to to get the rectangle of</param>
        /// <returns>Rectangle representing the window</returns>
        public static NM.tagRect GetWindowRectangleDIP(IntPtr window)
        {
            int Adjustment = 1;
            int TitleBarRight = 0;
            string WindowState = "";
            NM.tagRect ControlRect;
            bool topLevelWindow = NM.IsTopLevelWindow(window);

            if (topLevelWindow)
            {
                // Get the titlebar rectangle
                NM.TITLEBARINFO CurrentTitleBarInfo = new NM.TITLEBARINFO();
                CurrentTitleBarInfo.cbSize = (uint)Marshal.SizeOf(CurrentTitleBarInfo);
                NM.GetTitleBarInfo(window, ref CurrentTitleBarInfo);
                TitleBarRight = CurrentTitleBarInfo.rcTitleBar.right;

                // Get the windows current state
                NM.WindowPlacement CurrentWindowPlacement = new NM.WindowPlacement();
                CurrentWindowPlacement.length = (uint)Marshal.SizeOf(CurrentWindowPlacement);
                NM.GetWindowPlacement(window, ref CurrentWindowPlacement);
                if (CurrentWindowPlacement.showCmd.ToString() == "ShowMaximized")
                {
                    WindowState = "Maximized";
                }
            }

            //TODO screen capture broke for not toplevel windows / non dwm toplevel with dpi?
            NM.DwmIsCompositionEnabled(out bool desktopWindowManagerEnabled);

            if (desktopWindowManagerEnabled && topLevelWindow)
            {
                NM.DwmGetWindowAttribute(window, NM.DWMWINDOWATTRIBUTE.DWMWA_EXTENDED_FRAME_BOUNDS, out ControlRect, Marshal.SizeOf(typeof(NM.tagRect)));

                if (topLevelWindow && WindowState == "Maximized")
                {
                    Adjustment += ControlRect.right - TitleBarRight;
                }

                ControlRect.left = ControlRect.left + Adjustment;
                ControlRect.top = ControlRect.top + Adjustment;
                ControlRect.right = ControlRect.right - Adjustment;
                ControlRect.bottom = ControlRect.bottom - Adjustment;
            }
            else
            {
                NM.GetWindowRect(window, out ControlRect);
                NM.tagRect clipBox = NM.GetClipBox(window);

                ControlRect.right = ControlRect.left + clipBox.right;
                ControlRect.bottom = ControlRect.top + clipBox.bottom;

                if (topLevelWindow && WindowState == "Maximized")
                {
                    Adjustment += ControlRect.right - TitleBarRight;
                }

                float ScreenScalingFactor;
                using (Graphics desktopGraphics = Graphics.FromHwnd(window))
                {
                    IntPtr desktopDeviceContext = desktopGraphics.GetHdc();
                    int LogicalScreenHeight = NM.GetDeviceCaps(desktopDeviceContext, NM.DeviceCap.VERTRES);
                    int PhysicalScreenHeight = NM.GetDeviceCaps(desktopDeviceContext, NM.DeviceCap.DESKTOPVERTRES);
                    desktopGraphics.ReleaseHdc();
                    ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
                }

                ControlRect.left = (int)(Math.Round((float)(ControlRect.left + Adjustment) * ScreenScalingFactor));
                ControlRect.top = (int)(Math.Round((float)(ControlRect.top + Adjustment) * ScreenScalingFactor));
                ControlRect.right = (int)(Math.Round((float)(ControlRect.right - Adjustment) * ScreenScalingFactor));
                ControlRect.bottom = (int)(Math.Round((float)(ControlRect.bottom - Adjustment) * ScreenScalingFactor));
            }

            return ControlRect;
        }

        /// <summary>
        /// Captures a screenshot of the specified window or whole screen and returns it as an image
        /// </summary>
        /// <param name="window">The handle of the window to to capture or 0 for the whole screen</param>
        /// <returns>An image representation of the window</returns>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static Image ScreenCapture(IntPtr window)
        {
            PixelFormat GrabFormat;
            Bitmap windowBitmap;

            switch (m_ScreenPixelFormat)
            {
                case PixelFormat.Format16bppRgb555:
                case PixelFormat.Format16bppRgb565:
                case PixelFormat.Format24bppRgb:
                case PixelFormat.Format32bppArgb:
                case PixelFormat.Format32bppPArgb:
                case PixelFormat.Format32bppRgb:
                    GrabFormat = m_ScreenPixelFormat;
                    break;
                case PixelFormat.Format8bppIndexed:
                    GrabFormat = PixelFormat.Format32bppArgb;
                    break;
                default:
                    throw new Exception("Format " + m_ScreenPixelFormat.ToString() + " is not supported");
            }

            int Width;
            int Height;
            int X;
            int Y;

            if (window == IntPtr.Zero)
            {
                Width = Screen.PrimaryScreen.Bounds.Width;
                Height = Screen.PrimaryScreen.Bounds.Height;
                X = 0;
                Y = 0;
            }
            else
            {
                NM.tagRect WindowRect;
                if (NativeVersion.IsWindowsVistaOrHigher && NM.IsTopLevelWindow(window))
                {
                    NM.DwmIsCompositionEnabled(out bool DWMEnabled);

                    if (DWMEnabled)
                    {
                        NM.DwmGetWindowAttribute(window, NM.DWMWINDOWATTRIBUTE.DWMWA_EXTENDED_FRAME_BOUNDS, out WindowRect, Marshal.SizeOf(typeof(NM.tagRect)));
                    }
                    else
                    {
                        NM.GetWindowRect(window, out WindowRect);
                    }
                }
                else
                {
                    NM.GetWindowRect(window, out WindowRect);
                }

                Width = WindowRect.right - WindowRect.left;
                Height = WindowRect.bottom - WindowRect.top;
                X = WindowRect.left;
                Y = WindowRect.top;
            }

            windowBitmap = new Bitmap(Width, Height, GrabFormat);
            GetWindowBitmap(window, X, Y, ref windowBitmap, true, false);

            if (m_ScreenPixelFormat == PixelFormat.Format8bppIndexed)
            {
                WuQuantizer quantizer = new WuQuantizer();
                Image quantized = quantizer.QuantizeImage(windowBitmap);
                return quantized;
            }
            else
            {
                return windowBitmap;
            }
        }

        /// <summary>
        /// Captures a screenshot of the specified window or whole screen and saves it to the specified file
        /// </summary>
        /// <param name="fileName">The file name to save the image as</param>
        /// <param name="window">The handle of the window to to capture or 0 for the whole screen</param>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static void ScreenCapture(string fileName, IntPtr window)
        {
            Image bitmap = ScreenCapture(window);
            bitmap.Save(fileName);
        }

        /// <summary>
        /// Stops capturing a video
        /// </summary>
        public static void StopVideoCapture()
        {
            m_IsCapturingVideo = false;
            VideoThread.Join();
        }

        /// <summary>
        /// Starts capturing a video and saving it as the specified filename
        /// </summary>
        /// <param name="videoFilename">The file name to save the video as</param>
        public static void StartVideoCapture(string videoFilename)
        {
            m_VideoFilename = videoFilename;
            VideoThread = new Thread(new ThreadStart(CaptureVideo));
            VideoThread.SetApartmentState(ApartmentState.STA);
            VideoThread.Start();
        }

        private static void CheckIfCapturingVideo()
        {
            if (m_IsCapturingVideo)
            {
                throw new Exception("Can not change the capture window while recording a video");
            }
        }

        private static void CaptureVideo()
        {
            //TO DO set the defaults in the regristry for codec and install the SCPR codec
            //SCLS = MSU Screen Capture Lossless Codec v1.2
            //SCPR = Infognition ScreenPressor

            if (m_IsCapturingVideo)
            {
                throw new Exception("Video capture already in progress");
            }

            if (m_VideoFilename == "")
            {
                throw new Exception("Need to set the video file name property");
            }

            int Width;
            int Height;
            int X;
            int Y;

            if (m_VideoCaptureWindow == IntPtr.Zero)
            {
                Width = Screen.PrimaryScreen.Bounds.Width;
                Height = Screen.PrimaryScreen.Bounds.Height;
                X = 0;
                Y = 0;
            }
            else
            {
                NM.tagRect WindowRect;
                if (NativeVersion.IsWindowsVistaOrHigher && NM.IsTopLevelWindow(m_VideoCaptureWindow))
                {
                    NM.DwmIsCompositionEnabled(out bool DWMEnabled);

                    if (DWMEnabled)
                    {
                        NM.DwmGetWindowAttribute(m_VideoCaptureWindow, NM.DWMWINDOWATTRIBUTE.DWMWA_EXTENDED_FRAME_BOUNDS, out WindowRect, Marshal.SizeOf(typeof(NM.tagRect)));
                    }
                    else
                    {
                        NM.GetWindowRect(m_VideoCaptureWindow, out WindowRect);
                    }
                }
                else
                {
                    NM.GetWindowRect(m_VideoCaptureWindow, out WindowRect);
                }

                Width = WindowRect.right - WindowRect.left;
                Height = WindowRect.bottom - WindowRect.top;
                X = WindowRect.left;
                Y = WindowRect.top;
            }

            VfW.Codec = m_VideoCodec;
            VfW.FrameRate = m_VideoFrameRate;
            VfW.Quality = m_VideoQuality;
            VfW.KeyFrameEvery = m_VideoKeyFrameEvery;

            int TimeToSleep = 0;
            int Sleep = (int)(((float)1 / (float)m_VideoFrameRate) * (float)1000);
            Stopwatch timer = new Stopwatch();

            m_IsCapturingVideo = true;

            Bitmap screen = VfW.Open(m_VideoFilename, Width, Height, m_VideoPixelFormat);
            try
            {
                while (m_IsCapturingVideo)
                {
                    timer.Restart();

                    // If TimeToSleep is 0 or larger then we don't need to drop a frame so capture a new bitmap
                    // otherwise if TimeToSleep is negative we need to drop a frame to preserve the frame rate 
                    // so add the previous frame again
                    if (TimeToSleep >= 0)
                    {
                        // This uses gdi so is relatively slow (as in maybe 30ms)
                        GetWindowBitmap(m_VideoCaptureWindow, X, Y, ref screen, true, false);
                        TimeToSleep = 0;
                    }
                    VfW.AddFrame(screen);

                    TimeToSleep = Sleep - (int)timer.ElapsedMilliseconds + TimeToSleep;

                    if (TimeToSleep > 0)
                    {
                        Thread.Sleep(TimeToSleep);
                    }
                }
            }
            finally
            {
                VfW.Close();
            }
        }

        /// <summary>
        /// Captures the specified window as bitmap
        /// </summary>
        /// <param name="handle">Handle of the window to capture</param>
        /// <param name="x">The left side of the window</param>
        /// <param name="y">The top of the window</param>
        /// <param name="windowBitmap">The bitmap to store the capture in</param>
        /// <param name="captureMouse">Whether to include the mouse cursor in the capture</param>
        /// <param name="clearClientArea">Whether to clear the client area portion of the window</param>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static void GetWindowBitmap(IntPtr handle, int x, int y, ref Bitmap windowBitmap, bool captureMouse, bool clearClientArea)
        {
            using (Graphics gdest = Graphics.FromImage(windowBitmap))
            {
                using (Graphics gsrc = Graphics.FromHwnd(IntPtr.Zero))
                {
                    try
                    {
                        IntPtr hSrcDC = gsrc.GetHdc();
                        IntPtr hDC = gdest.GetHdc();

                        // BitBlt is faster than CopyFromScreen
                        bool retval = NM.BitBlt(hDC, 0, 0, windowBitmap.Width, windowBitmap.Height, hSrcDC, x, y, CopyPixelOperation.SourceCopy | CopyPixelOperation.CaptureBlt);

                        if (clearClientArea)
                        {
                            NM.GetClientRect(handle, out NM.tagRect theRect);

                            NM.tagPoint thePoint = new NM.tagPoint();
                            NM.ClientToScreen(handle, ref thePoint);

                            NM.Rectangle(hDC, thePoint.x - x, thePoint.y - y, (thePoint.x - x) + theRect.right, (thePoint.y - y) + theRect.bottom);
                        }

                        if (captureMouse)
                        {
                            NM.tagCURSORINFO cursorInfo = new NM.tagCURSORINFO();
                            cursorInfo.cbSize = Marshal.SizeOf(typeof(NM.tagCURSORINFO));

                            if (NM.GetCursorInfo(ref cursorInfo))
                            {
                                if (cursorInfo.flags.HasFlag(NM.CursorFlags.Cursor_Showing))
                                {
                                    if (NM.GetIconInfo(cursorInfo.hCursor, out NM.ICONINFO iconInfo))
                                    {
                                        NM.DrawIcon(hDC, cursorInfo.ptScreenPos.x - iconInfo.xHotspot, cursorInfo.ptScreenPos.y - iconInfo.yHotspot, cursorInfo.hCursor);
                                    }
                                }
                            }
                        }
                    }
                    finally
                    {
                        // Clean up
                        gdest.ReleaseHdc();
                        gsrc.ReleaseHdc();
                    }
                }
            }
        }

        /// <summary>
        /// Compares two bitmaps
        /// </summary>
        /// <param name="b1">The first bitmap</param>
        /// <param name="b2">The second bitmap</param>
        /// <returns>True if they match otherwise false</returns>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public static bool CompareBitmap(Bitmap b1, Bitmap b2)
        {
            if ((b1 == null) != (b2 == null)) return false;
            if (b1.Size != b2.Size) return false;

            BitmapData bd1 = b1.LockBits(new Rectangle(new Point(0, 0), b1.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
            BitmapData bd2 = b2.LockBits(new Rectangle(new Point(0, 0), b2.Size), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

            try
            {
                IntPtr bd1scan0 = bd1.Scan0;
                IntPtr bd2scan0 = bd2.Scan0;

                int stride = bd1.Stride;
                IntPtr len = (IntPtr)(stride * b1.Height);
                IntPtr actual = NM.RtlCompareMemory(bd1scan0, bd2scan0, len);

                return actual == len;
            }
            finally
            {
                b1.UnlockBits(bd1);
                b2.UnlockBits(bd2);
            }
        }
    }
}
