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
using System.Runtime.InteropServices;
using System.Drawing;
using System.Text;
using System.Diagnostics;
using System.Security;
using System.Threading;
using Accessibility;

namespace APE.Native
{
    //TODO update all DllImport lines to include ,CharSet=CharSet.Unicode

    public static class NativeMethods
    {
        // There are some badly behaved controls (TabStrip in MsComCtl.ocx for instance) which crash if we set this to 1
        public const int IntraClickDelay = 2;

        public const int EM_GETSEL = 0x00B0;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool SetProcessDPIAware();

        [DllImport("kernel32.dll", EntryPoint = "OpenProcess", SetLastError = true)]
        public static extern IntPtr OpenProcess(ProcessAccessFlags dwDesiredAccess, [MarshalAs(UnmanagedType.Bool)] bool bInheritHandle, int dwProcessId);

        private const int WM_SETREDRAW = 11;

        public static void SuspendDrawing(IntPtr form, uint timeout)
        {
            IntPtr messageResult;
            IntPtr sendResult;
            sendResult = SendMessageTimeout(form, WM_SETREDRAW, new IntPtr(0), IntPtr.Zero, SendMessageTimeoutFlags.SMTO_NORMAL, timeout, out messageResult);
            if (sendResult == IntPtr.Zero) //Failed
            {
                throw new Exception("Failed to access the suspend drawing");
            }
        }

        public static void ResumeDrawing(IntPtr form, uint timeout)
        {
            IntPtr messageResult;
            IntPtr sendResult;
            sendResult = SendMessageTimeout(form, WM_SETREDRAW, new IntPtr(1), IntPtr.Zero, SendMessageTimeoutFlags.SMTO_NORMAL, timeout, out messageResult);
            if (sendResult == IntPtr.Zero) //Failed
            {
                throw new Exception("Failed to access the suspend drawing");
            }
        }

        [Flags]
        public enum ProcessAccessFlags : uint
        {
            Terminate = 0x0001,
            CreateThread = 0x0002,
            SetSessionID = 0x0004,
            VMOperation = 0x0008,
            VMRead = 0x0010,
            VMWrite = 0x0020,
            DUPHandle = 0x0040,
            CreateProcess = 0x0080,
            SetQuota = 0x0100,
            SetInformation = 0x0200,
            QueryInformation = 0x0400,
            SuspendResume = 0x0800,
            QueryLimitedInformation = 0x1000,
            AllAccess = Synchronize | StandardRightsRequired | 0xFFFF,
            Synchronize = 0x100000,
            StandardRightsRequired = 0x0F0000,
        }

        [DllImport("kernel32.dll", EntryPoint = "OpenThread", SetLastError = true)]
        public static extern IntPtr OpenThread(ThreadAccess dwDesiredAccess, bool bInheritHandle, uint dwThreadId);

        [Flags]
        public enum ThreadAccess : int
        {
            TERMINATE = (0x0001),
            SUSPEND_RESUME = (0x0002),
            GET_CONTEXT = (0x0008),
            SET_CONTEXT = (0x0010),
            SET_INFORMATION = (0x0020),
            QUERY_INFORMATION = (0x0040),
            SET_THREAD_TOKEN = (0x0080),
            IMPERSONATE = (0x0100),
            DIRECT_IMPERSONATION = (0x0200)
        }

        [DllImport("kernel32.dll", EntryPoint = "CreateRemoteThread", SetLastError = true)]
        public static extern IntPtr CreateRemoteThread(IntPtr hProcess, IntPtr lpThreadAttributes, uint dwStackSize, IntPtr lpStartAddress, IntPtr lpParameter, uint dwCreationFlags, [Out] IntPtr lpThreadId);

        [DllImport("kernel32.dll", EntryPoint = "GetModuleHandleW", SetLastError = true)]
        public static extern IntPtr GetModuleHandle([MarshalAs(UnmanagedType.LPWStr)] string lpModuleName);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int LoadString(IntPtr hInstance, uint uID, StringBuilder lpBuffer, int nBufferMax);

        [DllImport("kernel32.dll", EntryPoint = "LoadLibrary", SetLastError = true)]
        public static extern IntPtr LoadLibrary(string lpFileName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool FreeLibrary(IntPtr hModule);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr GetProcAddress(IntPtr hModule, string procName);

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool WriteProcessMemory(IntPtr hProcess, IntPtr lpBaseAddress, IntPtr lpBuffer, uint nSize, out int lpNumberOfBytesWritten);

        [DllImport("kernel32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetExitCodeThread(IntPtr hThread, out uint lpExitCode);

        public const uint STILL_ACTIVE = 259;

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        public static extern IntPtr VirtualAllocEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, AllocationType flAllocationType, MemoryProtection flProtect);

        [Flags]
        public enum AllocationType : uint
        {
            Commit = 0x1000,
            Reserve = 0x2000,
            Decommit = 0x4000,
            Release = 0x8000,
            Free = 0x10000,
            Private = 0x20000,
            Mapped = 0x40000,
            Reset = 0x80000,
            TopDown = 0x100000,
            WriteWatch = 0x200000,
            Physical = 0x400000,
            Rotate = 0x800000,
            LargePages = 0x20000000,
            FourMbPages = 0x80000000,
        }

        public enum MemoryProtection : uint
        {
            NoAccess = 0x001,
            ReadOnly = 0x002,
            ReadWrite = 0x004,
            WriteCopy = 0x008,
            Execute = 0x010,
            ExecuteRead = 0x020,
            ExecuteReadWrite = 0x040,
            ExecuteWriteCopy = 0x080,
            PageGuard = 0x100,
            NoCache = 0x200,
            WriteCombine = 0x400,
        }

        [DllImport("kernel32.dll", SetLastError = true, ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool VirtualFreeEx(IntPtr hProcess, IntPtr lpAddress, uint dwSize, AllocationType dwFreeType);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WindowPlacement lpwndpl);

        [StructLayout(LayoutKind.Sequential)]
        public struct WindowPlacement
        {
            public uint length;
            public WindowPlacementFlags flags;
            public ShowWindowEnum showCmd;
            public tagPoint ptMinPosition;
            public tagPoint ptMaxPosition;
            public tagRect rcNormalPosition;
        }

        [Flags]
        public enum WindowPlacementFlags : uint
        {
            SetMinPosition = 0x0001,
            RestoreToMaximized = 0x0002,
            AsyncWindowPlacement = 0x0004,
        }

        public enum ShowWindowEnum : uint
        {
            Hide = 0,
            ShowNormal = 1,
            ShowMinimized = 2,
            ShowMaximized = 3,
            Maximize = 3,
            ShowNoActivate = 4,
            Show = 5,
            Minimize = 6,
            ShowMinNoActive = 7,
            ShowNA = 8,
            Restore = 9,
        }

        [DllImport("gdi32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetDeviceCaps(IntPtr hdc, DeviceCap nIndex);

        public enum DeviceCap : int
        {
            /// <summary>
            /// Device driver version
            /// </summary>
            DRIVERVERSION = 0,
            /// <summary>
            /// Device classification
            /// </summary>
            TECHNOLOGY = 2,
            /// <summary>
            /// Horizontal size in millimeters
            /// </summary>
            HORZSIZE = 4,
            /// <summary>
            /// Vertical size in millimeters
            /// </summary>
            VERTSIZE = 6,
            /// <summary>
            /// Horizontal width in pixels
            /// </summary>
            HORZRES = 8,
            /// <summary>
            /// Vertical height in pixels
            /// </summary>
            VERTRES = 10,
            /// <summary>
            /// Number of bits per pixel
            /// </summary>
            BITSPIXEL = 12,
            /// <summary>
            /// Number of planes
            /// </summary>
            PLANES = 14,
            /// <summary>
            /// Number of brushes the device has
            /// </summary>
            NUMBRUSHES = 16,
            /// <summary>
            /// Number of pens the device has
            /// </summary>
            NUMPENS = 18,
            /// <summary>
            /// Number of markers the device has
            /// </summary>
            NUMMARKERS = 20,
            /// <summary>
            /// Number of fonts the device has
            /// </summary>
            NUMFONTS = 22,
            /// <summary>
            /// Number of colors the device supports
            /// </summary>
            NUMCOLORS = 24,
            /// <summary>
            /// Size required for device descriptor
            /// </summary>
            PDEVICESIZE = 26,
            /// <summary>
            /// Curve capabilities
            /// </summary>
            CURVECAPS = 28,
            /// <summary>
            /// Line capabilities
            /// </summary>
            LINECAPS = 30,
            /// <summary>
            /// Polygonal capabilities
            /// </summary>
            POLYGONALCAPS = 32,
            /// <summary>
            /// Text capabilities
            /// </summary>
            TEXTCAPS = 34,
            /// <summary>
            /// Clipping capabilities
            /// </summary>
            CLIPCAPS = 36,
            /// <summary>
            /// Bitblt capabilities
            /// </summary>
            RASTERCAPS = 38,
            /// <summary>
            /// Length of the X leg
            /// </summary>
            ASPECTX = 40,
            /// <summary>
            /// Length of the Y leg
            /// </summary>
            ASPECTY = 42,
            /// <summary>
            /// Length of the hypotenuse
            /// </summary>
            ASPECTXY = 44,
            /// <summary>
            /// Shading and Blending caps
            /// </summary>
            SHADEBLENDCAPS = 45,

            /// <summary>
            /// Logical pixels inch in X
            /// </summary>
            LOGPIXELSX = 88,
            /// <summary>
            /// Logical pixels inch in Y
            /// </summary>
            LOGPIXELSY = 90,

            /// <summary>
            /// Number of entries in physical palette
            /// </summary>
            SIZEPALETTE = 104,
            /// <summary>
            /// Number of reserved entries in palette
            /// </summary>
            NUMRESERVED = 106,
            /// <summary>
            /// Actual color resolution
            /// </summary>
            COLORRES = 108,

            // Printing related DeviceCaps. These replace the appropriate Escapes
            /// <summary>
            /// Physical Width in device units
            /// </summary>
            PHYSICALWIDTH = 110,
            /// <summary>
            /// Physical Height in device units
            /// </summary>
            PHYSICALHEIGHT = 111,
            /// <summary>
            /// Physical Printable Area x margin
            /// </summary>
            PHYSICALOFFSETX = 112,
            /// <summary>
            /// Physical Printable Area y margin
            /// </summary>
            PHYSICALOFFSETY = 113,
            /// <summary>
            /// Scaling factor x
            /// </summary>
            SCALINGFACTORX = 114,
            /// <summary>
            /// Scaling factor y
            /// </summary>
            SCALINGFACTORY = 115,

            /// <summary>
            /// Current vertical refresh rate of the display device (for displays only) in Hz
            /// </summary>
            VREFRESH = 116,
            /// <summary>
            /// Vertical height of entire desktop in pixels
            /// </summary>
            DESKTOPVERTRES = 117,
            /// <summary>
            /// Horizontal width of entire desktop in pixels
            /// </summary>
            DESKTOPHORZRES = 118,
            /// <summary>
            /// Preferred blt alignment
            /// </summary>
            BLTALIGNMENT = 119
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct DATETIMEPICKERINFO
        {
            public uint cbSize;
            public tagRect rcCheck;
            public uint stateCheck;
            public tagRect rcButton;
            public uint stateButton;
            public IntPtr hwndEdit;
            public IntPtr hwndUD;
            public IntPtr hwndDropDown;
        }

        const int DTM_GETDATETIMEPICKERINFO = 0x1000 + 14;

        public static DATETIMEPICKERINFO GetDateTimePickerInfo(IntPtr hWnd)
        {
            DATETIMEPICKERINFO info = new DATETIMEPICKERINFO();
            info.cbSize  = (uint)Marshal.SizeOf(typeof(DATETIMEPICKERINFO));
            info.rcButton.top = -1;
            info.rcButton.left = -1;
            info.rcButton.bottom = -1;
            info.rcButton.right = -1;
            info.rcCheck.top = -1;
            info.rcCheck.left = -1;
            info.rcCheck.bottom = -1;
            info.rcCheck.right = -1;
            IntPtr theResult = IntPtr.Zero;
            IntPtr theReturn = SendMessageTimeout(hWnd, DTM_GETDATETIMEPICKERINFO, IntPtr.Zero, ref info, SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 30, out theResult);
            if (theReturn == IntPtr.Zero)
            {
                throw new Exception("SendMessageTimeout failed, does the window exists?");
            }
            
            if (info.rcCheck.top == -1 && info.rcCheck.left == -1 && info.rcCheck.bottom == -1 && info.rcCheck.right == -1)
            {
                // probably 5.82 version of the common controls which doesn't support this message so guess the coordinates
                info.rcCheck.top = 2;
                info.rcCheck.left = 2;
                info.rcCheck.bottom = 18;
                info.rcCheck.right = 18;
            }

            if (info.rcButton.top == -1 && info.rcButton.left == -1 && info.rcButton.bottom == -1 && info.rcButton.right == -1)
            {
                // probably 5.82 version of the common controls which doesn't support this message so guess the coordinates
                tagRect rect = GetClipBox(hWnd);
                info.rcButton.top = 2;
                info.rcButton.left = rect.right - 19;
                info.rcButton.bottom = 18;
                info.rcButton.right = rect.right - 2;
            }
            return info;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetTitleBarInfo(IntPtr hwnd, ref TITLEBARINFO pti);

        [StructLayout(LayoutKind.Sequential)]
        public struct TITLEBARINFO
        {
            public const int CCHILDREN_TITLEBAR = 5;
            public uint cbSize;
            public tagRect rcTitleBar;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = CCHILDREN_TITLEBAR + 1)]
            public uint[] rgstate;
        }

        [DllImport("user32.dll", EntryPoint = "GetWindowLong")]
        private static extern IntPtr GetWindowLongPtr32(IntPtr hWnd, GWL nIndex);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr")]
        private static extern IntPtr GetWindowLongPtr64(IntPtr hWnd, GWL nIndex);

        public static IntPtr GetWindowLongPtr(IntPtr hWnd, GWL nIndex)
        {
            if (IntPtr.Size == 8)
            {
                return GetWindowLongPtr64(hWnd, nIndex);
            }
            else if (IntPtr.Size == 4)
            {
                return GetWindowLongPtr32(hWnd, nIndex);
            }
            else
            {
                throw new Exception("Unsupported pointer size");
            }
        }

        [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
        private static extern IntPtr SetWindowLongPtr32(IntPtr hWnd, GWL nIndex, IntPtr dwNewLong);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
        private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, GWL nIndex, IntPtr dwNewLong);

        public static IntPtr SetWindowLongPtr(IntPtr hWnd, GWL nIndex, long dwNewLong)
        {
            if (IntPtr.Size == 8)
            {
                return SetWindowLongPtr64(hWnd, nIndex, new IntPtr(dwNewLong));
            }
            else if (IntPtr.Size == 4)
            {
                return SetWindowLongPtr32(hWnd, nIndex, new IntPtr((int)dwNewLong));
            }
            else
            {
                throw new Exception("Unsupported pointer size");
            }
        }

        [Flags]
        public enum GWL : int
        {
            GWL_WNDPROC = (-4),
            GWL_HINSTANCE = (-6),
            GWL_HWNDPARENT = (-8),
            GWL_STYLE = (-16),
            GWL_EXSTYLE = (-20),
            GWL_USERDATA = (-21),
            GWL_ID = (-12)
        }

        /// Return Type: BOOL->int
        ///hdc: HDC->HDC__*
        ///left: int
        ///top: int
        ///right: int
        ///bottom: int
        [DllImportAttribute("gdi32.dll", EntryPoint = "Rectangle")]
        [return: MarshalAsAttribute(UnmanagedType.Bool)]
        public static extern bool Rectangle([In] IntPtr hdc, int left, int top, int right, int bottom);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///lpPoint: LPPOINT->tagPOINT*
        [DllImport("user32.dll", EntryPoint = "ClientToScreen")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ClientToScreen([In] IntPtr hWnd, ref tagPoint lpPoint);

        [StructLayoutAttribute(LayoutKind.Sequential)]
        public struct tagPoint
        {
            /// LONG->int
            public int x;

            /// LONG->int
            public int y;
        }

        /// Return Type: HRESULT->Long->int
        ///pfEnabled: Bool
        [DllImport("dwmapi.dll", EntryPoint = "DwmIsCompositionEnabled", CharSet = CharSet.Unicode)]
        public static extern int DwmIsCompositionEnabled([Out] [MarshalAs(UnmanagedType.Bool)] out bool pfEnabled);

        /// Return Type: HRESULT->Long->int
        ///hwnd: HANDLE->void*
        ///dwAttribute: DWORD->unsigned int->int
        ///pvAttribute: LPRECT->tagRECT*
        ///cbAttribute: DWORD->unsigned int->int
        [DllImport("dwmapi.dll", EntryPoint = "DwmGetWindowAttribute", CharSet = CharSet.Unicode)]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE dwAttribute, [Out] out tagRect pvAttribute, int cbAttribute);

        /// Return Type: HRESULT->Long->int
        ///hwnd: HANDLE->void*
        ///dwAttribute: DWORD->unsigned int->int
        ///pvAttribute: Bool
        ///cbAttribute: DWORD->unsigned int->int
        [DllImport("dwmapi.dll", EntryPoint = "DwmGetWindowAttribute", CharSet = CharSet.Unicode)]
        public static extern int DwmGetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE dwAttribute, [Out] [MarshalAs(UnmanagedType.Bool)] out bool pvAttribute, int cbAttribute);

        /// Return Type: BOOL->int
        ///pci: PCURSORINFO->tagCURSORINFO*
        [DllImport("user32.dll", EntryPoint = "GetCursorInfo", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorInfo([In, Out] ref tagCURSORINFO pci);

        /// Return Type: BOOL->int
        ///hDC: HDC->HDC__*->IntPtr
        ///X: int
        ///Y: int
        ///hIcon: HICON->HICON__*->IntPtr
        [DllImport("user32.dll", EntryPoint = "DrawIcon", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DrawIcon([In] IntPtr hDC, int X, int Y, [In] IntPtr hIcon);

        /// Return Type: BOOL->int
        ///hIcon: HICON->HICON__*->IntPtr
        ///piconinfo: PICONINFO->ICONINFO*
        [DllImport("user32.dll", EntryPoint = "GetIconInfo", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetIconInfo([In] IntPtr hIcon, [Out] out ICONINFO piconinfo);

        /// Return Type: SIZE_T->ULONG_PTR->unsigned intptr->IntPtr
        ///Source1: void*
        ///Source2: void*
        ///Length: SIZE_T->ULONG_PTR->unsigned intptr->IntPtr
        [DllImport("ntdll.dll", EntryPoint = "RtlCompareMemory", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr RtlCompareMemory(IntPtr Source1, IntPtr Source2, IntPtr Length);

        /// Return Type: BOOL->int
        ///hdc: HDC->HDC__*->IntPtr
        ///x: int
        ///y: int
        ///cx: int
        ///cy: int
        ///hdcSrc: HDC->HDC__*->IntPtr
        ///x1: int
        ///y1: int
        ///rop: DWORD->unsigned int
        [DllImport("gdi32.dll", EntryPoint = "BitBlt", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool BitBlt([In] IntPtr hdc, int x, int y, int cx, int cy, [In] IntPtr hdcSrc, int x1, int y1, CopyPixelOperation rop);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*->IntPtr
        ///lpRect: LPRECT->tagRECT*
        [DllImport("user32.dll", EntryPoint = "GetWindowRect", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect([In] IntPtr hWnd, [Out] out tagRect lpRect);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*->IntPtr
        ///lpRect: LPRECT->tagRECT*
        [DllImport("user32.dll", EntryPoint = "GetClientRect", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetClientRect([In] IntPtr hWnd, [Out] out tagRect lpRect);

        /// Return Type: LPVOID->void*
        ///hFileMappingObject: HANDLE->void*
        ///dwDesiredAccess: DWORD->unsigned int
        ///dwFileOffsetHigh: DWORD->unsigned int
        ///dwFileOffsetLow: DWORD->unsigned int
        ///dwNumberOfBytesToMap: SIZE_T->ULONG_PTR->unsigned intptr 
        [DllImport("kernel32.dll", EntryPoint = "MapViewOfFile", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr MapViewOfFile([In] IntPtr hFileMappingObject, FileMapAccess dwDesiredAccess, uint dwFileOffsetHigh, uint dwFileOffsetLow, IntPtr dwNumberOfBytesToMap);

        /// Return Type: BOOL->int
        ///lpBaseAddress: LPCVOID->void*
        [DllImport("kernel32.dll", EntryPoint = "UnmapViewOfFile", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnmapViewOfFile([In] IntPtr lpBaseAddress);

        /// Return Type: HANDLE->void*
        ///lpEventAttributes: LPSECURITY_ATTRIBUTES->_SECURITY_ATTRIBUTES*->IntPtr
        ///bManualReset: BOOL->int
        ///bInitialState: BOOL->int
        ///lpName: LPCWSTR->WCHAR*
        [DllImport("kernel32.dll", EntryPoint = "CreateEvent", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateEvent([In] IntPtr lpEventAttributes, [MarshalAs(UnmanagedType.Bool)] bool bManualReset, [MarshalAs(UnmanagedType.Bool)] bool bInitialState, [In] [MarshalAs(UnmanagedType.LPWStr)] string lpName);

        /// Return Type: BOOL->int
        ///hEvent: HANDLE->void*
        [DllImport("kernel32.dll", EntryPoint = "PulseEvent", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PulseEvent([In] IntPtr hEvent);

        /// Return Type: BOOL->int
        ///hEvent: HANDLE->void*
        [DllImport("kernel32.dll", EntryPoint = "SetEvent", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetEvent([In] IntPtr hEvent);

        /// Return Type: HANDLE->void*
        ///hFile: HANDLE->void*
        ///lpFileMappingAttributes: LPSECURITY_ATTRIBUTES->_SECURITY_ATTRIBUTES*->Intptr
        ///flProtect: DWORD->unsigned int
        ///dwMaximumSizeHigh: DWORD->unsigned int
        ///dwMaximumSizeLow: DWORD->unsigned int
        ///lpName: LPCWSTR->WCHAR*
        [DllImport("kernel32.dll", EntryPoint = "CreateFileMapping", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr CreateFileMapping([In] IntPtr hFile, [In] IntPtr lpFileMappingAttributes, PageProtection flProtect, uint dwMaximumSizeHigh, uint dwMaximumSizeLow, [In] [MarshalAs(UnmanagedType.LPWStr)] string lpName);

        /// Return Type: BOOL->int
        ///hObject: HANDLE->void*
        [DllImport("kernel32.dll", EntryPoint = "CloseHandle", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAsAttribute(UnmanagedType.Bool)]
        public static extern bool CloseHandle([In] IntPtr hObject);

        [DllImport("kernel32.dll", EntryPoint = "LoadLibraryExW", SetLastError = true)]
        public static extern IntPtr LoadLibraryEx([MarshalAs(UnmanagedType.LPWStr)] string lpFileName, IntPtr hFile, LoadLibraryExFlags dwFlags);

        [Flags]
        public enum LoadLibraryExFlags : uint
        {
            DontResolveDllReferences = 0x00000001,
            LoadLibraryAsDatafile = 0x00000002,
            LoadLibraryWithAlteredSearchPath = 0x00000008,
            LoadIgnoreCodeAuthzLevel = 0x00000010,
            LoadLibraryAsImageResource = 0x00000020,
            LoadLibraryAsDatafileExclusive = 0x00000040,
        }

        /// Return Type: DWORD->unsigned int->int
        ///hHandle: HANDLE->void*
        ///dwMilliseconds: DWORD->unsigned int
        [DllImport("kernel32.dll", EntryPoint = "WaitForSingleObject", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern ThreadWaitValue WaitForSingleObject([In] IntPtr hHandle, int dwMilliseconds);

        public enum ThreadWaitValue : uint
        {
            Object0 = 0x00000000,
            Abandoned = 0x00000080,
            Timeout = 0x00000102,
            Failed = 0xFFFFFFFF,
            Infinite = 0xFFFFFFFF
        }

        /// Return Type: void*
        ///_Dst: void*
        ///_Src: void*
        ///_Size: size_t->unsigned intptr->IntPtr
        [DllImport("ntdll.dll", EntryPoint = "memcpy", CallingConvention = System.Runtime.InteropServices.CallingConvention.Cdecl)]
        public static extern IntPtr memcpy(IntPtr _Dst, [In] IntPtr _Src, IntPtr _Size);

        [DllImport("msvfw32.dll", EntryPoint = "ICInfo", CharSet = CharSet.Unicode)]
        public static extern bool ICInfo(int fccType, int fccHandler, ref ICINFO lpicinfo);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct ICINFO
        {
            public int dwSize;
            public int fccType;
            public int fccHandler;
            public int dwFlags;
            public int dwVersion;
            public int dwVersionICM;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 16)]
            public string szName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szDescription;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            public string szDriver;
        }

        [DllImport("avifil32.dll", EntryPoint = "AVIFileInit", CharSet = CharSet.Unicode)]
        public static extern void AVIFileInit();

        [DllImport("avifil32.dll", EntryPoint = "AVIFileExit", CharSet = CharSet.Unicode)]
        public static extern void AVIFileExit();

        /// Return Type: STDAPI->HRESULT->int
        ///ppfile: PAVIFILE->IntPtr
        ///szFile: LPCTSTR->WCHAR*
        ///mode: uint
        ///pclsidHandler: CLSID->GUID
        [DllImport("avifil32.dll", EntryPoint = "AVIFileOpen", CharSet = CharSet.Unicode)]
        public static extern int AVIFileOpen(out IntPtr ppfile, [MarshalAs(UnmanagedType.LPWStr)] string szFile, OpenFileMode mode, IntPtr pclsidHandler);

        /// Return Type: STDAPI->HRESULT->int
        ///pfile: PAVIFILE->IntPtr
        [DllImport("avifil32.dll", EntryPoint = "AVIFileRelease", CharSet = CharSet.Unicode)]
        public static extern int AVIFileRelease(IntPtr pfile);

        /// Return Type: STDAPI->HRESULT->int
        ///pfile: PAVIFILE->IntPtr
        ///ppavi: PAVISTREAM->IntPtr
        ///psi: AVISTREAMINFO
        [DllImport("avifil32.dll", EntryPoint = "AVIFileCreateStream", CharSet = CharSet.Unicode)]
        public static extern int AVIFileCreateStream(IntPtr pfile, out IntPtr ppavi, ref AVISTREAMINFO psi);

        /// Return Type: STDAPI->HRESULT->int
        ///ppavi: PAVISTREAM->IntPtr
        [DllImport("avifil32.dll", EntryPoint = "AVIStreamRelease", CharSet = CharSet.Unicode)]
        public static extern int AVIStreamRelease(IntPtr pavi);

        /// Return Type: STDAPI->HRESULT->int
        ///ppavi: PAVISTREAM->IntPtr
        ///lPos: LONG->int
        ///lpFormat: BITMAPINFOHEADER
        ///cbFormat: LONG->int
        [DllImport("avifil32.dll", EntryPoint = "AVIStreamSetFormat", CharSet = CharSet.Unicode)]
        public static extern int AVIStreamSetFormat(IntPtr pavi, int lPos, ref BITMAPINFOHEADER lpFormat, int cbFormat);

        /// Return Type: STDAPI->HRESULT->int
        ///ppavi: PAVISTREAM->IntPtr
        ///lStart: LONG->int
        ///lSamples: LONG->int
        ///lpBuffer: VOID*
        ///cbBuffer: LONG->int
        ///dwFlags: DWORD->unsigned int->int
        ///plSampWritten: VOID*
        ///plBytesWritten: VOID*
        [DllImport("avifil32.dll", EntryPoint = "AVIStreamWrite", CharSet = CharSet.Unicode)]
        public static extern int AVIStreamWrite(IntPtr pavi, int lStart, int lSamples, IntPtr lpBuffer, int cbBuffer, int dwFlags, IntPtr plSampWritten, IntPtr plBytesWritten);

        /// Return Type: STDAPI->HRESULT->int
        ///ppsCompressed: VOID*
        ///psSource: VOID*
        ///lpOptions: AVICOMPRESSOPTIONS
        ///pclsidHandler: VOID*
        [DllImport("avifil32.dll", EntryPoint = "AVIMakeCompressedStream", CharSet = CharSet.Unicode)]
        public static extern int AVIMakeCompressedStream(out IntPtr ppsCompressed, IntPtr psSource, ref AVICOMPRESSOPTIONS lpOptions, IntPtr pclsidHandler);

        [StructLayout(LayoutKind.Sequential)]
        public struct tagRect
        {
            /// LONG->int
            public int left;

            /// LONG->int
            public int top;

            /// LONG->int
            public int right;

            /// LONG->int
            public int bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ICONINFO
        {
            /// BOOL->int
            [MarshalAs(UnmanagedType.Bool)]
            public bool fIcon;

            /// DWORD->unsigned int->int
            public int xHotspot;

            /// DWORD->unsigned int->int
            public int yHotspot;

            /// HBITMAP->HBITMAP__*
            public IntPtr hbmMask;

            /// HBITMAP->HBITMAP__*
            public IntPtr hbmColor;
        }

        [StructLayoutAttribute(LayoutKind.Sequential)]
        public struct tagCURSORINFO
        {
            /// DWORD->unsigned int->int
            public int cbSize;

            /// DWORD->unsigned int
            public CursorFlags flags;

            /// HCURSOR->HICON->HICON__*
            public IntPtr hCursor;

            /// POINT->tagPOINT
            public tagPoint ptScreenPos;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Point
        {
            /// LONG->int
            public int x;

            /// LONG->int
            public int y;
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode, Pack = 1)]
        public struct AVISTREAMINFO
        {
            /// LONG->int
            public int fccType;

            /// LONG->int
            public int fccHandler;

            /// DWORD->unsigned int->int
            public int dwFlags;

            /// DWORD->unsigned int->int
            public int dwCaps;

            /// WORD->unsigned short->short
            public short wPriority;

            /// WORD->unsigned short->short
            public short wLanguage;

            /// DWORD->unsigned int->int
            public int dwScale;

            /// DWORD->unsigned int->int
            public int dwRate;

            /// DWORD->unsigned int->int
            public int dwStart;

            /// DWORD->unsigned int->int
            public int dwLength;

            /// DWORD->unsigned int->int
            public int dwInitialFrames;

            /// DWORD->unsigned int->int
            public int dwSuggestedBufferSize;

            /// DWORD->unsigned int->int
            public int dwQuality;

            /// DWORD->unsigned int->int
            public int dwSampleSize;

            /// RECT->tagRECT
            [MarshalAs(UnmanagedType.Struct, SizeConst = 16)]
            public tagRect rcFrame;

            /// DWORD->unsigned int->int
            public int dwEditCount;

            /// DWORD->unsigned int->int
            public int dwFormatChangeCount;

            /// TCHAR 
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 64)]
            public string szName;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct BITMAPINFOHEADER
        {
            /// DWORD->unsigned int->int
            public int biSize;

            /// LONG->int
            public int biWidth;

            /// LONG->int
            public int biHeight;

            /// WORD->unsigned short->short
            public short biPlanes;

            /// WORD->unsigned short->short
            public short biBitCount;

            /// DWORD->unsigned int->int
            public int biCompression;

            /// DWORD->unsigned int->int
            public int biSizeImage;

            /// LONG->int
            public int biXPelsPerMeter;

            /// LONG->int
            public int biYPelsPerMeter;

            /// DWORD->unsigned int->int
            public int biClrUsed;

            /// DWORD->unsigned int->int
            public int biClrImportant;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        public struct AVICOMPRESSOPTIONS
        {
            /// DWORD->unsigned int->int
            public int fccType;

            /// DWORD->unsigned int->int
            public int fccHandler;

            /// DWORD->unsigned int->int
            public int dwKeyFrameEvery;

            /// DWORD->unsigned int->int
            public int dwQuality;

            /// DWORD->unsigned int->int
            public int dwBytesPerSecond;

            /// DWORD->unsigned int
            public AviCompression dwFlags;


            /// LPVOID->IntPtr
            public IntPtr lpFormat;

            /// DWORD->unsigned int->int
            public int cbFormat;

            /// LPVOID->IntPtr
            public int lpParms;

            /// DWORD->unsigned int->int
            public int cbParms;

            /// DWORD->unsigned int->int
            public int dwInterleaveEvery;
        }

        /// WAIT_OBJECT_0 -> 0 
        public const int WAIT_OBJECT_0 = 0;

        /// 
        [Flags]
        public enum WindowStyles : long
        {
            WS_EX_LEFT = 0,
            WS_TILED = 0,
            WS_EX_LTRREADING = 0,
            WS_OVERLAPPED = 0,
            WS_EX_RIGHTSCROLLBAR = 0,
            WS_EX_DLGMODALFRAME = 1,
            WS_EX_NOPARENTNOTIFY = 4,
            WS_EX_TOPMOST = 8,
            WS_EX_ACCEPTFILES = 16,
            WS_EX_TRANSPARENT = 32,
            WS_EX_MDICHILD = 64,
            WS_EX_TOOLWINDOW = 128,
            WS_EX_WINDOWEDGE = 256,
            WS_EX_PALETTEWINDOW = 392,
            WS_EX_CLIENTEDGE = 512,
            WS_EX_OVERLAPPEDWINDOW = 768,
            WS_EX_CONTEXTHELP = 1024,
            WS_EX_RIGHT = 4096,
            WS_EX_RTLREADING = 8192,
            WS_EX_LEFTSCROLLBAR = 16384,
            WS_MAXIMIZEBOX = 65536,
            WS_TABSTOP = 65536,
            WS_EX_CONTROLPARENT = 65536,
            WS_GROUP = 131072,
            WS_EX_STATICEDGE = 131072,
            WS_MINIMIZEBOX = 131072,
            WS_SIZEBOX = 262144,
            WS_THICKFRAME = 262144,
            WS_EX_APPWINDOW = 262144,
            WS_SYSMENU = 524288,
            WS_EX_LAYERED = 524288,
            WS_HSCROLL = 1048576,
            WS_EX_NOINHERITLAYOUT = 1048576,
            WS_VSCROLL = 2097152,
            WS_EX_LAYOUTRTL = 4194304,
            WS_DLGFRAME = 4194304,
            WS_BORDER = 8388608,
            WS_CAPTION = 12582912,
            WS_OVERLAPPEDWINDOW = 13565952,
            WS_TILEDWINDOW = 13565952,
            WS_MAXIMIZE = 16777216,
            WS_CLIPCHILDREN = 33554432,
            WS_EX_COMPOSITED = 33554432,
            WS_CLIPSIBLINGS = 67108864,
            WS_DISABLED = 134217728,
            WS_EX_NOACTIVATE = 134217728,
            WS_VISIBLE = 268435456,
            WS_MINIMIZE = 536870912,
            WS_ICONIC = 536870912,
            WS_CHILD = 1073741824,
            WS_CHILDWINDOW = 1073741824,
            WS_POPUP = 2147483648,
            WS_POPUPWINDOW = 2156396544,
        }

        ///  DwmGetWindowAttribute
        [Flags]
        public enum DWMWINDOWATTRIBUTE : uint
        {
            DWMWA_NCRENDERING_ENABLED = 1,
            DWMWA_NCRENDERING_POLICY,
            DWMWA_TRANSITIONS_FORCEDISABLED,
            DWMWA_ALLOW_NCPAINT,
            DWMWA_CAPTION_BUTTON_BOUNDS,
            DWMWA_NONCLIENT_RTL_LAYOUT,
            DWMWA_FORCE_ICONIC_REPRESENTATION,
            DWMWA_FLIP3D_POLICY,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            DWMWA_HAS_ICONIC_BITMAP,
            DWMWA_DISALLOW_PEEK,
            DWMWA_EXCLUDED_FROM_PEEK,
            DWMWA_CLOAK,
            DWMWA_CLOAKED,
            DWMWA_FREEZE_REPRESENTATION,
            DWMWA_LAST,
        }

        ///  CreateFileMapping
        [Flags]
        public enum PageProtection : uint
        {
            NoAccess = 0x01,
            Readonly = 0x02,
            ReadWrite = 0x04,
            WriteCopy = 0x08,
            Execute = 0x10,
            ExecuteRead = 0x20,
            ExecuteReadWrite = 0x40,
            ExecuteWriteCopy = 0x80,
            Guard = 0x100,
            NoCache = 0x200,
            WriteCombine = 0x400,
        }

        /// MapViewOfFile
        [Flags]
        public enum FileMapAccess : uint
        {
            FileMapCopy = 0x0001,
            FileMapWrite = 0x0002,
            FileMapRead = 0x0004,
            FileMapAllAccess = 0x001f,
            FileMapExecute = 0x0020,
        }

        /// CURSORINFO 
        [Flags]
        public enum CursorFlags : uint
        {
            Cursor_Hidden = 0x00000000,
            Cursor_Showing = 0x00000001,
            Cursor_Suppressed = 0x00000002,
        }

        /// AVIFileOpen
        [Flags]
        public enum OpenFileMode : uint
        {
            Read = 0x00000000,
            Write = 0x00000001,
            ReadWrite = 0x00000002,
            ShareCompat = 0x00000000,
            ShareExclusive = 0x00000010,
            ShareDenyWrite = 0x00000020,
            ShareDenyRead = 0x00000030,
            ShareDenyNone = 0x00000040,
            Parse = 0x00000100,
            Delete = 0x00000200,
            Verify = 0x00000400,
            Cancel = 0x00000800,
            Create = 0x00001000,
            Prompt = 0x00002000,
            Exist = 0x00004000,
            Reopen = 0x00008000
        }

        /// AVICOMPRESSOPTIONS
        [Flags]
        public enum AviCompression : uint
        {
            AVICOMPRESSF_DATARATE = 0x00000002,
            AVICOMPRESSF_INTERLEAVE = 0x00000001,
            AVICOMPRESSF_KEYFRAMES = 0x00000004,
            AVICOMPRESSF_VALID = 0x00000008
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct APPBARDATA
        {
            public uint cbSize;
            public IntPtr hWnd;
            public uint uCallbackMessage;
            public uint uEdge;
            public tagRect rc;
            public int lParam;
        }

        // Sends an appbar message to the system. 
        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern UIntPtr SHAppBarMessage(
            AppBarMessages dwMessage,					// Appbar message value to send.
            ref APPBARDATA pData);				// Address of an APPBARDATA structure. The content of the structure 
        // depends on the value set in the dwMessage parameter. 

        // The RegisterWindowMessage function defines a new window message that is guaranteed to be unique throughout 
        // the system. The message value can be used when sending or posting messages. 
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern UInt32 RegisterWindowMessage(
            [MarshalAs(UnmanagedType.LPTStr)]
			String lpString);		// Pointer to a null-terminated string that specifies the message to be registered. 

        public enum AppBarMessages : uint
        {
            /// <summary>
            /// Registers a new appbar and specifies the message identifier
            /// that the system should use to send notification messages to 
            /// the appbar. 
            /// </summary>
            New = 0x00000000,
            /// <summary>
            /// Unregisters an appbar, removing the bar from the system's 
            /// public list.
            /// </summary>
            Remove = 0x00000001,
            /// <summary>
            /// Requests a size and screen position for an appbar.
            /// </summary>
            QueryPos = 0x00000002,
            /// <summary>
            /// Sets the size and screen position of an appbar. 
            /// </summary>
            SetPos = 0x00000003,
            /// <summary>
            /// Retrieves the autohide and always-on-top states of the 
            /// Microsoft® Windows® taskbar. 
            /// </summary>
            GetState = 0x00000004,
            /// <summary>
            /// Retrieves the bounding rectangle of the Windows taskbar. 
            /// </summary>
            GetTaskBarPos = 0x00000005,
            /// <summary>
            /// Notifies the system that an appbar has been activated. 
            /// </summary>
            Activate = 0x00000006,
            /// <summary>
            /// Retrieves the handle to the autohide appbar associated with
            /// a particular edge of the screen. 
            /// </summary>
            GetAutoHideBar = 0x00000007,
            /// <summary>
            /// Registers or unregisters an autohide appbar for an edge of 
            /// the screen. 
            /// </summary>
            SetAutoHideBar = 0x00000008,
            /// <summary>
            /// Notifies the system when an appbar's position has changed. 
            /// </summary>
            WindowPosChanged = 0x00000009,
            /// <summary>
            /// Sets the state of the appbar's autohide and always-on-top 
            /// attributes.
            /// </summary>
            SetState = 0x0000000a
        }

        /// Return Type: BOOL->int
        ///lpKeyState: LPBYTE->BYTE*
        [DllImport("user32.dll", EntryPoint = "SetKeyboardState")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetKeyboardState([MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I1, SizeConst = 256)] byte[] lpKeyState);

        /// Return Type: HWND->HWND__*
        ///Point: POINT->tagPOINT
        [DllImport("user32.dll", EntryPoint = "WindowFromPoint")]
        public static extern IntPtr WindowFromPoint(tagPoint Point);

        /// Return Type: HWND->HWND__*
        [DllImport("user32.dll", EntryPoint = "GetDesktopWindow")]
        public static extern IntPtr GetDesktopWindow();

        /// Return Type: HWND->HWND__*
        ///hwndParent: HWND->HWND__*
        ///ptParentClientCoords: POINT->tagPOINT
        [DllImport("user32.dll", EntryPoint = "RealChildWindowFromPoint")]
        public static extern IntPtr RealChildWindowFromPoint([InAttribute()] IntPtr hwndParent, tagPoint ptParentClientCoords);

        /// Return Type: HWND->HWND__*
        ///hWndParent: HWND->HWND__*
        ///Point: POINT->tagPOINT
        [DllImport("user32.dll", EntryPoint = "ChildWindowFromPoint")]
        public static extern IntPtr ChildWindowFromPoint([In] IntPtr hWndParent, tagPoint Point);

        /// Return Type: HWND->HWND__*
        [DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
        public static extern IntPtr GetForegroundWindow();

        /// SW_RESTORE -> 9
        public const int SW_RESTORE = 9;

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///nCmdShow: int
        [DllImport("user32.dll", EntryPoint = "ShowWindow")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow([In] IntPtr hWnd, int nCmdShow);

        /// Return Type: BOOL->int
        ///lpPoint: LPPOINT->tagPOINT*
        [DllImport("user32.dll", EntryPoint = "GetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos([Out] out tagPoint lpPoint);

        /// Return Type: BOOL->int
        ///X: int
        ///Y: int
        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetCursorPos(int X, int Y);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///Msg: UINT->unsigned int
        ///wParam: WPARAM->UINT_PTR->IntPtr
        ///lParam: LPARAM->LONG_PTR->IntPtr
        [DllImport("user32.dll", EntryPoint = "PostMessage")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostMessage([In] IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        public const int WM_GETTITLEBARINFOEX = 0x033F;
        public const int WM_GETTEXT = 0x000D;
        public const int CCHILDREN_TITLEBAR = 5;

        //typedef struct tagTITLEBARINFOEX {
        //DWORD cbSize;
        //RECT  rcTitleBar;
        //DWORD rgstate[CCHILDREN_TITLEBAR+1];
        //RECT  rgrect[CCHILDREN_TITLEBAR+1];

        [StructLayout(LayoutKind.Sequential)]
        public struct TITLEBARINFOEX
        {
            public int cbSize;
            public tagRect rcTitleBar;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = CCHILDREN_TITLEBAR + 1)]
            public StateSystem[] rgstate;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = CCHILDREN_TITLEBAR + 1)]
            public tagRect[] rgrect;
        }

        public enum TitleBarStateElement : int
        {
            TitleBar = 0,
            //Reserved = 1,
            Minimize = 2,
            Maximize = 3,
            Help = 4,
            Close = 5,
        }

        public enum TitleBarLocationElement : int
        {
            //Reserved = 0,
            //Reserved = 1,
            Minimize = 2,
            Maximize = 3,
            Help = 4,
            Close = 5,
        }

        [Flags]
        public enum StateSystem : uint
        {
            STATE_SYSTEM_ALERT_HIGH = 0x10000000,
            STATE_SYSTEM_ALERT_LOW = 0x04000000,
            STATE_SYSTEM_ALERT_MEDIUM = 0x08000000,
            STATE_SYSTEM_ANIMATED = 0x00004000,
            STATE_SYSTEM_BUSY = 0x00000800,
            STATE_SYSTEM_CHECKED = 0x00000010,
            STATE_SYSTEM_COLLAPSED = 0x00000400,
            STATE_SYSTEM_DEFAULT = 0x00000100,
            STATE_SYSTEM_EXPANDED = 0x00000200,
            STATE_SYSTEM_EXTSELECTABLE = 0x02000000,
            STATE_SYSTEM_FLOATING = 0x00001000,
            STATE_SYSTEM_FOCUSABLE = 0x00100000,
            STATE_SYSTEM_FOCUSED = 0x00000004,
            STATE_SYSTEM_HASPOPUP = 0x40000000,
            STATE_SYSTEM_HOTTRACKED = 0x00000080,
            STATE_SYSTEM_INVISIBLE = 0x00008000,
            STATE_SYSTEM_LINKED = 0x00400000,
            STATE_SYSTEM_MARQUEED = 0x00002000,
            STATE_SYSTEM_MIXED = 0x00000020,
            STATE_SYSTEM_MOVEABLE = 0x00040000,
            STATE_SYSTEM_MULTISELECTABLE = 0x01000000,
            STATE_SYSTEM_NORMAL = 0x00000000,
            STATE_SYSTEM_OFFSCREEN = 0x00010000,
            STATE_SYSTEM_PRESSED = 0x00000008,
            STATE_SYSTEM_READONLY = 0x00000040,
            STATE_SYSTEM_SELECTABLE = 0x00200000,
            STATE_SYSTEM_SELECTED = 0x00000002,
            STATE_SYSTEM_SELFVOICING = 0x00080000,
            STATE_SYSTEM_SIZEABLE = 0x00020000,
            STATE_SYSTEM_TRAVERSED = 0x00800000,
            STATE_SYSTEM_UNAVAILABLE = 0x00000001,
            STATE_SYSTEM_VALID = 0x1FFFFFFF,
        }

        public const int PM_REMOVE = 1;
        public const int HC_ACTION = 0;
        public const int WH_MOUSE = 7;
        public const int WH_GETMESSAGE = 3;
        public const int WH_CALLWNDPROC = 4;
        public const int WH_FOREGROUNDIDLE = 11;
        public const int WM_MOUSEMOVE = 0x0200;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_LBUTTONUP = 0x0202;
        public const int WM_LBUTTONDBLCLK = 0x0203;
        public const int WM_RBUTTONDOWN = 0x0204;
        public const int WM_RBUTTONUP = 0x0205;
        public const int WM_RBUTTONDBLCLK = 0x0206;
        public const int WM_MBUTTONDOWN = 0x0207;
        public const int WM_MBUTTONUP = 0x0208;
        public const int WM_MBUTTONDBLCLK = 0x0209;
        public const int WM_NCMOUSEMOVE = 0x00A0;
        public const int WM_NCLBUTTONDOWN = 0x00A1;
        public const int WM_NCLBUTTONUP = 0x00A2;
        public const int WM_NCLBUTTONDBLCLK = 0x00A3;
        public const int WM_NCRBUTTONDOWN = 0x00A4;
        public const int WM_NCRBUTTONUP = 0x00A5;
        public const int WM_NCRBUTTONDBLCLK = 0x00A6;
        public const int WM_NCMBUTTONDOWN = 0x00A7;
        public const int WM_NCMBUTTONUP = 0x00A8;
        public const int WM_NCMBUTTONDBLCLK = 0x00A9;
        public const int WM_NULL = 0;
        public const int WM_GETTEXTLENGTH = 0x000E;

        [StructLayout(LayoutKind.Sequential)]
        public class MouseHookStruct
        {
            public tagPoint pt;
            public IntPtr hwnd;
            public int wHitTestCode;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CWPSTRUCT
        {
            public IntPtr lparam;
            public IntPtr wparam;
            public int message;
            public IntPtr hwnd;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSG
        {
            public IntPtr hwnd;
            public int message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public int pointX;
            public int pointY;
        }

        internal enum OBJID : uint
        {
            WINDOW = 0x00000000,
            SYSMENU = 0xFFFFFFFF,
            TITLEBAR = 0xFFFFFFFE,
            MENU = 0xFFFFFFFD,
            CLIENT = 0xFFFFFFFC,
            VSCROLL = 0xFFFFFFFB,
            HSCROLL = 0xFFFFFFFA,
            SIZEGRIP = 0xFFFFFFF9,
            CARET = 0xFFFFFFF8,
            CURSOR = 0xFFFFFFF7,
            ALERT = 0xFFFFFFF6,
            SOUND = 0xFFFFFFF5,
        }

        [DllImport("oleacc.dll")]
        internal static extern int AccessibleObjectFromWindow(IntPtr hwnd, OBJID id, ref Guid iid, [In, Out, MarshalAs(UnmanagedType.IUnknown)] ref object ppvObject);

        public static IAccessible IAccessibleObjectFromWindow(IntPtr hWnd)
        {
            Guid guid = new Guid("{618736E0-3C3D-11CF-810C-00AA00389B71}");
            object obj = null;
            int retVal = AccessibleObjectFromWindow(hWnd, OBJID.WINDOW, ref guid, ref obj);
            return (IAccessible)obj;
        }

        // local process hooking
        public delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        public delegate bool EnumWindow(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "SetWindowsHookEx", CharSet = CharSet.Unicode)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "UnhookWindowsHookEx", CharSet = CharSet.Unicode)]
        public static extern bool UnhookWindowsHookEx(int idHook);

        [DllImport("user32.dll", SetLastError = true, EntryPoint = "CallNextHookEx", CharSet = CharSet.Unicode)]
        public static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

        /// Return Type: DWORD->unsigned int->int
        ///hWnd: HWND->HWND__*
        ///lpdwProcessId: LPDWORD->DWORD*
        [DllImport("user32.dll", EntryPoint = "GetWindowThreadProcessId", CharSet = CharSet.Unicode)]
        public static extern int GetWindowThreadProcessId([In] IntPtr hWnd, IntPtr lpdwProcessId);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumThreadWindows(uint dwThreadId, EnumWindow lpfn, IntPtr lParam);


        [DllImport("user32.dll", EntryPoint = "EnumChildWindows")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows([In] IntPtr hWndParent, EnumWindow lpEnumFunc, IntPtr lParam);

        /// Return Type: BOOL
        ///uInterval: UINT->unsigned int
        [DllImport("user32.dll", EntryPoint = "SetDoubleClickTime")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetDoubleClickTime(uint uInterval);

        [DllImport("Kernel32.dll", SetLastError = false)]
        public static extern void CopyMemory(IntPtr dest, IntPtr src, UIntPtr size);

        public const Int64 INVALID_HANDLE_VALUE = -1;

        [DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateFileMapping(
            IntPtr hFile,
            IntPtr lpFileMappingAttributes,
            FileMapProtection flProtect,
            uint dwMaximumSizeHigh,
            uint dwMaximumSizeLow,
            [MarshalAs(UnmanagedType.LPTStr)] string lpName);

        [Flags]
        public enum FileMapProtection : uint
        {
            PageReadonly = 0x02,
            PageReadWrite = 0x04,
            PageWriteCopy = 0x08,
            PageExecuteRead = 0x20,
            PageExecuteReadWrite = 0x40,
            SectionCommit = 0x8000000,
            SectionImage = 0x1000000,
            SectionNoCache = 0x10000000,
            SectionReserve = 0x4000000,
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern IntPtr MapViewOfFile(
            IntPtr hFileMappingObject,
            FileMapAccess dwDesiredAccess,
            UInt32 dwFileOffsetHigh,
            UInt32 dwFileOffsetLow,
            UIntPtr dwNumberOfBytesToMap);

        [DllImport("kernel32.dll", SetLastError = true, CallingConvention = CallingConvention.Winapi)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsWow64Process([In] IntPtr process, [Out] out bool wow64Process);

        private static bool IsWow64Process(IntPtr processHandle)
        {
            IsWow64Process(processHandle, out bool retVal);
            return retVal;
        }

        public static bool IsWow64Process(Process process)
        {
            return IsWow64Process(process.Handle);
        }

        public static bool Is32BitProcess(Process process)
        {
            if (!Environment.Is64BitOperatingSystem)
            {
                return true;
            }

            IntPtr processHandle;
            try
            {
                processHandle = process.Handle;
            }
            catch
            {
                return false; // access is denied to the process
            }

            return IsWow64Process(processHandle);
        }

        [DllImport("gdi32.dll")]
        public static extern uint GetPixel(IntPtr hdc, int nXPos, int nYPos);

        /// Return Type: HDC->HDC__*
        ///hWnd: HWND->HWND__*
        [DllImport("user32.dll", EntryPoint = "GetWindowDC")]
        public static extern IntPtr GetWindowDC([In] IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern Int32 ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("user32.dll")]
        public static extern IntPtr GetSystemMenu(IntPtr hWnd, bool bRevert);

        [DllImport("user32.dll", EntryPoint = "GetMenu")]
        private static extern IntPtr GetMenuAPI(IntPtr hWnd);

        public static IntPtr GetMenu(IntPtr hWnd)
        {
            IntPtr theResult = IntPtr.Zero;
            if (IsTopLevelWindow(hWnd))
            {
                theResult = GetMenuAPI(hWnd);
            }
            return theResult;
        }

        private const int MN_GETHMENU = 0x1E1;

        public static IntPtr GetContextMenu(IntPtr hWnd)
        {
            IntPtr theResult = IntPtr.Zero;
            if (IsTopLevelWindow(hWnd))
            {
                IntPtr theReturn = SendMessageTimeout(hWnd, MN_GETHMENU, IntPtr.Zero, IntPtr.Zero, SendMessageTimeoutFlags.SMTO_ABORTIFHUNG, 30, out theResult);
                if (theReturn == IntPtr.Zero)
                {
                    throw new Exception("SendMessageTimeout failed, does the window exists?");
                }
            }
            return theResult;
        }

        [DllImport("user32.dll")]
        public static extern IntPtr GetSubMenu(IntPtr hMenu, int nPos);

        [DllImport("user32.dll")]
        public static extern int GetMenuItemCount(IntPtr hMenu);

        [DllImport("user32.dll")]
        public static extern bool GetMenuItemRect(IntPtr hWnd, IntPtr hMenu, uint uItem, out tagRect lprcItem);

        [DllImport("user32.dll")]
        private static extern int GetMenuString(IntPtr hMenu, uint uIDItem, [Out, MarshalAs(UnmanagedType.LPStr)] StringBuilder lpString, int nMaxCount, GetMenuFlag uFlag);

        public static string GetMenuString(IntPtr hMenu, uint uIDItem, GetMenuFlag uFlag)
        {
            StringBuilder menuString = new StringBuilder(1024);
            GetMenuString(hMenu, uIDItem, menuString, menuString.Capacity - 1, uFlag);
            return menuString.ToString();
        }

        [DllImport("user32.dll")]
        public static extern GetMenuStateMask GetMenuState(IntPtr hMenu, uint uId, GetMenuFlag uFlags);

        [Flags]
        public enum GetMenuFlag : uint
        {
            MF_BYCOMMAND = 0x00000000,
            MF_BYPOSITION = 0x00000400
        }

        [Flags]
        public enum GetMenuStateMask : uint
        {
            MF_GRAYED = 0x00000001,
            MF_DISABLED = 0x00000002,
            MF_CHECKED = 0x00000008,
            MF_POPUP = 0x00000010,
            MF_MENUBARBREAK = 0x00000020,
            MF_MENUBREAK = 0x00000040,
            MF_HILITE = 0x00000080,
            MF_OWNERDRAW = 0x00000100,
            MF_SEPARATOR = 0x00000800
        }

        public static bool IsTopLevelWindow(IntPtr Window)
        {
            WindowStyles Style = (WindowStyles)(long)GetWindowLongPtr(Window, GWL.GWL_STYLE);

            // WS_OVERLAPPED and WS_POPUP indicate a top level window.
            // WS_OVERLAPPED constant is 0, it does not make a good mask.  But all
            // WS_OVERLAPPED windows MUST have a caption so use WS_CAPTION instead.
            return Style.HasFlag(WindowStyles.WS_CAPTION) || Style.HasFlag(WindowStyles.WS_POPUP);
        }

        public const int LVM_FIRST = 0x1000;
        public const int LVM_GETITEMRECT = LVM_FIRST + 14;
        public const int LVM_GETGROUPRECT = LVM_FIRST + 98;
        public const int LVGGR_LABEL = 0x2;
        public const int LVIR_LABEL = 0x2;

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, int uMsg, IntPtr wParam, ref ComboBoxExItem lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, int uMsg, IntPtr wParam, ref tagRect lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, int uMsg, IntPtr wParam, ref DATETIMEPICKERINFO lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, int uMsg, IntPtr wParam, ref TITLEBARINFOEX lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, uint uMsg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, int uMsg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", EntryPoint = "SendMessageTimeout", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hwnd, int uMsg, ref int wParam, ref int lParam, SendMessageTimeoutFlags flags, uint uTimeout, out IntPtr result);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, ListBoxMessages Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, ListBoxMessages Msg, IntPtr wParam, string lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, ListBoxMessages Msg, IntPtr wParam, int[] lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0,
            SMTO_BLOCK = 0x1,
            SMTO_ABORTIFHUNG = 0x2,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x8
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool BringWindowToTop(IntPtr hWnd);

        /// Return Type: BOOL->int
        ///hWndLock: HWND->HWND__*
        [DllImport("user32.dll", EntryPoint = "LockWindowUpdate")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool LockWindowUpdate([In] IntPtr hWndLock);

        [DllImport("user32.dll")]
        public static extern bool UpdateWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        public static extern Boolean RedrawWindow(IntPtr hWnd, IntPtr lpRectUpdate, IntPtr hrgnUpdate, RedrawWindowFlags RedrawWindowFlags);

        [Flags]
        public enum RedrawWindowFlags : uint
        {
            /// <summary>
            /// Invalidates the rectangle or region that you specify in lprcUpdate or hrgnUpdate.
            /// You can set only one of these parameters to a non-NULL value. If both are NULL, RDW_INVALIDATE invalidates the entire window.
            /// </summary>
            Invalidate = 0x1,

            /// <summary>Causes the OS to post a WM_PAINT message to the window regardless of whether a portion of the window is invalid.</summary>
            InternalPaint = 0x2,

            /// <summary>
            /// Causes the window to receive a WM_ERASEBKGND message when the window is repainted.
            /// Specify this value in combination with the RDW_INVALIDATE value; otherwise, RDW_ERASE has no effect.
            /// </summary>
            Erase = 0x4,

            /// <summary>
            /// Validates the rectangle or region that you specify in lprcUpdate or hrgnUpdate.
            /// You can set only one of these parameters to a non-NULL value. If both are NULL, RDW_VALIDATE validates the entire window.
            /// This value does not affect internal WM_PAINT messages.
            /// </summary>
            Validate = 0x8,

            NoInternalPaint = 0x10,

            /// <summary>Suppresses any pending WM_ERASEBKGND messages.</summary>
            NoErase = 0x20,

            /// <summary>Excludes child windows, if any, from the repainting operation.</summary>
            NoChildren = 0x40,

            /// <summary>Includes child windows, if any, in the repainting operation.</summary>
            AllChildren = 0x80,

            /// <summary>Causes the affected windows, which you specify by setting the RDW_ALLCHILDREN and RDW_NOCHILDREN values, to receive WM_ERASEBKGND and WM_PAINT messages before the RedrawWindow returns, if necessary.</summary>
            UpdateNow = 0x100,

            /// <summary>
            /// Causes the affected windows, which you specify by setting the RDW_ALLCHILDREN and RDW_NOCHILDREN values, to receive WM_ERASEBKGND messages before RedrawWindow returns, if necessary.
            /// The affected windows receive WM_PAINT messages at the ordinary time.
            /// </summary>
            EraseNow = 0x200,

            Frame = 0x400,

            NoFrame = 0x800
        }

        public const int BM_GETCHECK = 0xF0;
        public const int BST_UNCHECKED = 0x0;
        public const int BST_CHECKED = 0x1;
        public const int BST_INDETERMINATE = 0x2;

        public enum ComboBoxStyle : int
        {
            Simple = 1,
            DropDown = 2,
            DropDownList = 3,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct COMBOBOXINFO
        {
            public Int32 cbSize;
            public tagRect rcItem;
            public tagRect rcButton;
            public ComboBoxButtonState buttonState;
            public IntPtr hwndCombo;
            public IntPtr hwndEdit;
            public IntPtr hwndList;
        }

        public enum ComboBoxButtonState
        {
            STATE_SYSTEM_NONE = 0,
            STATE_SYSTEM_INVISIBLE = 0x00008000,
            STATE_SYSTEM_PRESSED = 0x00000008
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern bool GetComboBoxInfo(IntPtr hWnd, ref COMBOBOXINFO pcbi);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowEnabled(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        public static string GetClassName(IntPtr hWnd)
        {
            StringBuilder className = new StringBuilder(1024);
            GetClassName(hWnd, className, className.Capacity - 1);
            return className.ToString();
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint GetWindowModuleFileName(IntPtr hwnd, StringBuilder lpszFileName, uint cchFileNameMax);

        /// <summary>
        /// Gets the module filename that the window handle belongs to.
        /// Has to be called from the same process that the window belongs to.
        /// </summary>
        /// <param name="hWnd">The window handle to get the module file name for </param>
        /// <returns></returns>
        public static string GetWindowModuleFileName(IntPtr hWnd)
        {
            StringBuilder fullPath = new StringBuilder(1024);
            GetWindowModuleFileName(hWnd, fullPath, (uint)fullPath.Capacity - 1);
            return fullPath.ToString();
        }

        [DllImport("user32.dll")]
        public static extern int GetDlgCtrlID(IntPtr hwndCtl);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr GetWindow(IntPtr hWnd, GetWindowCmd uCmd);

        public enum GetWindowCmd : uint
        {
            GW_HWNDFIRST = 0,
            GW_HWNDLAST = 1,
            GW_HWNDNEXT = 2,
            GW_HWNDPREV = 3,
            GW_OWNER = 4,
            GW_CHILD = 5,
            GW_ENABLEDPOPUP = 6
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetGUIThreadInfo(uint idThread, ref GUITHREADINFO lpgui);

        [StructLayout(LayoutKind.Sequential)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int flags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public tagRect rcCaret;
        }

        [DllImport("user32.dll")]
        public static extern uint GetQueueStatus(QueueStatusFlags flags);

        [Flags]
        public enum QueueStatusFlags : uint
        {
            QS_KEY = 0x0001,
            QS_MOUSEMOVE = 0x0002,
            QS_MOUSEBUTTON = 0x0004,
            QS_POSTMESSAGE = 0x0008,
            QS_TIMER = 0x0010,
            QS_PAINT = 0x0020,
            QS_SENDMESSAGE = 0x0040,
            QS_HOTKEY = 0x0080,
            QS_ALLPOSTMESSAGE = 0x0100,
            QS_RAWINPUT = 0x0400,
            QS_MOUSE = (QS_MOUSEMOVE | QS_MOUSEBUTTON),
            QS_INPUT = (QS_MOUSE | QS_KEY | QS_RAWINPUT),
            QS_REFRESH = (QS_HOTKEY | QS_KEY | QS_MOUSEBUTTON | QS_PAINT),
            QS_ALLEVENTS = (QS_INPUT | QS_POSTMESSAGE | QS_TIMER | QS_PAINT | QS_HOTKEY),
            QS_ALLINPUT = (QS_INPUT | QS_POSTMESSAGE | QS_TIMER | QS_PAINT | QS_HOTKEY | QS_SENDMESSAGE)
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SendInput(uint nInputs, INPUT[] pInputs, int cbSize);

        [StructLayout(LayoutKind.Sequential)]
        public struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public uint mouseData;
            public MOUSEEVENTF dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct INPUT
        {
            public INPUT_TYPE type;
            public InputUnion U;
            public static int Size
            {
                get { return Marshal.SizeOf(typeof(INPUT)); }
            }
        }

        [StructLayout(LayoutKind.Explicit)]
        public struct InputUnion
        {
            [FieldOffset(0)]
            public MOUSEINPUT mi;
            [FieldOffset(0)]
            public KEYBDINPUT ki;
            [FieldOffset(0)]
            public HARDWAREINPUT hi;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KEYBDINPUT
        {
            public VirtualKeyShort wVk;
            public ScanCodeShort wScan;
            public KEYEVENTF dwFlags;
            public uint time;
            public UIntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct HARDWAREINPUT
        {
            public int uMsg;
            public short wParamL;
            public short wParamH;
        }

        [Flags]
        public enum KEYEVENTF : uint
        {
            NONE = 0,
            EXTENDEDKEY = 0x0001,
            KEYUP = 0x0002,
            SCANCODE = 0x0008,
            UNICODE = 0x0004
        }

        public enum VirtualKeyShort : short
        {
            ///<summary>
            ///Left mouse button
            ///</summary>
            LBUTTON = 0x01,
            ///<summary>
            ///Right mouse button
            ///</summary>
            RBUTTON = 0x02,
            ///<summary>
            ///Control-break processing
            ///</summary>
            CANCEL = 0x03,
            ///<summary>
            ///Middle mouse button (three-button mouse)
            ///</summary>
            MBUTTON = 0x04,
            ///<summary>
            ///Windows 2000/XP: X1 mouse button
            ///</summary>
            XBUTTON1 = 0x05,
            ///<summary>
            ///Windows 2000/XP: X2 mouse button
            ///</summary>
            XBUTTON2 = 0x06,
            ///<summary>
            ///BACKSPACE key
            ///</summary>
            BACK = 0x08,
            ///<summary>
            ///TAB key
            ///</summary>
            TAB = 0x09,
            ///<summary>
            ///CLEAR key
            ///</summary>
            CLEAR = 0x0C,
            ///<summary>
            ///ENTER key
            ///</summary>
            RETURN = 0x0D,
            ///<summary>
            ///SHIFT key
            ///</summary>
            SHIFT = 0x10,
            ///<summary>
            ///CTRL key
            ///</summary>
            CONTROL = 0x11,
            ///<summary>
            ///ALT key
            ///</summary>
            MENU = 0x12,
            ///<summary>
            ///PAUSE key
            ///</summary>
            PAUSE = 0x13,
            ///<summary>
            ///CAPS LOCK key
            ///</summary>
            CAPITAL = 0x14,
            ///<summary>
            ///Input Method Editor (IME) Kana mode
            ///</summary>
            KANA = 0x15,
            ///<summary>
            ///IME Hangul mode
            ///</summary>
            HANGUL = 0x15,
            ///<summary>
            ///IME Junja mode
            ///</summary>
            JUNJA = 0x17,
            ///<summary>
            ///IME final mode
            ///</summary>
            FINAL = 0x18,
            ///<summary>
            ///IME Hanja mode
            ///</summary>
            HANJA = 0x19,
            ///<summary>
            ///IME Kanji mode
            ///</summary>
            KANJI = 0x19,
            ///<summary>
            ///ESC key
            ///</summary>
            ESCAPE = 0x1B,
            ///<summary>
            ///IME convert
            ///</summary>
            CONVERT = 0x1C,
            ///<summary>
            ///IME nonconvert
            ///</summary>
            NONCONVERT = 0x1D,
            ///<summary>
            ///IME accept
            ///</summary>
            ACCEPT = 0x1E,
            ///<summary>
            ///IME mode change request
            ///</summary>
            MODECHANGE = 0x1F,
            ///<summary>
            ///SPACEBAR
            ///</summary>
            SPACE = 0x20,
            ///<summary>
            ///PAGE UP key
            ///</summary>
            PRIOR = 0x21,
            ///<summary>
            ///PAGE DOWN key
            ///</summary>
            NEXT = 0x22,
            ///<summary>
            ///END key
            ///</summary>
            END = 0x23,
            ///<summary>
            ///HOME key
            ///</summary>
            HOME = 0x24,
            ///<summary>
            ///LEFT ARROW key
            ///</summary>
            LEFT = 0x25,
            ///<summary>
            ///UP ARROW key
            ///</summary>
            UP = 0x26,
            ///<summary>
            ///RIGHT ARROW key
            ///</summary>
            RIGHT = 0x27,
            ///<summary>
            ///DOWN ARROW key
            ///</summary>
            DOWN = 0x28,
            ///<summary>
            ///SELECT key
            ///</summary>
            SELECT = 0x29,
            ///<summary>
            ///PRINT key
            ///</summary>
            PRINT = 0x2A,
            ///<summary>
            ///EXECUTE key
            ///</summary>
            EXECUTE = 0x2B,
            ///<summary>
            ///PRINT SCREEN key
            ///</summary>
            SNAPSHOT = 0x2C,
            ///<summary>
            ///INS key
            ///</summary>
            INSERT = 0x2D,
            ///<summary>
            ///DEL key
            ///</summary>
            DELETE = 0x2E,
            ///<summary>
            ///HELP key
            ///</summary>
            HELP = 0x2F,
            ///<summary>
            ///0 key
            ///</summary>
            KEY_0 = 0x30,
            ///<summary>
            ///1 key
            ///</summary>
            KEY_1 = 0x31,
            ///<summary>
            ///2 key
            ///</summary>
            KEY_2 = 0x32,
            ///<summary>
            ///3 key
            ///</summary>
            KEY_3 = 0x33,
            ///<summary>
            ///4 key
            ///</summary>
            KEY_4 = 0x34,
            ///<summary>
            ///5 key
            ///</summary>
            KEY_5 = 0x35,
            ///<summary>
            ///6 key
            ///</summary>
            KEY_6 = 0x36,
            ///<summary>
            ///7 key
            ///</summary>
            KEY_7 = 0x37,
            ///<summary>
            ///8 key
            ///</summary>
            KEY_8 = 0x38,
            ///<summary>
            ///9 key
            ///</summary>
            KEY_9 = 0x39,
            ///<summary>
            ///A key
            ///</summary>
            KEY_A = 0x41,
            ///<summary>
            ///B key
            ///</summary>
            KEY_B = 0x42,
            ///<summary>
            ///C key
            ///</summary>
            KEY_C = 0x43,
            ///<summary>
            ///D key
            ///</summary>
            KEY_D = 0x44,
            ///<summary>
            ///E key
            ///</summary>
            KEY_E = 0x45,
            ///<summary>
            ///F key
            ///</summary>
            KEY_F = 0x46,
            ///<summary>
            ///G key
            ///</summary>
            KEY_G = 0x47,
            ///<summary>
            ///H key
            ///</summary>
            KEY_H = 0x48,
            ///<summary>
            ///I key
            ///</summary>
            KEY_I = 0x49,
            ///<summary>
            ///J key
            ///</summary>
            KEY_J = 0x4A,
            ///<summary>
            ///K key
            ///</summary>
            KEY_K = 0x4B,
            ///<summary>
            ///L key
            ///</summary>
            KEY_L = 0x4C,
            ///<summary>
            ///M key
            ///</summary>
            KEY_M = 0x4D,
            ///<summary>
            ///N key
            ///</summary>
            KEY_N = 0x4E,
            ///<summary>
            ///O key
            ///</summary>
            KEY_O = 0x4F,
            ///<summary>
            ///P key
            ///</summary>
            KEY_P = 0x50,
            ///<summary>
            ///Q key
            ///</summary>
            KEY_Q = 0x51,
            ///<summary>
            ///R key
            ///</summary>
            KEY_R = 0x52,
            ///<summary>
            ///S key
            ///</summary>
            KEY_S = 0x53,
            ///<summary>
            ///T key
            ///</summary>
            KEY_T = 0x54,
            ///<summary>
            ///U key
            ///</summary>
            KEY_U = 0x55,
            ///<summary>
            ///V key
            ///</summary>
            KEY_V = 0x56,
            ///<summary>
            ///W key
            ///</summary>
            KEY_W = 0x57,
            ///<summary>
            ///X key
            ///</summary>
            KEY_X = 0x58,
            ///<summary>
            ///Y key
            ///</summary>
            KEY_Y = 0x59,
            ///<summary>
            ///Z key
            ///</summary>
            KEY_Z = 0x5A,
            ///<summary>
            ///Left Windows key (Microsoft Natural keyboard) 
            ///</summary>
            LWIN = 0x5B,
            ///<summary>
            ///Right Windows key (Natural keyboard)
            ///</summary>
            RWIN = 0x5C,
            ///<summary>
            ///Applications key (Natural keyboard)
            ///</summary>
            APPS = 0x5D,
            ///<summary>
            ///Computer Sleep key
            ///</summary>
            SLEEP = 0x5F,
            ///<summary>
            ///Numeric keypad 0 key
            ///</summary>
            NUMPAD0 = 0x60,
            ///<summary>
            ///Numeric keypad 1 key
            ///</summary>
            NUMPAD1 = 0x61,
            ///<summary>
            ///Numeric keypad 2 key
            ///</summary>
            NUMPAD2 = 0x62,
            ///<summary>
            ///Numeric keypad 3 key
            ///</summary>
            NUMPAD3 = 0x63,
            ///<summary>
            ///Numeric keypad 4 key
            ///</summary>
            NUMPAD4 = 0x64,
            ///<summary>
            ///Numeric keypad 5 key
            ///</summary>
            NUMPAD5 = 0x65,
            ///<summary>
            ///Numeric keypad 6 key
            ///</summary>
            NUMPAD6 = 0x66,
            ///<summary>
            ///Numeric keypad 7 key
            ///</summary>
            NUMPAD7 = 0x67,
            ///<summary>
            ///Numeric keypad 8 key
            ///</summary>
            NUMPAD8 = 0x68,
            ///<summary>
            ///Numeric keypad 9 key
            ///</summary>
            NUMPAD9 = 0x69,
            ///<summary>
            ///Multiply key
            ///</summary>
            MULTIPLY = 0x6A,
            ///<summary>
            ///Add key
            ///</summary>
            ADD = 0x6B,
            ///<summary>
            ///Separator key
            ///</summary>
            SEPARATOR = 0x6C,
            ///<summary>
            ///Subtract key
            ///</summary>
            SUBTRACT = 0x6D,
            ///<summary>
            ///Decimal key
            ///</summary>
            DECIMAL = 0x6E,
            ///<summary>
            ///Divide key
            ///</summary>
            DIVIDE = 0x6F,
            ///<summary>
            ///F1 key
            ///</summary>
            F1 = 0x70,
            ///<summary>
            ///F2 key
            ///</summary>
            F2 = 0x71,
            ///<summary>
            ///F3 key
            ///</summary>
            F3 = 0x72,
            ///<summary>
            ///F4 key
            ///</summary>
            F4 = 0x73,
            ///<summary>
            ///F5 key
            ///</summary>
            F5 = 0x74,
            ///<summary>
            ///F6 key
            ///</summary>
            F6 = 0x75,
            ///<summary>
            ///F7 key
            ///</summary>
            F7 = 0x76,
            ///<summary>
            ///F8 key
            ///</summary>
            F8 = 0x77,
            ///<summary>
            ///F9 key
            ///</summary>
            F9 = 0x78,
            ///<summary>
            ///F10 key
            ///</summary>
            F10 = 0x79,
            ///<summary>
            ///F11 key
            ///</summary>
            F11 = 0x7A,
            ///<summary>
            ///F12 key
            ///</summary>
            F12 = 0x7B,
            ///<summary>
            ///F13 key
            ///</summary>
            F13 = 0x7C,
            ///<summary>
            ///F14 key
            ///</summary>
            F14 = 0x7D,
            ///<summary>
            ///F15 key
            ///</summary>
            F15 = 0x7E,
            ///<summary>
            ///F16 key
            ///</summary>
            F16 = 0x7F,
            ///<summary>
            ///F17 key  
            ///</summary>
            F17 = 0x80,
            ///<summary>
            ///F18 key  
            ///</summary>
            F18 = 0x81,
            ///<summary>
            ///F19 key  
            ///</summary>
            F19 = 0x82,
            ///<summary>
            ///F20 key  
            ///</summary>
            F20 = 0x83,
            ///<summary>
            ///F21 key  
            ///</summary>
            F21 = 0x84,
            ///<summary>
            ///F22 key, (PPC only) Key used to lock device.
            ///</summary>
            F22 = 0x85,
            ///<summary>
            ///F23 key  
            ///</summary>
            F23 = 0x86,
            ///<summary>
            ///F24 key  
            ///</summary>
            F24 = 0x87,
            ///<summary>
            ///NUM LOCK key
            ///</summary>
            NUMLOCK = 0x90,
            ///<summary>
            ///SCROLL LOCK key
            ///</summary>
            SCROLL = 0x91,
            ///<summary>
            ///Left SHIFT key
            ///</summary>
            LSHIFT = 0xA0,
            ///<summary>
            ///Right SHIFT key
            ///</summary>
            RSHIFT = 0xA1,
            ///<summary>
            ///Left CONTROL key
            ///</summary>
            LCONTROL = 0xA2,
            ///<summary>
            ///Right CONTROL key
            ///</summary>
            RCONTROL = 0xA3,
            ///<summary>
            ///Left MENU key
            ///</summary>
            LMENU = 0xA4,
            ///<summary>
            ///Right MENU key
            ///</summary>
            RMENU = 0xA5,
            ///<summary>
            ///Windows 2000/XP: Browser Back key
            ///</summary>
            BROWSER_BACK = 0xA6,
            ///<summary>
            ///Windows 2000/XP: Browser Forward key
            ///</summary>
            BROWSER_FORWARD = 0xA7,
            ///<summary>
            ///Windows 2000/XP: Browser Refresh key
            ///</summary>
            BROWSER_REFRESH = 0xA8,
            ///<summary>
            ///Windows 2000/XP: Browser Stop key
            ///</summary>
            BROWSER_STOP = 0xA9,
            ///<summary>
            ///Windows 2000/XP: Browser Search key 
            ///</summary>
            BROWSER_SEARCH = 0xAA,
            ///<summary>
            ///Windows 2000/XP: Browser Favorites key
            ///</summary>
            BROWSER_FAVORITES = 0xAB,
            ///<summary>
            ///Windows 2000/XP: Browser Start and Home key
            ///</summary>
            BROWSER_HOME = 0xAC,
            ///<summary>
            ///Windows 2000/XP: Volume Mute key
            ///</summary>
            VOLUME_MUTE = 0xAD,
            ///<summary>
            ///Windows 2000/XP: Volume Down key
            ///</summary>
            VOLUME_DOWN = 0xAE,
            ///<summary>
            ///Windows 2000/XP: Volume Up key
            ///</summary>
            VOLUME_UP = 0xAF,
            ///<summary>
            ///Windows 2000/XP: Next Track key
            ///</summary>
            MEDIA_NEXT_TRACK = 0xB0,
            ///<summary>
            ///Windows 2000/XP: Previous Track key
            ///</summary>
            MEDIA_PREV_TRACK = 0xB1,
            ///<summary>
            ///Windows 2000/XP: Stop Media key
            ///</summary>
            MEDIA_STOP = 0xB2,
            ///<summary>
            ///Windows 2000/XP: Play/Pause Media key
            ///</summary>
            MEDIA_PLAY_PAUSE = 0xB3,
            ///<summary>
            ///Windows 2000/XP: Start Mail key
            ///</summary>
            LAUNCH_MAIL = 0xB4,
            ///<summary>
            ///Windows 2000/XP: Select Media key
            ///</summary>
            LAUNCH_MEDIA_SELECT = 0xB5,
            ///<summary>
            ///Windows 2000/XP: Start Application 1 key
            ///</summary>
            LAUNCH_APP1 = 0xB6,
            ///<summary>
            ///Windows 2000/XP: Start Application 2 key
            ///</summary>
            LAUNCH_APP2 = 0xB7,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard.
            ///</summary>
            OEM_1 = 0xBA,
            ///<summary>
            ///Windows 2000/XP: For any country/region, the '+' key
            ///</summary>
            OEM_PLUS = 0xBB,
            ///<summary>
            ///Windows 2000/XP: For any country/region, the ',' key
            ///</summary>
            OEM_COMMA = 0xBC,
            ///<summary>
            ///Windows 2000/XP: For any country/region, the '-' key
            ///</summary>
            OEM_MINUS = 0xBD,
            ///<summary>
            ///Windows 2000/XP: For any country/region, the '.' key
            ///</summary>
            OEM_PERIOD = 0xBE,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard.
            ///</summary>
            OEM_2 = 0xBF,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard. 
            ///</summary>
            OEM_3 = 0xC0,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard. 
            ///</summary>
            OEM_4 = 0xDB,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard. 
            ///</summary>
            OEM_5 = 0xDC,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard. 
            ///</summary>
            OEM_6 = 0xDD,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard. 
            ///</summary>
            OEM_7 = 0xDE,
            ///<summary>
            ///Used for miscellaneous characters; it can vary by keyboard.
            ///</summary>
            OEM_8 = 0xDF,
            ///<summary>
            ///Windows 2000/XP: Either the angle bracket key or the backslash key on the RT 102-key keyboard
            ///</summary>
            OEM_102 = 0xE2,
            ///<summary>
            ///Windows 95/98/Me, Windows NT 4.0, Windows 2000/XP: IME PROCESS key
            ///</summary>
            PROCESSKEY = 0xE5,
            ///<summary>
            ///Windows 2000/XP: Used to pass Unicode characters as if they were keystrokes.
            ///The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information,
            ///see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
            ///</summary>
            PACKET = 0xE7,
            ///<summary>
            ///Attn key
            ///</summary>
            ATTN = 0xF6,
            ///<summary>
            ///CrSel key
            ///</summary>
            CRSEL = 0xF7,
            ///<summary>
            ///ExSel key
            ///</summary>
            EXSEL = 0xF8,
            ///<summary>
            ///Erase EOF key
            ///</summary>
            EREOF = 0xF9,
            ///<summary>
            ///Play key
            ///</summary>
            PLAY = 0xFA,
            ///<summary>
            ///Zoom key
            ///</summary>
            ZOOM = 0xFB,
            ///<summary>
            ///Reserved 
            ///</summary>
            NONAME = 0xFC,
            ///<summary>
            ///PA1 key
            ///</summary>
            PA1 = 0xFD,
            ///<summary>
            ///Clear key
            ///</summary>
            OEM_CLEAR = 0xFE
        }

        public enum ScanCodeShort : short
        {
            LBUTTON = 0,
            RBUTTON = 0,
            CANCEL = 70,
            MBUTTON = 0,
            XBUTTON1 = 0,
            XBUTTON2 = 0,
            BACK = 14,
            TAB = 15,
            CLEAR = 76,
            RETURN = 28,
            SHIFT = 42,
            CONTROL = 29,
            MENU = 56,
            PAUSE = 0,
            CAPITAL = 58,
            KANA = 0,
            HANGUL = 0,
            JUNJA = 0,
            FINAL = 0,
            HANJA = 0,
            KANJI = 0,
            ESCAPE = 1,
            CONVERT = 0,
            NONCONVERT = 0,
            ACCEPT = 0,
            MODECHANGE = 0,
            SPACE = 57,
            PRIOR = 73,
            NEXT = 81,
            END = 79,
            HOME = 71,
            LEFT = 75,
            UP = 72,
            RIGHT = 77,
            DOWN = 80,
            SELECT = 0,
            PRINT = 0,
            EXECUTE = 0,
            SNAPSHOT = 84,
            INSERT = 82,
            DELETE = 83,
            HELP = 99,
            KEY_0 = 11,
            KEY_1 = 2,
            KEY_2 = 3,
            KEY_3 = 4,
            KEY_4 = 5,
            KEY_5 = 6,
            KEY_6 = 7,
            KEY_7 = 8,
            KEY_8 = 9,
            KEY_9 = 10,
            KEY_A = 30,
            KEY_B = 48,
            KEY_C = 46,
            KEY_D = 32,
            KEY_E = 18,
            KEY_F = 33,
            KEY_G = 34,
            KEY_H = 35,
            KEY_I = 23,
            KEY_J = 36,
            KEY_K = 37,
            KEY_L = 38,
            KEY_M = 50,
            KEY_N = 49,
            KEY_O = 24,
            KEY_P = 25,
            KEY_Q = 16,
            KEY_R = 19,
            KEY_S = 31,
            KEY_T = 20,
            KEY_U = 22,
            KEY_V = 47,
            KEY_W = 17,
            KEY_X = 45,
            KEY_Y = 21,
            KEY_Z = 44,
            LWIN = 91,
            RWIN = 92,
            APPS = 93,
            SLEEP = 95,
            NUMPAD0 = 82,
            NUMPAD1 = 79,
            NUMPAD2 = 80,
            NUMPAD3 = 81,
            NUMPAD4 = 75,
            NUMPAD5 = 76,
            NUMPAD6 = 77,
            NUMPAD7 = 71,
            NUMPAD8 = 72,
            NUMPAD9 = 73,
            MULTIPLY = 55,
            ADD = 78,
            SEPARATOR = 0,
            SUBTRACT = 74,
            DECIMAL = 83,
            DIVIDE = 53,
            F1 = 59,
            F2 = 60,
            F3 = 61,
            F4 = 62,
            F5 = 63,
            F6 = 64,
            F7 = 65,
            F8 = 66,
            F9 = 67,
            F10 = 68,
            F11 = 87,
            F12 = 88,
            F13 = 100,
            F14 = 101,
            F15 = 102,
            F16 = 103,
            F17 = 104,
            F18 = 105,
            F19 = 106,
            F20 = 107,
            F21 = 108,
            F22 = 109,
            F23 = 110,
            F24 = 118,
            NUMLOCK = 69,
            SCROLL = 70,
            LSHIFT = 42,
            RSHIFT = 54,
            LCONTROL = 29,
            RCONTROL = 29,
            LMENU = 56,
            RMENU = 56,
            BROWSER_BACK = 106,
            BROWSER_FORWARD = 105,
            BROWSER_REFRESH = 103,
            BROWSER_STOP = 104,
            BROWSER_SEARCH = 101,
            BROWSER_FAVORITES = 102,
            BROWSER_HOME = 50,
            VOLUME_MUTE = 32,
            VOLUME_DOWN = 46,
            VOLUME_UP = 48,
            MEDIA_NEXT_TRACK = 25,
            MEDIA_PREV_TRACK = 16,
            MEDIA_STOP = 36,
            MEDIA_PLAY_PAUSE = 34,
            LAUNCH_MAIL = 108,
            LAUNCH_MEDIA_SELECT = 109,
            LAUNCH_APP1 = 107,
            LAUNCH_APP2 = 33,
            OEM_1 = 39,
            OEM_PLUS = 13,
            OEM_COMMA = 51,
            OEM_MINUS = 12,
            OEM_PERIOD = 52,
            OEM_2 = 53,
            OEM_3 = 41,
            OEM_4 = 26,
            OEM_5 = 43,
            OEM_6 = 27,
            OEM_7 = 40,
            OEM_8 = 0,
            OEM_102 = 86,
            PROCESSKEY = 0,
            PACKET = 0,
            ATTN = 0,
            CRSEL = 0,
            EXSEL = 0,
            EREOF = 93,
            PLAY = 0,
            ZOOM = 98,
            NONAME = 0,
            PA1 = 0,
            OEM_CLEAR = 0,
        }


        [Flags]
        public enum MOUSEEVENTF : uint
        {
            ABSOLUTE = 0x8000,
            HWHEEL = 0x01000,
            MOVE = 0x0001,
            MOVE_NOCOALESCE = 0x2000,
            LEFTDOWN = 0x0002,
            LEFTUP = 0x0004,
            RIGHTDOWN = 0x0008,
            RIGHTUP = 0x0010,
            MIDDLEDOWN = 0x0020,
            MIDDLEUP = 0x0040,
            VIRTUALDESK = 0x4000,
            WHEEL = 0x0800,
            XDOWN = 0x0080,
            XUP = 0x0100
        }

        public enum INPUT_TYPE : uint
        {
            INPUT_MOUSE = 0,
            INPUT_KEYBOARD = 1,
            INPUT_HARDWAEE = 2,
        }

        [DllImport("user32.dll")]
        public static extern int GetSystemMetrics(SystemMetric smIndex);

        public enum SystemMetric
        {
            SM_CXSCREEN = 0,  // 0x00
            SM_CYSCREEN = 1,  // 0x01
            SM_CXVSCROLL = 2,  // 0x02
            SM_CYHSCROLL = 3,  // 0x03
            SM_CYCAPTION = 4,  // 0x04
            SM_CXBORDER = 5,  // 0x05
            SM_CYBORDER = 6,  // 0x06
            SM_CXDLGFRAME = 7,  // 0x07
            SM_CXFIXEDFRAME = 7,  // 0x07
            SM_CYDLGFRAME = 8,  // 0x08
            SM_CYFIXEDFRAME = 8,  // 0x08
            SM_CYVTHUMB = 9,  // 0x09
            SM_CXHTHUMB = 10, // 0x0A
            SM_CXICON = 11, // 0x0B
            SM_CYICON = 12, // 0x0C
            SM_CXCURSOR = 13, // 0x0D
            SM_CYCURSOR = 14, // 0x0E
            SM_CYMENU = 15, // 0x0F
            SM_CXFULLSCREEN = 16, // 0x10
            SM_CYFULLSCREEN = 17, // 0x11
            SM_CYKANJIWINDOW = 18, // 0x12
            SM_MOUSEPRESENT = 19, // 0x13
            SM_CYVSCROLL = 20, // 0x14
            SM_CXHSCROLL = 21, // 0x15
            SM_DEBUG = 22, // 0x16
            SM_SWAPBUTTON = 23, // 0x17
            SM_CXMIN = 28, // 0x1C
            SM_CYMIN = 29, // 0x1D
            SM_CXSIZE = 30, // 0x1E
            SM_CYSIZE = 31, // 0x1F
            SM_CXSIZEFRAME = 32, // 0x20
            SM_CXFRAME = 32, // 0x20
            SM_CYSIZEFRAME = 33, // 0x21
            SM_CYFRAME = 33, // 0x21
            SM_CXMINTRACK = 34, // 0x22
            SM_CYMINTRACK = 35, // 0x23
            SM_CXDOUBLECLK = 36, // 0x24
            SM_CYDOUBLECLK = 37, // 0x25
            SM_CXICONSPACING = 38, // 0x26
            SM_CYICONSPACING = 39, // 0x27
            SM_MENUDROPALIGNMENT = 40, // 0x28
            SM_PENWINDOWS = 41, // 0x29
            SM_DBCSENABLED = 42, // 0x2A
            SM_CMOUSEBUTTONS = 43, // 0x2B
            SM_SECURE = 44, // 0x2C
            SM_CXEDGE = 45, // 0x2D
            SM_CYEDGE = 46, // 0x2E
            SM_CXMINSPACING = 47, // 0x2F
            SM_CYMINSPACING = 48, // 0x30
            SM_CXSMICON = 49, // 0x31
            SM_CYSMICON = 50, // 0x32
            SM_CYSMCAPTION = 51, // 0x33
            SM_CXSMSIZE = 52, // 0x34
            SM_CYSMSIZE = 53, // 0x35
            SM_CXMENUSIZE = 54, // 0x36
            SM_CYMENUSIZE = 55, // 0x37
            SM_ARRANGE = 56, // 0x38
            SM_CXMINIMIZED = 57, // 0x39
            SM_CYMINIMIZED = 58, // 0x3A
            SM_CXMAXTRACK = 59, // 0x3B
            SM_CYMAXTRACK = 60, // 0x3C
            SM_CXMAXIMIZED = 61, // 0x3D
            SM_CYMAXIMIZED = 62, // 0x3E
            SM_NETWORK = 63, // 0x3F
            SM_CLEANBOOT = 67, // 0x43
            SM_CXDRAG = 68, // 0x44
            SM_CYDRAG = 69, // 0x45
            SM_SHOWSOUNDS = 70, // 0x46
            SM_CXMENUCHECK = 71, // 0x47
            SM_CYMENUCHECK = 72, // 0x48
            SM_SLOWMACHINE = 73, // 0x49
            SM_MIDEASTENABLED = 74, // 0x4A
            SM_MOUSEWHEELPRESENT = 75, // 0x4B
            SM_XVIRTUALSCREEN = 76, // 0x4C
            SM_YVIRTUALSCREEN = 77, // 0x4D
            SM_CXVIRTUALSCREEN = 78, // 0x4E
            SM_CYVIRTUALSCREEN = 79, // 0x4F
            SM_CMONITORS = 80, // 0x50
            SM_SAMEDISPLAYFORMAT = 81, // 0x51
            SM_IMMENABLED = 82, // 0x52
            SM_CXFOCUSBORDER = 83, // 0x53
            SM_CYFOCUSBORDER = 84, // 0x54
            SM_TABLETPC = 86, // 0x56
            SM_MEDIACENTER = 87, // 0x57
            SM_STARTER = 88, // 0x58
            SM_SERVERR2 = 89, // 0x59
            SM_MOUSEHORIZONTALWHEELPRESENT = 91, // 0x5B
            SM_CXPADDEDBORDER = 92, // 0x5C

            SM_DIGITIZER = 94, // 0x5E
            SM_MAXIMUMTOUCHES = 95, // 0x5F

            SM_REMOTESESSION = 0x1000, // 0x1000

            SM_SHUTTINGDOWN = 0x2000, // 0x2000
            SM_REMOTECONTROL = 0x2001, // 0x2001
            SM_CONVERTABLESLATEMODE = 0x2003,
            SM_SYSTEMDOCKED = 0x2004
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr MonitorFromPoint(tagPoint pt, MonitorOptions dwFlags);

        public enum MonitorOptions : uint
        {
            MONITOR_DEFAULTTONULL = 0x00000000,
            MONITOR_DEFAULTTOPRIMARY = 0x00000001,
            MONITOR_DEFAULTTONEAREST = 0x00000002
        }

        /// LB_ERR -> (-1)
        public const int LB_ERR = -1;

        [Flags]
        public enum ListBoxMessages : int
        {
            LB_ADDSTRING = 0x0180,
            LB_SETSEL = 0x0185,
            LB_GETSELITEMS = 0x0191,
            LB_GETSELCOUNT = 0x0190,
            LB_GETCURSEL = 0x0188,
            LB_SELECTSTRING = 0x018C,
            LB_SETCURSEL = 0x0186,
            LB_FINDSTRING = 0x018F,
            LB_FINDSTRINGEXACT = 0x01A2,
            LB_GETCOUNT = 0x018B,
            LB_GETSEL = 0x0187,
            LB_GETTEXT = 0x0189,
            LB_RESETCONTENT = 0x0184,
            LB_SETHORIZONTALEXTENT = 0x0194,
            LB_GETHORIZONTALEXTENT = 0x0193,
            LB_GETTOPINDEX = 0x018E,
            LB_SETTOPINDEX = 0x0197,
            LB_GETITEMRECT = 0x0198,
            LB_INSERTSTRING = 0x0181,
            LB_DELETESTRING = 0x0182,
            LB_GETITEMDATA = 0x0199,
        }

        /// SB_LEFT -> 6
        public const int SB_LEFT = 6;

        /// WM_HSCROLL -> 0x0114
        public const int WM_HSCROLL = 276;

        /// WM_VSCROLL -> 0x0115
        public const int WM_VSCROLL = 277;

        /// Return Type: HWND->HWND__*
        ///hWndParent: HWND->HWND__*
        ///hWndChildAfter: HWND->HWND__*
        ///lpszClass: LPCWSTR->WCHAR*
        ///lpszWindow: LPCWSTR->WCHAR*
        [DllImport("user32.dll", EntryPoint = "FindWindowEx", CharSet = CharSet.Unicode)]
        public static extern IntPtr FindWindowEx([In] IntPtr hWndParent, [In] IntPtr hWndChildAfter, [In] [MarshalAs(UnmanagedType.LPWStr)] string lpszClass, [In] [MarshalAsAttribute(UnmanagedType.LPWStr)] string lpszWindow);

        /// CB_ERR -> (-1)
        public const int CB_ERR = -1;

        public enum ComboBoxMessages : int
        {
            CB_GETEDITSEL = 0x0140,
            CB_LIMITTEXT = 0x0141,
            CB_SETEDITSEL = 0x0142,
            CB_ADDSTRING = 0x0143,
            CB_DELETESTRING = 0x0144,
            CB_DIR = 0x0145,
            CB_GETCOUNT = 0x0146,
            CB_GETCURSEL = 0x0147,
            CB_GETLBTEXT = 0x0148,
            CB_GETLBTEXTLEN = 0x0149,
            CB_INSERTSTRING = 0x014A,
            CB_RESETCONTENT = 0x014B,
            CB_FINDSTRING = 0x014C,
            CB_SELECTSTRING = 0x014D,
            CB_SETCURSEL = 0x014E,
            CB_SHOWDROPDOWN = 0x014F,
            CB_GETITEMDATA = 0x0150,
            CB_SETITEMDATA = 0x0151,
            CB_GETDROPPEDCONTROLRECT = 0x0152,
            CB_SETITEMHEIGHT = 0x0153,
            CB_GETITEMHEIGHT = 0x0154,
            CB_SETEXTENDEDUI = 0x0155,
            CB_GETEXTENDEDUI = 0x0156,
            CB_GETDROPPEDSTATE = 0x0157,
            CB_FINDSTRINGEXACT = 0x0158,
            CB_SETLOCALE = 0x0159,
            CB_GETLOCALE = 0x015A,
            CB_GETTOPINDEX = 0x015B,
            CB_SETTOPINDEX = 0x015C,
            CB_GETHORIZONTALEXTENT = 0x015D,
            CB_SETHORIZONTALEXTENT = 0x015E,
            CB_GETDROPPEDWIDTH = 0x015F,
            CB_SETDROPPEDWIDTH = 0x0160,
            CB_INITSTORAGE = 0x0161
            /*  CB_MSGMAX    = 354*/
        }

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, ComboBoxMessages Msg, IntPtr wParam, string lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, ComboBoxMessages Msg, IntPtr wParam, IntPtr lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, ListBoxMessages Msg, IntPtr wParam, ref tagRect lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, int msg, IntPtr wParam, ref ToolInfo lParam, SendMessageTimeoutFlags fuFlags, uint uTimeout, out IntPtr lpdwResult);

        public const int WM_USER = 0x400;
        
        public const int CBEIF_TEXT = 0x1;
        public const int CBEM_GETUNICODEFORMAT = (0x2000 + 6);

        public const int CBEM_GETITEMA = WM_USER + 4;
        public const int CBEM_GETCOMBOCONTROL = WM_USER + 6;
        public const int CBEM_GETEDITCONTROL = WM_USER + 7;

        public const int CBEM_GETITEMW = WM_USER + 13;

        public struct ComboBoxExItem
        {
            public int mask;
            public IntPtr iItem;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string pszText;
            public int cchTextMax;
            public int iImage;
            public int iSelectedImage;
            public int iOverlay;
            public int iIndent;
            public IntPtr lParam;
        }

        public struct ToolInfo
        {
            public int cbSize;
            public int uFlags;
            public IntPtr hWnd;
            public IntPtr uId;
            public tagRect rect;
            public IntPtr hInst;
            [MarshalAs(UnmanagedType.LPTStr)]
            public string lpszText;
            public IntPtr lParam;
        }

        private const int TTM_GETCURRENTTOOLW = (0x0400 + 59);

        public static ToolInfo GetToolInfo(IntPtr hWnd, uint timeout)
        {
            ToolInfo info = new ToolInfo();
            info.cbSize = Marshal.SizeOf(info);
            
            IntPtr messageResult;
            IntPtr sendResult = SendMessageTimeout(hWnd, TTM_GETCURRENTTOOLW, IntPtr.Zero, ref info, SendMessageTimeoutFlags.SMTO_NORMAL, timeout, out messageResult);
            if (sendResult == IntPtr.Zero) //Failed
            {
                return new ToolInfo();
            }
            else
            {
                if (messageResult == IntPtr.Zero)
                {
                    //Failed
                    return new ToolInfo();
                }
                return info;
            }
        }

        [DllImport("User32.dll")]
        public static extern int GetWindowThreadProcessId(IntPtr hWnd, out int pid);

        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("user32.dll", CharSet = CharSet.Unicode, ExactSpelling = true)]
        public static extern bool BlockInput([In, MarshalAs(UnmanagedType.Bool)] bool fBlockIt);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsZoomed(IntPtr hWnd);

        /// Return Type: SHORT->short
        ///nVirtKey: int
        [DllImport("user32.dll", EntryPoint = "GetKeyState")]
        public static extern short GetKeyState(int nVirtKey);

        /// Return Type: void
        ///dwFlags: DWORD->unsigned int
        ///dx: DWORD->unsigned int
        ///dy: DWORD->unsigned int
        ///dwData: DWORD->unsigned int
        ///dwExtraInfo: ULONG_PTR->UIntPtr
        [DllImport("user32.dll", EntryPoint = "mouse_event")]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);

        /// MOUSEEVENTF_LEFTUP -> 0x0004
        public const int MOUSEEVENTF_LEFTUP = 4;

        /// MOUSEEVENTF_RIGHTUP -> 0x0010
        public const int MOUSEEVENTF_RIGHTUP = 16;

        /// MOUSEEVENTF_MIDDLEUP -> 0x0040
        public const int MOUSEEVENTF_MIDDLEUP = 64;

        /// MOUSEEVENTF_XUP -> 0x0100
        public const int MOUSEEVENTF_XUP = 256;

        /// VK_LBUTTON -> 0x01
        public const int VK_LBUTTON = 1;

        /// VK_RBUTTON -> 0x02
        public const int VK_RBUTTON = 2;

        /// VK_MBUTTON -> 0x04
        public const int VK_MBUTTON = 4;

        /// VK_XBUTTON1 -> 0x05
        public const int VK_XBUTTON1 = 5;

        /// VK_XBUTTON2 -> 0x06
        public const int VK_XBUTTON2 = 6;

        /// Return Type: void
        ///bVk: BYTE->unsigned char
        ///bScan: BYTE->unsigned char
        ///dwFlags: DWORD->unsigned int
        ///dwExtraInfo: ULONG_PTR->UIntPtr
        [DllImport("user32.dll", EntryPoint = "keybd_event")]
        public static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);

        /// KEYEVENTF_KEYUP -> 0x0002
        public const int KEYEVENTF_KEYUP = 2;

        /// KEYEVENTF_KEYDOWN -> 0x0000
        public const int KEYEVENTF_KEYDOWN = 0;

        /// Return Type: BOOL->int
        ///lpKeyState: PBYTE->BYTE*
        [DllImport("user32.dll", EntryPoint = "GetKeyboardState")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetKeyboardState(byte[] lpKeyState);

        /// VK_CANCEL -> 0x03
        public const int VK_CANCEL = 3;

        /// VK_CAPITAL -> 0x14
        public const int VK_CAPITAL = 20;

        /// VK_NUMLOCK -> 0x90
        public const int VK_NUMLOCK = 144;

        /// VK_SCROLL -> 0x91
        public const int VK_SCROLL = 145;

        public const int VK_SHIFT = 0x10;

        /// Return Type: BOOL->int
        ///hWndParent: HWND->HWND__*
        ///hWnd: HWND->HWND__*
        [DllImport("user32.dll", EntryPoint = "IsChild")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsChild([In] IntPtr hWndParent, [In] IntPtr hWnd);

        public static bool IsSibling(IntPtr hWnd1, IntPtr hWnd2)
        {
            IntPtr sibling;

            sibling = hWnd1;
            do
            {
                sibling = NativeMethods.GetWindow(sibling, GetWindowCmd.GW_HWNDPREV);

                if (sibling == hWnd2)
                {
                    return true;
                }
            }
            while (sibling != IntPtr.Zero);


            sibling = hWnd1;
            do
            {
                sibling = NativeMethods.GetWindow(sibling, GetWindowCmd.GW_HWNDNEXT);

                if (sibling == hWnd2)
                {
                    return true;
                }
            }
            while (sibling != IntPtr.Zero);

            return false;
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        public const int VK_ESCAPE = 27;

        /// VK_CONTROL -> 0x11
        public const int VK_CONTROL = 17;

        /// VK_MENU -> 0x12
        public const int VK_MENU = 18;

        /// VK_PAUSE -> 0x13
        public const int VK_PAUSE = 19;

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///id: int
        ///fsModifiers: UINT->unsigned int
        ///vk: UINT->unsigned int
        [DllImport("user32.dll", EntryPoint = "RegisterHotKey")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool RegisterHotKey([In] IntPtr hWnd, int id, uint fsModifiers, uint vk);

        /// WM_HOTKEY -> 0x0312
        public const int WM_HOTKEY = 786;

        /// MOD_ALT -> 0x0001
        public const int MOD_ALT = 1;

        /// MOD_CONTROL -> 0x0002
        public const int MOD_CONTROL = 2;

        /// MOD_SHIFT -> 0x0004
        public const int MOD_SHIFT = 4;

        /// MOD_WIN -> 0x0008
        public const int MOD_WIN = 8;

        /// KEYEVENTF_EXTENDEDKEY -> 0x0001
        public const int KEYEVENTF_EXTENDEDKEY = 1;

        /// Return Type: int
        ///hWnd: HWND->HWND__*
        ///nBar: int
        ///nPos: int
        ///bRedraw: BOOL->intu
        [DllImport("user32.dll", EntryPoint = "SetScrollPos")]
        public static extern int SetScrollPos([In] IntPtr hWnd, int nBar, int nPos, [MarshalAs(UnmanagedType.Bool)] bool bRedraw);

        /// Return Type: int
        ///hWnd: HWND->HWND__*
        ///nBar: int
        [DllImport("user32.dll", EntryPoint = "GetScrollPos")]
        public static extern int GetScrollPos([In] IntPtr hWnd, int nBar);

        /// SBS_HORZ -> 0x0000L
        public const int SBS_HORZ = 0;

        /// SBS_VERT -> 0x0001L
        public const int SBS_VERT = 1;

        /// SB_THUMBPOSITION -> 4
        public const uint SB_THUMBPOSITION = 4;

        /// SB_LINELEFT -> 0
        public const int SB_LINELEFT = 0;

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [Flags]
        public enum ToolHelpCreateSnapshotFlags : uint
        {
            HeapList = 0x00000001,
            Process = 0x00000002,
            Thread = 0x00000004,
            Module = 0x00000008,
            Module32 = 0x00000010,
            Inherit = 0x80000000,
            All = 0x0000001F,
            NoHeaps = 0x40000000,
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct ProcessEntry32
        {
            const int MAX_PATH = 260;
            public UInt32 dwSize;
            public UInt32 cntUsage;
            public UInt32 th32ProcessID;
            public IntPtr th32DefaultHeapID;
            public UInt32 th32ModuleID;
            public UInt32 cntThreads;
            public UInt32 th32ParentProcessID;
            public Int32 pcPriClassBase;
            public UInt32 dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = MAX_PATH)]
            public string szExeFile;
        }

        [DllImport("kernel32", SetLastError = true, ExactSpelling = false, CharSet = CharSet.Unicode)]
        public static extern IntPtr CreateToolhelp32Snapshot([In] ToolHelpCreateSnapshotFlags dwFlags, [In] uint th32ProcessID);

        [DllImport("kernel32", SetLastError = true, ExactSpelling = false, CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool Process32First([In] IntPtr hSnapshot, ref ProcessEntry32 lppe);

        [DllImport("kernel32", SetLastError = true, ExactSpelling = false, CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool Process32Next([In] IntPtr hSnapshot, ref ProcessEntry32 lppe);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///hWndInsertAfter: HWND->HWND__*
        ///X: int
        ///Y: int
        ///cx: int
        ///cy: int
        ///uFlags: UINT->unsigned int
        [DllImport("user32.dll", EntryPoint = "SetWindowPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos([In] IntPtr hWnd, [In] IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, SetWindowPosFlags uFlags);

        public enum SetWindowPosFlags : uint
        {
            /// <summary>If the calling thread and the thread that owns the window are attached to different input queues, 
            /// the system posts the request to the thread that owns the window. This prevents the calling thread from 
            /// blocking its execution while other threads process the request.</summary>
            /// <remarks>SWP_ASYNCWINDOWPOS</remarks>
            AsynchronousWindowPosition = 0x4000,
            /// <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
            /// <remarks>SWP_DEFERERASE</remarks>
            DeferErase = 0x2000,
            /// <summary>Draws a frame (defined in the window's class description) around the window.</summary>
            /// <remarks>SWP_DRAWFRAME</remarks>
            DrawFrame = 0x0020,
            /// <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to 
            /// the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE 
            /// is sent only when the window's size is being changed.</summary>
            /// <remarks>SWP_FRAMECHANGED</remarks>
            FrameChanged = 0x0020,
            /// <summary>Hides the window.</summary>
            /// <remarks>SWP_HIDEWINDOW</remarks>
            HideWindow = 0x0080,
            /// <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the 
            /// top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter 
            /// parameter).</summary>
            /// <remarks>SWP_NOACTIVATE</remarks>
            DoNotActivate = 0x0010,
            /// <summary>Discards the entire contents of the client area. If this flag is not specified, the valid 
            /// contents of the client area are saved and copied back into the client area after the window is sized or 
            /// repositioned.</summary>
            /// <remarks>SWP_NOCOPYBITS</remarks>
            DoNotCopyBits = 0x0100,
            /// <summary>Retains the current position (ignores X and Y parameters).</summary>
            /// <remarks>SWP_NOMOVE</remarks>
            IgnoreMove = 0x0002,
            /// <summary>Does not change the owner window's position in the Z order.</summary>
            /// <remarks>SWP_NOOWNERZORDER</remarks>
            DoNotChangeOwnerZOrder = 0x0200,
            /// <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to 
            /// the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent 
            /// window uncovered as a result of the window being moved. When this flag is set, the application must 
            /// explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
            /// <remarks>SWP_NOREDRAW</remarks>
            DoNotRedraw = 0x0008,
            /// <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
            /// <remarks>SWP_NOREPOSITION</remarks>
            DoNotReposition = 0x0200,
            /// <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
            /// <remarks>SWP_NOSENDCHANGING</remarks>
            DoNotSendChangingEvent = 0x0400,
            /// <summary>Retains the current size (ignores the cx and cy parameters).</summary>
            /// <remarks>SWP_NOSIZE</remarks>
            IgnoreResize = 0x0001,
            /// <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
            /// <remarks>SWP_NOZORDER</remarks>
            IgnoreZOrder = 0x0004,
            /// <summary>Displays the window.</summary>
            /// <remarks>SWP_SHOWWINDOW</remarks>
            ShowWindow = 0x0040,
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MemoryBasicInformation
        {
            public UIntPtr BaseAddress;
            public UIntPtr AllocationBase;
            public MemoryProtect AllocationProtect;
            public UIntPtr RegionSize;
            public MemoryState State;
            public MemoryProtect Protect;
            public MemoryType Type;
        }

        [Flags]
        public enum MemoryState : uint
        {
            Commit = 0x1000,
            Free = 0x10000,
            Reserve = 0x2000,
        }

        [Flags]
        public enum MemoryType : uint
        {
            Image = 0x1000000,
            Mapped = 0x40000,
            Private = 0x20000,
        }

        [Flags]
        public enum MemoryProtect : uint
        {
            NoAccess = 0x01,
            ReadOnly = 0x02,
            ReadWrite = 0x04,
            WriteCopy = 0x08,
            Execute = 0x10,
            ExecuteRead = 0x20,
            ExecuteReadWrite = 0x40,
            ExecuteWriteCopy = 0x80,
            Guard = 0x100,
            NoCache = 0x200,
            WriteCombine = 0x400,
            TargetsInvalid = 0x40000000,
            TargetsNoUpdate = 0x40000000,
        }

        [DllImport("kernel32.dll", EntryPoint = "VirtualQueryEx", SetLastError = true)]
        public static extern int VirtualQueryEx(IntPtr hProcess, UIntPtr lpAddress, out MemoryBasicInformation lpBuffer, uint dwLength);

        [DllImport("psapi.dll", EntryPoint = "GetMappedFileName", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetMappedFileName(IntPtr m_hProcess, UIntPtr lpv, StringBuilder lpFilename, uint nSize);

        /// Return Type: HWND->HWND__*
        ///hWnd: HWND->HWND__*
        ///uCmd: UINT->unsigned int
        [DllImportAttribute("user32.dll", EntryPoint = "GetWindow")]
        public static extern IntPtr GetWindow([In] IntPtr hWnd, GetWindowEnum uCmd);

        public enum GetWindowEnum : uint
        {
            HWNDFIRST = 0,
            HWNDLAST = 1,
            HWNDNEXT = 2,
            HWNDPREV = 3,
            OWNER = 4,
            CHILD = 5,
            ENABLEDPOPUP = 6,
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        // Delegate to filter which windows to include 
        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool RedrawWindow(IntPtr hWnd, [In] ref tagRect lprcUpdate, IntPtr hrgnUpdate, RedrawWindowFlags flags);

        [DllImport("gdi32.dll", EntryPoint = "SelectObject")]
        public static extern IntPtr SelectObject([In] IntPtr hdc, [In] IntPtr hgdiobj);

        public enum StockObjects
        {
            WHITE_BRUSH = 0,
            LTGRAY_BRUSH = 1,
            GRAY_BRUSH = 2,
            DKGRAY_BRUSH = 3,
            BLACK_BRUSH = 4,
            NULL_BRUSH = 5,
            HOLLOW_BRUSH = NULL_BRUSH,
            WHITE_PEN = 6,
            BLACK_PEN = 7,
            NULL_PEN = 8,
            OEM_FIXED_FONT = 10,
            ANSI_FIXED_FONT = 11,
            ANSI_VAR_FONT = 12,
            SYSTEM_FONT = 13,
            DEVICE_DEFAULT_FONT = 14,
            DEFAULT_PALETTE = 15,
            SYSTEM_FIXED_FONT = 16,
            DEFAULT_GUI_FONT = 17,
            DC_BRUSH = 18,
            DC_PEN = 19,
        }

        [DllImport("gdi32.dll")]
        public static extern IntPtr GetStockObject(StockObjects fnObject);

        [DllImport("gdi32.dll")]
        public static extern IntPtr CreatePen(PenStyle fnPenStyle, int nWidth, uint crColor);

        public enum PenStyle : int
        {
            PS_SOLID = 0, //The pen is solid.
            PS_DASH = 1, //The pen is dashed.
            PS_DOT = 2, //The pen is dotted.
            PS_DASHDOT = 3, //The pen has alternating dashes and dots.
            PS_DASHDOTDOT = 4, //The pen has alternating dashes and double dots.
            PS_NULL = 5, //The pen is invisible.
            PS_INSIDEFRAME = 6,// Normally when the edge is drawn, it’s centred on the outer edge meaning that half the width of the pen is drawn
            // outside the shape’s edge, half is inside the shape’s edge. When PS_INSIDEFRAME is specified the edge is drawn 
            //completely inside the outer edge of the shape.
            PS_USERSTYLE = 7,
            PS_ALTERNATE = 8,
            PS_STYLE_MASK = 0x0000000F,

            PS_ENDCAP_ROUND = 0x00000000,
            PS_ENDCAP_SQUARE = 0x00000100,
            PS_ENDCAP_FLAT = 0x00000200,
            PS_ENDCAP_MASK = 0x00000F00,

            PS_JOIN_ROUND = 0x00000000,
            PS_JOIN_BEVEL = 0x00001000,
            PS_JOIN_MITER = 0x00002000,
            PS_JOIN_MASK = 0x0000F000,

            PS_COSMETIC = 0x00000000,
            PS_GEOMETRIC = 0x00010000,
            PS_TYPE_MASK = 0x000F0000
        };

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnregisterHotKey(IntPtr hWnd, int id);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(Point p);

        [Flags()]
        public enum ChildWindowFromPointExFlags : uint
        {
            All = 0x0,
            SKIPINVISIBLE = 0x1,
            SKIPDISABLED = 0x2,
            SKIPTRANSPARENT = 0x4,
        }

        [DllImport("user32.dll")]
        public static extern IntPtr ChildWindowFromPointEx(IntPtr hWndParent, tagPoint Point, ChildWindowFromPointExFlags uFlags);

        /// Return Type: BOOL->int
        ///hWnd: HWND->HWND__*
        ///lpPoint: LPPOINT->tagPOINT*
        [DllImport("user32.dll", EntryPoint = "ScreenToClient")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ScreenToClient([In] IntPtr hWnd, ref tagPoint lpPoint);

        public enum GetAncestorFlags
        {
            /// <summary>
            /// Retrieves the parent window. This does not include the owner, as it does with the GetParent function. 
            /// </summary>
            GetParent = 1,
            /// <summary>
            /// Retrieves the root window by walking the chain of parent windows.
            /// </summary>
            GetRoot = 2,
            /// <summary>
            /// Retrieves the owned root window by walking the chain of parent and owner windows returned by GetParent. 
            /// </summary>
            GetRootOwner = 3
        }

        [DllImport("user32.dll", ExactSpelling = true)]
        public static extern IntPtr GetAncestor(IntPtr hwnd, GetAncestorFlags flags);

        [DllImport("user32.dll", SetLastError = true, ExactSpelling = false, CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        public const uint STATUS_SUCCESS = 0;
        public const uint STATUS_TIMER_RESOLUTION_NOT_SET = 0xC0000245;

        [DllImport("ntdll.dll", SetLastError = false, ExactSpelling = false, CharSet = CharSet.Unicode)]
        public static extern uint NtQueryTimerResolution(out uint MinimumResolution, out uint MaximumResolution, out uint CurrentResolution);

        [DllImport("ntdll.dll", SetLastError = false, ExactSpelling = false, CharSet = CharSet.Unicode)]
        public static extern uint NtSetTimerResolution(uint DesiredResolution, bool SetResolution, out uint CurrentResolution);

        [DllImport("gdi32.dll")]
        private static extern GetClipBoxReturn GetClipBox(IntPtr hdc, out tagRect lprc);

        private enum GetClipBoxReturn : int
        {
            Error = 0,
            NullRegion = 1,
            SimpleRegion = 2,
            ComplexRegion = 3
        }

        //public static bool WindowObscured(IntPtr control) //This wont work when DWM is enabled (which it pretty much is everywhere now)
        //{
        //    tagRect rect;
        //    IntPtr hdc = GetDC(control);
        //    GetClipBoxReturn ret = GetClipBox(hdc, out rect);
        //    ReleaseDC(control, hdc);
        //    if (ret == GetClipBoxReturn.Error)
        //    {
        //        throw new Exception("GetClipBox failed");
        //    }
        //    if (ret == GetClipBoxReturn.NullRegion)
        //    {
        //        return true;
        //    }
        //    return false;
        //}

        public static tagRect GetClipBox(IntPtr control)
        {
            IntPtr hdc = GetWindowDC(control);
            GetClipBoxReturn ret = GetClipBox(hdc, out tagRect rect);
            ReleaseDC(control, hdc);
            if (ret == GetClipBoxReturn.Error)
            {
                throw new Exception("GetClipBox failed");
            }
            return rect;
        }

    }
}
