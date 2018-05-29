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
using System.Diagnostics;

namespace APE.Native
{
    public static class NativeVersion
    {
        private static float m_KernelVersion;

        [DllImport("ntdll.dll", SetLastError = true)]
        private static extern int RtlGetVersion(ref OSVERSIONINFOEX versionInfo);

        [StructLayout(LayoutKind.Sequential)]
        private struct OSVERSIONINFOEX
        {
            // The OSVersionInfoSize field must be set to Marshal.SizeOf(typeof(OSVERSIONINFOEX))
            internal int OSVersionInfoSize;
            internal int MajorVersion;
            internal int MinorVersion;
            internal int BuildNumber;
            internal int PlatformId;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 128)]
            internal string CSDVersion;
            internal ushort ServicePackMajor;
            internal ushort ServicePackMinor;
            internal short SuiteMask;
            internal byte ProductType;
            internal byte Reserved;
        }

        static NativeVersion()
        {
            OSVERSIONINFOEX osVersionInfo = new OSVERSIONINFOEX { OSVersionInfoSize = Marshal.SizeOf(typeof(OSVERSIONINFOEX)) };
            if (RtlGetVersion(ref osVersionInfo) != 0)
            {
                throw new Exception("Failed to call RtlGetVersion");
            }
            m_KernelVersion = (float)osVersionInfo.MajorVersion + (float)osVersionInfo.MinorVersion;
        }

        public static bool IsWindowsVistaOrHigher
        {
            get
            {
                return (m_KernelVersion >= 6);
            }
        }

        public static bool IsWindows7OrHigher
        {
            get
            {
                return (m_KernelVersion >= 6.1);
            }
        }

        public static bool IsWindows8OrHigher
        {
            get
            {
                return (m_KernelVersion >= 6.2);
            }
        }

        public static bool IsWindows81OrHigher
        {
            get
            {
                return (m_KernelVersion >= 6.3);
            }
        }

        public static bool IsWindows10OrHigher
        {
            get
            {
                return (m_KernelVersion >= 10);
            }
        }
    }
}