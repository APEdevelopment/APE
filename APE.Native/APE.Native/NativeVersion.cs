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

namespace APE.Native
{
    public static class NativeVersion
    {
        private static float m_KernelVersion;

        static NativeVersion()
        {
            FileVersionInfo kernelFileVersion = FileVersionInfo.GetVersionInfo(Environment.GetEnvironmentVariable("SystemRoot") + @"\System32\kernel32.dll");
            string majorMinor = kernelFileVersion.FileMajorPart + "." + kernelFileVersion.FileMinorPart;
            m_KernelVersion = float.Parse(majorMinor);
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
