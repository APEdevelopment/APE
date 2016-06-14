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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//
using System.Runtime.InteropServices;
using System.IO;
using System.Diagnostics;
using NM = APE.Native.NativeMethods;

namespace APE.Spy
{
    static class Modules
    {
        public static Dictionary<string, string>Get(Process Application)
        {
            Dictionary<string, string> Modules = new Dictionary<string, string>();

            try
            {
                IntPtr ProcessHandle = Application.Handle;
                NM.MemoryBasicInformation MBI = new NM.MemoryBasicInformation();
                UIntPtr Address = UIntPtr.Zero;
                
                while (true)
                {
                    if (NM.VirtualQueryEx(ProcessHandle, Address, out MBI, (uint)Marshal.SizeOf(MBI)) == 0)
                    {
                        break;
                    }

                    if (MBI.Type.HasFlag(NM.MemoryType.Image))
                    {
                        StringBuilder FileName = new StringBuilder(2048);
                        uint ret = NM.GetMappedFileName(ProcessHandle, Address, FileName, (uint)FileName.Capacity);

                        if (ret != 0)
                        {
                            string PathAndFileName = FileName.ToString();
                            if (!Modules.Keys.Contains(PathAndFileName))
                            {
                                Modules.Add(PathAndFileName, Path.GetFileNameWithoutExtension(PathAndFileName));
                            }
                        }
                    }

                    //We need to not prefer 32bit in the build options as if we are on a 64bit system pointers need to be 64bit
                    Address = new UIntPtr(MBI.BaseAddress.ToUInt64() + MBI.RegionSize.ToUInt64());
                }
            }
            catch
            {
            }

            return Modules;
        }
    }
}
