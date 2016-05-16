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
