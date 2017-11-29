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
using APE.Syringe;
using Microsoft.Win32;

namespace APE.Injector
{
#if x64
    public static class Dummy64
#else
    public static class Dummy32
#endif
    {
    }

    internal static class Injector
    {
        static int Main(string[] args)
        {
            if (args.Length != 4)
            {
                Console.WriteLine("<Target PID>, <Assembly>, <Method>, <Source PID>");
                return -1;
            }
            else
            {
                //Console.WriteLine(args[0]);
                //Console.WriteLine(args[1]);
                //Console.WriteLine(args[2]);
                //Console.WriteLine(args[3]);

                int TargetPID = int.Parse(args[0]);
                string Assembly = args[1];
                string Method = args[2];
                int SourcePID = int.Parse(args[3]);

                Process p = Process.GetProcessById(TargetPID);
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software").CreateSubKey("APE", RegistryKeyPermissionCheck.Default, RegistryOptions.Volatile);
                key.SetValue(SourcePID + "_Attach_Status", "In_Process");
                int result = (int)Needle.Inject(p, SourcePID, Assembly, Method);
                key.SetValue(SourcePID + "_Attach_Status", "Success");
                return result;
            }
        }
    }
}
