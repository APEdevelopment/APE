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
using System.Windows.Forms;
using System.IO;
using System.Text;
using System.Diagnostics;
using NM = APE.Native.NativeMethods;
using System.Runtime.InteropServices;

namespace APE.Spy
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length > 0)
            {
                switch (args[0])
                {
                    case "-pin":
                        PinUnpinTaskBar(true);
                        break;
                    case "-unpin":
                        PinUnpinTaskBar(false);
                        break;
                }
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new APESpy());
            }
        }

        static void PinUnpinTaskBar(bool pin)
        {
            string filePath = Process.GetCurrentProcess().MainModule.FileName;

            // 5386 is the DLL index for"Pin to Tas&kbar", ref. http://www.win7dll.info/shell32_dll.html
            int actionIndex = pin ? 5386 : 5387;
            StringBuilder szPinToStartLocalized = new StringBuilder(2048);
            IntPtr hShell32 = NM.LoadLibrary("Shell32.dll");
            if (hShell32 == IntPtr.Zero)
            {
                throw new Exception("Failed to load Shell32.dll: Last Error was " + Marshal.GetLastWin32Error().ToString());
            }

            NM.LoadString(hShell32, (uint)actionIndex, szPinToStartLocalized, szPinToStartLocalized.Capacity);
            string localizedVerb = szPinToStartLocalized.ToString();

            // create the shell application object
            dynamic shellApplication = Activator.CreateInstance(Type.GetTypeFromProgID("Shell.Application"));

            string path = Path.GetDirectoryName(filePath);
            string fileName = Path.GetFileName(filePath);

            dynamic directory = shellApplication.NameSpace(path);
            dynamic link = directory.ParseName(fileName);

            dynamic verbs = link.Verbs();
            for (int i = 0; i < verbs.Count(); i++)
            {
                dynamic verb = verbs.Item(i);

                if ((pin && verb.Name.Equals(localizedVerb)) || (!pin && verb.Name.Equals(localizedVerb)))
                {
                    verb.DoIt();
                    break;
                }
            }
        }
    }
}
