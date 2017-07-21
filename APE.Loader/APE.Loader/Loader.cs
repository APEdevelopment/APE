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
using RGiesecke.DllExport;  //[Un]install-Package UnmanagedExports
using System.Reflection;
using Microsoft.Win32;
using System.Diagnostics;
using System.Linq;

namespace APE.Loader
{
#if x64
    public static class Dummy64
#else
    public static class Dummy32
#endif
    {
    }

    internal static class Loader
    {
        [DllExport]
        static int LoadAPEIPC(IntPtr apePid)
        {
            try
            {
                int APEPID = apePid.ToInt32();
                string APEProcessId = apePid.ToString();
                string AUTProcessId = Process.GetCurrentProcess().Id.ToString();

                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software").CreateSubKey("APE");
                string APEPath = (string)key.GetValue(APEProcessId + "_Path_" + AUTProcessId, null);
                string AppDomainToLoadInto = (string)key.GetValue(APEProcessId + "_AppDomain_" + AUTProcessId, null);
                key.DeleteValue(APEProcessId + "_Path_" + AUTProcessId);
                key.DeleteValue(APEProcessId + "_AppDomain_" + AUTProcessId);
                key.Close();

                Assembly assembly = Assembly.LoadFrom(APEPath + @"\APE.Domain.dll");
                Assembly native = Assembly.LoadFrom(APEPath + @"\APE.Native.dll");
                Type myAPEDomainType = assembly.GetType("APE.Domain.DomainSearch");
                AppDomain appDom = null;

                if (AppDomainToLoadInto == "DefaultDomain")
                {
                    if (AppDomain.CurrentDomain.IsDefaultAppDomain())
                    {
                        appDom = AppDomain.CurrentDomain;
                    }
                }
                else
                {
                    appDom = (AppDomain)myAPEDomainType.InvokeMember("GetAppDomain", BindingFlags.Default | BindingFlags.InvokeMethod, null, null, new object[] { AppDomainToLoadInto });
                }

                if (appDom == null)
                {
                    throw new Exception("Failed to find the appdomain " + AppDomainToLoadInto);
                }

                bool WPF;
                Assembly assemblyWPF = appDom.GetAssemblies().FirstOrDefault(x => x.GetName().Name == "WindowsBase");
                if (assemblyWPF == null)
                {
                    WPF = false;
                }
                else
                {
                    WPF = true;
                }

                appDom.CreateInstanceFrom(APEPath + @"\APE.Communication.dll", "APE.Communication.APEIPC", false, BindingFlags.Default, null, new object[] { APEPID, AppDomainToLoadInto, WPF }, null, null);

                return 0;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\r\n" + e.StackTrace);
                return -1;
            }
        }
    }
}
