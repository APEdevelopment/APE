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
using System.Text;
using System.Diagnostics;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.IO;
using NM = APE.Native.NativeMethods;
using System.Windows.Forms;

namespace APE.Syringe
{
    public static class Needle
    {
        public static uint Inject(Process processToInject, int apePid, string assembly, string method)
        {
            try
            {
                if (processToInject == null)
                {
                    throw new Exception("Null process");
                }

                if (processToInject.Id == Process.GetCurrentProcess().Id)
                {
                    throw new Exception("Cannot inject the current process");
                }

                IntPtr handleOfProcessToInject = NM.OpenProcess(NM.ProcessAccessFlags.QueryInformation | NM.ProcessAccessFlags.CreateThread |
                    NM.ProcessAccessFlags.VMOperation | NM.ProcessAccessFlags.VMWrite | NM.ProcessAccessFlags.VMRead, false, processToInject.Id);

                if (handleOfProcessToInject == IntPtr.Zero)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                if (!File.Exists(assembly))
                {
                    throw new FileNotFoundException("Cannot find file " + assembly);
                }

                string absoluteAssemblyPath = Path.GetFullPath(assembly);
                string assemblyFilename = Path.GetFileName(absoluteAssemblyPath);

                IntPtr pAssemblyPathInThisProcess = Marshal.StringToHGlobalUni(absoluteAssemblyPath);
                IntPtr pAssemblyPathInRemoteProcess = IntPtr.Zero;
                IntPtr hRemoteThread = IntPtr.Zero;
                ProcessModule moduleInRemoteProcess = null;

                try
                {
                    IntPtr hKernel32 = NM.GetModuleHandle("Kernel32");
                    if (hKernel32 == IntPtr.Zero)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    IntPtr hLoadLibrary = NM.GetProcAddress(hKernel32, "LoadLibraryW");
                    if (hLoadLibrary == IntPtr.Zero)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    uint sizeOfAssemblyPath = (uint)Encoding.Unicode.GetByteCount(absoluteAssemblyPath);

                    // Allocate memory to the remote process for absoluteAssemblyPath
                    pAssemblyPathInRemoteProcess = NM.VirtualAllocEx(handleOfProcessToInject, IntPtr.Zero,
                        sizeOfAssemblyPath, NM.AllocationType.Commit, NM.MemoryProtection.ReadWrite);
                    if (pAssemblyPathInRemoteProcess == IntPtr.Zero)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    // Write absoluteAssemblyPath to the remote process
                    int bytesWritten;
                    if (!NM.WriteProcessMemory(handleOfProcessToInject, pAssemblyPathInRemoteProcess,
                        pAssemblyPathInThisProcess, sizeOfAssemblyPath, out bytesWritten))
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    // Load dll via call to LoadLibrary using CreateRemoteThread
                    hRemoteThread = NM.CreateRemoteThread(handleOfProcessToInject, IntPtr.Zero, 0, hLoadLibrary, pAssemblyPathInRemoteProcess, 0, IntPtr.Zero);
                    if (hRemoteThread == IntPtr.Zero)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    // Wait for the thread to exit
                    if (NM.WaitForSingleObject(hRemoteThread, 30000) != NM.ThreadWaitValue.Object0)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    // Iterate modules in target process to find our newly injected module
                    processToInject.Refresh();
                    foreach (ProcessModule module in processToInject.Modules)
                    {
                        if (module.ModuleName == assemblyFilename)
                        {
                            moduleInRemoteProcess = module;
                            break;
                        }
                    }
                    if (moduleInRemoteProcess == null)
                    {
                        throw new Exception("Could not find module in remote process!");
                    }
                }
                finally
                {
                    // Tidy up
                    Marshal.FreeHGlobal(pAssemblyPathInThisProcess);
                    NM.CloseHandle(hRemoteThread);
                    NM.VirtualFreeEx(handleOfProcessToInject, pAssemblyPathInRemoteProcess, 0, NM.AllocationType.Release);
                }

                IntPtr pFunc = FindExport(moduleInRemoteProcess, method);
                IntPtr parameter = new IntPtr(apePid);

                try
                {
                    // Start a thread in the remote process calling the method and passing the APE pid as a parameter
                    hRemoteThread = NM.CreateRemoteThread(handleOfProcessToInject, IntPtr.Zero, 0, pFunc, parameter, 0, IntPtr.Zero);
                    if (hRemoteThread == IntPtr.Zero)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    // Wait for the thread to exit
                    if (NM.WaitForSingleObject(hRemoteThread, 30000) != NM.ThreadWaitValue.Object0)
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    uint exitCode;
                    if (!NM.GetExitCodeThread(hRemoteThread, out exitCode))
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }

                    return exitCode;
                }
                finally
                {
                    NM.CloseHandle(hRemoteThread);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message + "\r\n" + ex.StackTrace);
                return 1;
            }
        }

        // Loads the module as data, finds relative virtual address (RVA) of the method and uses 
        // that to find the address in the target process
        private static IntPtr FindExport(ProcessModule module, string methodName)
        {
            IntPtr hModule = IntPtr.Zero;
            try
            {
                // Load module into local process address space
                hModule = NM.LoadLibraryEx(module.FileName, IntPtr.Zero, NM.LoadLibraryExFlags.DontResolveDllReferences);
                if (hModule == IntPtr.Zero)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                // Grt the address of the function in the module locally
                IntPtr pFunc = NM.GetProcAddress(hModule, methodName);
                if (pFunc == IntPtr.Zero)
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }

                // Get RVA of export and add to base address of injected module
                IntPtr pExportAddr;
                if (IntPtr.Size == 8)
                {
                    pExportAddr = new IntPtr(module.BaseAddress.ToInt64() + (pFunc.ToInt64() - hModule.ToInt64()));
                }
                else
                {
                    pExportAddr = new IntPtr(module.BaseAddress.ToInt32() + (pFunc.ToInt32() - hModule.ToInt32()));
                }

                return pExportAddr;
            }
            finally
            {
                if (hModule.ToInt64() != 0)
                {
                    NM.FreeLibrary(hModule);
                }
            }
        }
    }
}
