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
using System.Linq;
using System.Linq.Expressions;
using System.IO;
using System.Runtime.InteropServices;
using Fasterflect;  //[Un]Install-Package fasterflect
using System.Threading;
using WF = System.Windows.Forms;
using WPF = System.Windows;
using System.Reflection;
using System.Diagnostics;
using System.ComponentModel;
using APE.Domain;
using NM = APE.Native.NativeMethods;
using Microsoft.Win32;
using System.Runtime;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using APE.Bridge;
using Accessibility;

namespace APE.Communication
{
    [StructLayout(LayoutKind.Sequential, Pack = 1)]
    public struct ControlIdentifier
    {
        public IntPtr ParentHandle;
        public IntPtr Handle;
        public string Name;
        public string TechnologyType;
        public string TypeNameSpace;
        public string TypeName;
        public string ModuleName;
        public string AssemblyName;
        public int Index;
        public string Text;
        public IntPtr ChildOf;
        public IntPtr SiblingOf;
        public IntPtr ParentOf;
        public string Description;
        public string AccessibilityObjectName;
        public string UniqueId;
    }

    public enum DataStores : int
    {
        Store0 = 0,
        Store1 = 1,
        Store2 = 2,
        Store3 = 3,
        Store4 = 4,
        Store5 = 5,
        Store6 = 6,
        Store7 = 7,
        Store8 = 8,
        Store9 = 9,
    }

    public partial class APEIPC : IDisposable
    {
        private const int ERROR_ALREADY_EXISTS = 183;
        internal const int StringSpaceBytes = 10485760; //1024 * 1024 * 10

        //Memory Map File for IPC for message strings
        private IntPtr m_HandleMemoryMappedFileStringStore;
        internal IntPtr m_IntPtrMemoryMappedFileViewStringStore;
        internal int m_StringStoreOffset;

        //Memory Map File for IPC for message value types
        private IntPtr m_HandleMemoryMappedFileMessageStore;
        private IntPtr m_IntPtrMemoryMappedFileViewMessageStore;
        private unsafe MessageStore* m_PtrMessageStore;

        private unsafe int m_SizeOfMessage = sizeof(Message);
        private ParametersTypeCircularList ParametersTypeCache = new ParametersTypeCircularList(3);    //3 is optimal here as there are only a couple of very frequently used types
        private ParameterTypeCircularList ParameterTypeCache = new ParameterTypeCircularList(10);
        private MemberGetterCircularList MemberGetterCache = new MemberGetterCircularList(10);
        private MethodInvokerCircularList MethodInvokerCache = new MethodInvokerCircularList(10);
        private ConstructorInvokerCircularList ConstructorInvokerCache = new ConstructorInvokerCircularList(10);

        //Cache all the primitive types (for performance)
        private Type m_TypeBoolean = typeof(Boolean);
        private Type m_TypeBooleanByRef = typeof(Boolean).MakeByRefType();
        private Type m_TypeChar = typeof(Char);
        private Type m_TypeCharByRef = typeof(Char).MakeByRefType();
        private Type m_TypeSByte = typeof(SByte);
        private Type m_TypeSByteByRef = typeof(SByte).MakeByRefType();
        private Type m_TypeByte = typeof(Byte);
        private Type m_TypeByteByRef = typeof(Byte).MakeByRefType();
        private Type m_TypeInt16 = typeof(Int16);
        private Type m_TypeInt16ByRef = typeof(Int16).MakeByRefType();
        private Type m_TypeUInt16 = typeof(UInt16);
        private Type m_TypeUInt16ByRef = typeof(UInt16).MakeByRefType();
        private Type m_TypeInt32 = typeof(Int32);
        private Type m_TypeInt32ByRef = typeof(Int32).MakeByRefType();
        private Type m_TypeUInt32 = typeof(UInt32);
        private Type m_TypeUInt32ByRef = typeof(UInt32).MakeByRefType();
        private Type m_TypeInt64 = typeof(Int64);
        private Type m_TypeInt64ByRef = typeof(Int64).MakeByRefType();
        private Type m_TypeUInt64 = typeof(UInt64);
        private Type m_TypeUInt64ByRef = typeof(UInt64).MakeByRefType();
        private Type m_TypeSingle = typeof(Single);
        private Type m_TypeSingleByRef = typeof(Single).MakeByRefType();
        private Type m_TypeDouble = typeof(Double);
        private Type m_TypeDoubleByRef = typeof(Double).MakeByRefType();
        private Type m_TypeDecimal = typeof(Decimal);
        private Type m_TypeDecimalByRef = typeof(Decimal).MakeByRefType();
        private Type m_TypeDateTime = typeof(DateTime);
        private Type m_TypeDateTimeByRef = typeof(DateTime).MakeByRefType();
        private Type m_TypeString = typeof(String);
        private Type m_TypeStringByRef = typeof(String).MakeByRefType();
        private Type m_TypeIntPtr = typeof(IntPtr);
        private Type m_TypeIntPtrByRef = typeof(IntPtr).MakeByRefType();

        private object tempStore0;
        private bool tempStore0ReleaseComObject = false;
        private object tempStore1;
        private bool tempStore1ReleaseComObject = false;
        private object tempStore2;
        private bool tempStore2ReleaseComObject = false;
        private object tempStore3;
        private bool tempStore3ReleaseComObject = false;
        private object tempStore4;
        private bool tempStore4ReleaseComObject = false;
        private object tempStore5;
        private bool tempStore5ReleaseComObject = false;
        private object tempStore6;
        private bool tempStore6ReleaseComObject = false;
        private object tempStore7;
        private bool tempStore7ReleaseComObject = false;
        private object tempStore8;
        private bool tempStore8ReleaseComObject = false;
        private object tempStore9;
        private bool tempStore9ReleaseComObject = false;

        private Boolean m_DoneFind;
        private Boolean m_DoneQuery;
        private Boolean m_DoneGet;

        EventWaitHandle m_eventIPC;
        IntPtr m_Handle;
        string m_Name;
        string m_Text;

        private bool m_WPF = false;
        private bool m_AssemblyLoadEventHandlerActive = false;

        private delegate string GetTextDelegate(WF.Control theControl);
        private GetTextDelegate m_GetTextDelegater;
        private delegate string GetAccessibilityObjectNameDelegate(WF.Control theControl);
        private GetAccessibilityObjectNameDelegate m_GetAccessibilityObjectNameDelegater;
        private delegate object ConvertTypeDelegate(Type theTyoe, object theObject);
        private ConvertTypeDelegate m_ConvertTypeDelegater;
        private delegate void GetWPFHandleAndNameAndTitleDelegate(WPF.Window theWindow);
        private GetWPFHandleAndNameAndTitleDelegate m_GetWPFHandleAndNameAndTitleDelegater;
        private NM.EnumWindow EnumThreadProcedue;
        private int m_ManagedThreadId;

        List<IntPtr> m_AllControls;
        EventSet Side;
        Process ApeProcess = null;
        public Process AUTProcess = null;
        public string AUTProcessName = null;
        public string AUTProcessId = null;
        bool m_Abort = false;
        uint m_TimeOut = 0;
        uint m_HangTimeOut = 30000;
        //int m_Sleep = 62;               //sleep time ms best to use one of the following values rounded down (or a multiple)
        //15.625
        //31.25
        //46.875
        //62.5
        //78.125
        //93.75
        //109.375
        //125

        private int m_MessageNumber = 0;

        [Flags]
        public enum MessageResult : int
        {
            Success = 1,
            Failure = 2,
        }

        private void ApeAssemblyLoadEventHandler(object sender, AssemblyLoadEventArgs args)
        {
            if (args.LoadedAssembly.GetName().Name == "WindowsBase")
            {
                m_WPF = true;
            }
        }

        private void BuildListOfAllControls(IntPtr parentWindow)
        {
            IntPtr child = NM.GetWindow(parentWindow, NM.GetWindowEnum.CHILD);

            if (child != IntPtr.Zero)
            {
                if (NM.IsWindowVisible(child))
                {
                    m_AllControls.Add(child);
                    BuildListOfAllControls(child);
                }

                IntPtr sibling = child;
                do
                {
                    sibling = NM.GetWindow(sibling, NM.GetWindowEnum.HWNDNEXT);
                    if (sibling != IntPtr.Zero)
                    {
                        if (NM.IsWindowVisible(sibling))
                        {
                            m_AllControls.Add(sibling);
                            BuildListOfAllControls(sibling);
                        }
                    }
                }
                while (sibling != IntPtr.Zero);
            }
        }

        private bool EnumThreadCallback(IntPtr hWnd, IntPtr lParam)
        {
            if (NM.IsWindowVisible(hWnd))
            {
                m_AllControls.Add(hWnd);
            }
            return true;
        }

        public unsafe APEIPC(Process AUTProcess)
        {
            InjectAPEIPC(AUTProcess, "DefaultDomain");
        }

        public unsafe APEIPC(Process AUTProcess, string AppDomain)
        {
            InjectAPEIPC(AUTProcess, AppDomain);
        }

        private unsafe void InjectAPEIPC(Process autProcess, string appDomain)
        {
            m_ManagedThreadId = Thread.CurrentThread.ManagedThreadId;
            this.AUTProcess = autProcess;
            this.AUTProcessId = AUTProcess.Id.ToString();
            this.AUTProcessName = AUTProcess.ProcessName;

            string path = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string method = "LoadAPEIPC";
            string apeProcessId = Process.GetCurrentProcess().Id.ToString();
            
            m_HandleMemoryMappedFileStringStore = NM.CreateFileMapping((IntPtr)(NM.INVALID_HANDLE_VALUE), (IntPtr)0, NM.FileMapProtection.PageReadWrite, 0, StringSpaceBytes, apeProcessId + "_String_" + appDomain + "_" + AUTProcessId);

            if (m_HandleMemoryMappedFileStringStore == null)
            {
                throw new Exception("Failed to create memory mapped file error: " + Marshal.GetLastWin32Error().ToString());
            }

            if (Marshal.GetLastWin32Error() != ERROR_ALREADY_EXISTS)
            {
                int exitCode = -1;
                Process Injector = null;
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software").CreateSubKey("APE", RegistryKeyPermissionCheck.Default, RegistryOptions.Volatile);
                string tempPath = null;
                try
                {
                    key.SetValue(apeProcessId + "_Path_" + AUTProcessId, path);
                    key.SetValue(apeProcessId + "_AppDomain_" + AUTProcessId, appDomain);
                    key.SetValue(apeProcessId + "_Attach_Status", "Starting");
                    
                    string assembly;
                    string exe;
                    if (NM.Is32BitProcess(autProcess))
                    {
                        assembly = "APE.Loader_x86.dll";
                        exe = "APE.Injector_x86.exe";
                    }
                    else
                    {
                        assembly = "APE.Loader_x64.dll";
                        exe = "APE.Injector_x64.exe";
                    }

                    tempPath = Path.GetTempPath();
                    File.Delete(tempPath + @"devenv.exe");
                    File.Delete(tempPath + @"APE.Syringe.dll");
                    File.Delete(tempPath + @"APE.Native.dll");
                    File.Copy(path + @"\" + exe, tempPath + @"devenv.exe");
                    File.Copy(path + @"\APE.Syringe.dll", tempPath + @"APE.Syringe.dll");
                    File.Copy(path + @"\APE.Native.dll", tempPath + @"APE.Native.dll");

                    Injector = new Process();
                    Injector.StartInfo = new ProcessStartInfo(tempPath + @"devenv.exe", AUTProcessId + " " + assembly + " " + method + " " + apeProcessId);
                    Injector.StartInfo.WorkingDirectory = path;
                    Injector.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    Injector.Start();

                    Injector.WaitForExit();
                    exitCode = Injector.ExitCode;
                    if (exitCode != 0)
                    {
                        throw new Exception("Failed to attach to process: exit code: " + exitCode.ToString());
                    }
                    string state = (string)key.GetValue(apeProcessId + "_Attach_Status", "");
                    if (state != "Success")
                    {
                        string additionalInformation = "";
                        if (state == "In_Process")
                        {
                            additionalInformation = Environment.NewLine + "Please check any virus scanning software for possible error";
                        }
                        throw new Exception("Failed to attach to process: last state: " + state + additionalInformation);
                    }
                }
                finally
                {
                    if (Injector != null)
                    {
                        Injector.Dispose();
                    }
                    key.DeleteValue(apeProcessId + "_Attach_Status", false);
                    key.DeleteValue(apeProcessId + "_Path_" + AUTProcessId, false);
                    key.DeleteValue(apeProcessId + "_AppDomain_" + AUTProcessId, false);
                    key.Dispose();
                    if (tempPath != null)
                    {
                        try
                        {
                            File.Delete(tempPath + @"devenv.exe");
                            File.Delete(tempPath + @"APE.Syringe.dll");
                            File.Delete(tempPath + @"APE.Native.dll");
                        }
                        catch   //Shouldn't be needed but worth an attempt if it ever does
                        {
                            Thread.Sleep(1000);
                            File.Delete(tempPath + @"devenv.exe");
                            File.Delete(tempPath + @"APE.Syringe.dll");
                            File.Delete(tempPath + @"APE.Native.dll");
                        }
                    }
                }

                if (exitCode != 0)
                {
                    throw new Exception("Injector exited with status: " + Injector.ExitCode.ToString());
                }
            }

            m_IntPtrMemoryMappedFileViewStringStore = NM.MapViewOfFile(m_HandleMemoryMappedFileStringStore, NM.FileMapAccess.FileMapAllAccess, 0, 0, (UIntPtr)StringSpaceBytes);
            m_HandleMemoryMappedFileMessageStore = NM.CreateFileMapping((IntPtr)(NM.INVALID_HANDLE_VALUE), (IntPtr)0, NM.FileMapProtection.PageReadWrite, 0, (uint)sizeof(MessageStore), apeProcessId + "_Message_" + appDomain + "_" + AUTProcessId);
            m_IntPtrMemoryMappedFileViewMessageStore = NM.MapViewOfFile(m_HandleMemoryMappedFileMessageStore, NM.FileMapAccess.FileMapAllAccess, 0, 0, (UIntPtr)sizeof(MessageStore));
            m_PtrMessageStore = (MessageStore*)m_IntPtrMemoryMappedFileViewMessageStore.ToPointer();

            m_eventIPC = new EventWaitHandle(false, EventResetMode.AutoReset, apeProcessId + "_EventIPC_" + appDomain + "_" + AUTProcessId);
            Side = EventSet.APE;
            m_Abort = false;
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public unsafe APEIPC(int APEPID, string AppDomainToLoadInto)
        {
            try
            {
                AppDomain.CurrentDomain.AssemblyLoad += new AssemblyLoadEventHandler(ApeAssemblyLoadEventHandler);
                m_AssemblyLoadEventHandlerActive = true;

                Assembly assemblyWPF = AppDomain.CurrentDomain.GetAssemblies().FirstOrDefault(x => x.GetName().Name == "WindowsBase");
                if (assemblyWPF == null)
                {
                    m_WPF = false;
                }
                else
                {
                    m_WPF = true;
                }

                Thread myThread = new Thread(() => ProcessMessages(APEPID, AppDomainToLoadInto));
                myThread.SetApartmentState(ApartmentState.STA);
                myThread.IsBackground = true;
                myThread.Start();
            }
            catch (Exception ex)
            {
                TextWriter log = File.AppendText(Environment.GetEnvironmentVariable("TEMP") + @"\APE_Critical.log");
                log.WriteLine(DateTime.Now.ToString() + "\t" + ex.Message);
                log.WriteLine(DateTime.Now.ToString() + "\t" + ex.StackTrace);
                log.Close();
                throw;
            }
        }

        ~APEIPC()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (m_IntPtrMemoryMappedFileViewStringStore != IntPtr.Zero)
            {
                NM.UnmapViewOfFile(m_IntPtrMemoryMappedFileViewStringStore);
                m_IntPtrMemoryMappedFileViewStringStore = IntPtr.Zero;
                NM.CloseHandle(m_HandleMemoryMappedFileStringStore);
                m_HandleMemoryMappedFileStringStore = IntPtr.Zero;
            }

            if (m_IntPtrMemoryMappedFileViewMessageStore != IntPtr.Zero)
            {
                NM.UnmapViewOfFile(m_IntPtrMemoryMappedFileViewMessageStore);
                m_IntPtrMemoryMappedFileViewMessageStore = IntPtr.Zero;
                NM.CloseHandle(m_HandleMemoryMappedFileMessageStore);
                m_HandleMemoryMappedFileMessageStore = IntPtr.Zero;
            }

            if (disposing)
            {
                if (m_eventIPC != null)
                {
                    m_eventIPC.Dispose();
                }

                if (m_AssemblyLoadEventHandlerActive)
                {
                    AppDomain.CurrentDomain.AssemblyLoad -= new AssemblyLoadEventHandler(ApeAssemblyLoadEventHandler);
                    m_AssemblyLoadEventHandlerActive = false;
                }
            }
        }

        private unsafe void GetAppDomains(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("GetAppDomains must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            string AppDomains = DomainSearch.GetAllAppDomainNames();

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            AddReturnValue(new Parameter(this, AppDomains));
        }

        private unsafe void SetTimeOuts(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("SetTimeOuts must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = TimeOut
            if ((PtrMessage->Parameter.TypeCode[0]) == (Int32)ApeTypeCode.UInt32)
            {
                m_TimeOut = PtrMessage->Parameter.UInt32[0];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.UInt32 got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        private MethodInfo MethodInfoGetAllRepositories = null;
        private void FlushAllLog4NetLogs()
        {
            if (MethodInfoGetAllRepositories == null)
            {
                Assembly[] allAssemblies = AppDomain.CurrentDomain.GetAssemblies();
                for (int i = 0; i < allAssemblies.Count(); i++)
                {
                    if (allAssemblies[i].FullName.StartsWith("log4net, "))
                    {
                        Type typeLogManager = allAssemblies[i].GetType("log4net.LogManager");
                        if (typeLogManager != null)
                        {
                            MethodInfoGetAllRepositories = typeLogManager.GetMethod("GetAllRepositories", BindingFlags.Public | BindingFlags.Static);
                        }
                        break;
                    }
                }
            }

            if (MethodInfoGetAllRepositories != null)
            {
                dynamic repositories = MethodInfoGetAllRepositories.Invoke(null, null);
                foreach (dynamic repository in repositories)
                {
                    try
                    {
                        repository.Flush((int)TimeOut);
                    }
                    catch
                    {
                    }
                }
            }
        }

        private unsafe void GarbageCollect(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GarbageCollect must be first message");
            }

            int generation = GetParameterInt32(ptrMessage, 0);
            CleanUpMessage(ptrMessage);

            // Make sure the the logs are all flushed
            //FlushAllLog4NetLogs();

            //Tempoary workaround for painting issues with GC
            List<IntPtr> formList = new List<IntPtr>();
            for (int index = WF.Application.OpenForms.Count - 1; index > -1; index--)
            {
                try
                {
                    WF.Form form = WF.Application.OpenForms[index];
                    GetHandleAndName(form);
                    if (NM.IsWindowVisible(m_Handle))
                    {
                        formList.Add(m_Handle);
                        NM.SuspendDrawing(m_Handle, TimeOut);
                    }
                }
                catch
                {
                }
            }

            // Do a single GC (more than this can have undesirable effects on performance)
            if (generation == GC.MaxGeneration)
            {
                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
                GC.Collect(generation, GCCollectionMode.Forced, true);
            }
            else
            {
                GC.Collect(generation);
            }
            GC.WaitForPendingFinalizers();

            //Tempoary workaround for painting issues with GC
            foreach (IntPtr form in formList)
            {
                try
                {
                    NM.ResumeDrawing(form, TimeOut);
                }
                catch
                {
                }
            }

            Thread.Sleep(150);  //A small sleep after GC seems to make the performance timings more accurate
        }

        private unsafe void VisualStyleSupported(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("VisualStyleSupported must be first message");
            }

            CleanUpMessage(ptrMessage);
            AddReturnValue(new Parameter(this, WF.VisualStyles.VisualStyleRenderer.IsSupported));
        }

        public unsafe void GetContextMenuStrip(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("GetContextMenuStrip must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = handle
            IntPtr Handle;
            if ((PtrMessage->Parameter.TypeCode[0]) == (int)ApeTypeCode.IntPtr)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            m_AllControls = new List<IntPtr>();
            NM.EnumThreadWindows((uint)NM.GetWindowThreadProcessId(Handle, IntPtr.Zero), EnumThreadProcedue, IntPtr.Zero);

            IntPtr contextMenuStrip = IntPtr.Zero;
            foreach (IntPtr hWnd in m_AllControls)
            {
                if (NM.IsWindowVisible(hWnd))
                {
                    WF.Control FoundControl = WF.Control.FromHandle(hWnd);
                    if (FoundControl != null)
                    {
                        if (FoundControl.GetType().Name == "ContextMenuStrip")
                        {
                            contextMenuStrip = hWnd;
                            break;
                        }
                    }
                }
            }
            AddReturnValue(new Parameter(this, contextMenuStrip));
        }

        public unsafe void GetListViewItemRectangle(int MessageNumber)
        {
            //
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("GetListViewItemRectangle must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = handle
            IntPtr Handle;
            if ((PtrMessage->Parameter.TypeCode[0]) == (int)ApeTypeCode.IntPtr)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = Index
            int Index;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)ApeTypeCode.Int32)
            {
                Index = PtrMessage->Parameter.Int32[1];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.Int32 got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
            }

            NM.tagRect theRect = new NM.tagRect();
            theRect.left = NM.LVIR_LABEL;
            IntPtr Result;
            IntPtr Return;
            Return = NM.SendMessageTimeout(Handle, NM.LVM_GETITEMRECT, new IntPtr(Index), ref theRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out Result);
            if (Return == IntPtr.Zero || Result == IntPtr.Zero)
            {
                throw new Exception("Failed to get the rect of the item");
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            AddReturnValue(new Parameter(this, theRect.top));
            AddReturnValue(new Parameter(this, theRect.left));
            AddReturnValue(new Parameter(this, theRect.bottom));
            AddReturnValue(new Parameter(this, theRect.right));
        }

        public unsafe void GetListViewGroupRectangle(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("GetListViewGroupRectangle must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = handle
            IntPtr Handle;
            if ((PtrMessage->Parameter.TypeCode[0]) == (int)ApeTypeCode.IntPtr)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = Group
            int GroupID;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)ApeTypeCode.Int32)
            {
                GroupID = PtrMessage->Parameter.Int32[1];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.Int32 got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
            }

            NM.tagRect theRect = new NM.tagRect();
            theRect.top = NM.LVGGR_LABEL;
            IntPtr Result;
            IntPtr Return;
            Return = NM.SendMessageTimeout(Handle, NM.LVM_GETGROUPRECT, new IntPtr(GroupID), ref theRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out Result);
            if (Return == IntPtr.Zero || Result == IntPtr.Zero)
            {
                throw new Exception("Failed to get the rect of the group");
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            AddReturnValue(new Parameter(this, theRect.top));
            AddReturnValue(new Parameter(this, theRect.left));
            AddReturnValue(new Parameter(this, theRect.bottom));
            AddReturnValue(new Parameter(this, theRect.right));
        }

        public unsafe void RemoveFileMapping()
        {
            Dispose(true);
        }

        unsafe private void GetResultMessage()
        {
            string ResultMessage;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((m_PtrMessageStore->NumberOfMessages - 1) * m_SizeOfMessage));

            MessageResult Result = (MessageResult)PtrMessage->Parameter.Int32[0];

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            if (Result != MessageResult.Success)
            {
                if (PtrMessage->Parameter.StringLength[1] == -1)
                {
                    ResultMessage = null;
                }
                else
                {
                    ResultMessage = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[1]), 0, PtrMessage->Parameter.StringLength[1]);
                }

                throw new Exception(ResultMessage);
            }
        }

        unsafe public dynamic GetValueFromMessage()
        {
            if (m_Abort == true)
            {
                if (AUTProcess == null)
                {
                    throw new Exception("Not attached to a process");
                }
                if (AUTProcessHasExited())
                {
                    throw new Exception(AUTProcess.ProcessName + " has exited");
                }
            }

            m_MessageNumber++;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((m_MessageNumber - 1) * m_SizeOfMessage));
            int parameter = 0;
            dynamic returnValue;

            switch (PtrMessage->Parameter.TypeCode[parameter])
            {
                case (int)ApeTypeCode.Boolean:
                    returnValue = PtrMessage->Parameter.Boolean[parameter];
                    break;
                case (int)ApeTypeCode.Char:
                    returnValue = PtrMessage->Parameter.Char[parameter];
                    break;
                case (int)ApeTypeCode.SByte:
                    returnValue = PtrMessage->Parameter.SByte[parameter];
                    break;
                case (int)ApeTypeCode.Byte:
                    returnValue = PtrMessage->Parameter.Byte[parameter];
                    break;
                case (int)ApeTypeCode.Int16:
                    returnValue = PtrMessage->Parameter.Int16[parameter];
                    break;
                case (int)ApeTypeCode.UInt16:
                    returnValue = PtrMessage->Parameter.UInt16[parameter];
                    break;
                case (int)ApeTypeCode.Int32:
                    returnValue = PtrMessage->Parameter.Int32[parameter];
                    break;
                case (int)ApeTypeCode.UInt32:
                    returnValue = PtrMessage->Parameter.UInt32[parameter];
                    break;
                case (int)ApeTypeCode.Int64:
                    returnValue = PtrMessage->Parameter.Int64[parameter];
                    break;
                case (int)ApeTypeCode.UInt64:
                    returnValue = PtrMessage->Parameter.UInt64[parameter];
                    break;
                case (int)ApeTypeCode.Single:
                    returnValue = PtrMessage->Parameter.Single[parameter];
                    break;
                case (int)ApeTypeCode.Double:
                    returnValue = PtrMessage->Parameter.Double[parameter];
                    break;
                case (int)ApeTypeCode.Decimal:
                    int[] DecimalBits = new int[4];
                    DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[parameter];
                    DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[parameter];
                    DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[parameter];
                    DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[parameter];
                    returnValue = new decimal(DecimalBits);
                    break;
                case (int)ApeTypeCode.DateTime:
                    returnValue = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[parameter]);
                    break;
                case (int)ApeTypeCode.String:
                    if (PtrMessage->Parameter.StringLength[parameter] == -1)
                    {
                        string Empty = null;
                        returnValue = Empty;
                    }
                    else
                    {
                        returnValue = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[parameter]), 0, PtrMessage->Parameter.StringLength[parameter]);
                    }
                    break;
                case (int)ApeTypeCode.IntPtr:
                    returnValue = new IntPtr(PtrMessage->Parameter.IntPtr[parameter]);
                    break;
                case (int)ApeTypeCode.Empty:
                    returnValue = null;
                    break;
                case (int)ApeTypeCode.Image:
                    if (PtrMessage->Parameter.StringLength[parameter] == -1)
                    {
                        Image Empty = null;
                        returnValue = Empty;
                    }
                    else
                    {
                        byte[] imageAsByteArray = new byte[PtrMessage->Parameter.StringLength[parameter]];
                        fixed (void* PtrByteArray = imageAsByteArray)
                        {
                            NM.CopyMemory((IntPtr)PtrByteArray, m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[parameter], (UIntPtr)(imageAsByteArray.Length));
                        }

                        returnValue = Image.FromStream(new MemoryStream(imageAsByteArray));
                    }
                    break;
                default:
                    throw new Exception("Unsupported ApeTypeCode: " + PtrMessage->Parameter.TypeCode[parameter].ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            return returnValue;
        }

        //TODO is it worth doing to avoid dynamic for performance?
        //unsafe public void GetValueFromMessage(int messageNumber, out string value)
        //{
        //    Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
        //    int parameter = 0;

        //    if (PtrMessage->Parameter.StringLength[parameter] == -1)
        //    {
        //        string Empty = null;
        //        value = Empty;
        //    }
        //    else
        //    {
        //        value = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[parameter]), 0, PtrMessage->Parameter.StringLength[parameter]);
        //    }

        //    //cleanup the message
        //    PtrMessage->TypeCodeKey = 0;
        //    PtrMessage->NumberOfParameters = 0;
        //    PtrMessage->NameOffset = 0;
        //    PtrMessage->NameLength = 0;
        //    PtrMessage->Action = MessageAction.None;
        //}

        unsafe public void SendMessages(EventSet WhoIsSending)
        {
            if (Thread.CurrentThread.ManagedThreadId != m_ManagedThreadId)
            {
                throw new InvalidOperationException("Cross-thread operation not valid");
            }

            //TODO check if m_Abort is set as if it is probably don't actually want to try send the messages

            //signal the other process
            m_eventIPC.Set();
            Stopwatch timer = Stopwatch.StartNew();

            // Yield the current process until the other process wakes up and notifies us via the mmf

            // 10 million loops should take a couple of seconds at most
            for (int loop = 0; loop < 10000000; loop++)
            {
                Thread.Yield();

                if (m_PtrMessageStore->LastWake != WhoIsSending)
                {
                    return;
                }
            }

            // If the above loop didn't handle it then fall back to a less aggressive loop
            // This should never happen under normal circumstances
            while (true)
            {
                // Checking if a process has exited is slow so we only check once every second or so
                if (WhoIsSending == EventSet.APE)
                {
                    // Check if the AUT is still running
                    if (AUTProcess == null)
                    {
                        m_Abort = true;
                        throw new Exception("Not attached to a process");
                    }

                    if (AUTProcessHasExited())
                    {
                        m_Abort = true;
                        throw new Exception(AUTProcess.ProcessName + " has exited");
                    }                    

                    //if last awake hasn't been set in m_TimeOut + m_HangTimeout ms then report a hang (might happen if the process still exists but is showing a 'stopped working dialog')
                    if (timer.ElapsedMilliseconds > m_TimeOut + m_HangTimeOut)
                    {
                        m_Abort = true;
                        throw new Exception(AUTProcess.ProcessName + " has hung");
                    }
                }
                else
                {
                    // Check if the APE is still running (if it crashed it wont ask us to unload)
                    // If its being debugged HasExited will throw a permission denined exception
                    try
                    {
                        if (ApeProcess == null || ApeProcess.HasExited)
                        {
                            m_Abort = true;
                            return;
                        }
                    }
                    catch
                    {
                    }
                }

                // Loop for about a second
                for (int loop = 0; loop < 1000; loop++)
                {
                    Thread.Sleep(1);

                    if (m_PtrMessageStore->LastWake != WhoIsSending)
                    {
                        return;
                    }
                }
            }
        }

        public bool AUTProcessHasExited()
        {
            // If its being debugged HasExited will throw a permission denined exception
            try
            {
                if (AUTProcess.HasExited)
                {
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        unsafe public void WaitForMessages(EventSet WhoIsWaiting)
        {
            int waitTimeMs;
            Process processToCheck;
            bool signaled;
            Stopwatch timer = null;

            if (Side == EventSet.APE)
            {
                waitTimeMs = 1000;
                processToCheck = AUTProcess;
                timer = Stopwatch.StartNew();
            }
            else
            {
                waitTimeMs = 30000;
                processToCheck = ApeProcess;
            }

            while (true)
            {
                signaled = m_eventIPC.WaitOne(waitTimeMs);

                // Check if the process is still running
                // If its being debugged HasExited will throw a permission denied exception
                try
                {
                    if (processToCheck == null || processToCheck.HasExited)
                    {
                        // We may be expecting it to have exited (IE if the form was closed)
                        m_Abort = true;
                        break;
                    }
                }
                catch
                {
                }

                if (signaled)
                {
                    //writes to the mmf to let the other process know we have woken up
                    m_PtrMessageStore->LastWake = WhoIsWaiting;
                    if (Side == EventSet.APE)
                    {
                        GetResultMessage();
                    }
                    break;
                }

                //if a response hasn't been recieved in m_TimeOut + m_HangTimeout ms then report a hang
                if (Side == EventSet.APE)
                {
                    if (timer.ElapsedMilliseconds > m_TimeOut + m_HangTimeOut)
                    {
                        m_Abort = true;
                        throw new Exception(processToCheck.ProcessName + " has hung");
                    }
                }
            }            
        }

        unsafe public void AddRetrieveMessageGetValue(DataStores SourceStore)
        {
            m_MessageNumber = 0;

            if (m_DoneFind == false)
            {
                throw new Exception("Need to find the control before getting a value");
            }

            if (m_DoneQuery == false)
            {
                throw new Exception("Need to query a control before getting a value");
            }

            //add a result request
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetResult;
            PtrMessage->SourceStore = SourceStore;
            PtrMessage->DestinationStore = SourceStore;
            PtrMessage->NameOffset = -1;
            PtrMessage->NameLength = -1;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneGet = true;
        }

        //TODO Is it worth doing this, it would avoid using dynamic for improved performance
        //unsafe public void AddRetrieveMessageGetValue(DataStores SourceStore, int TypeCode)
        //{
        //    m_MessageNumber = 0;

        //    if (m_DoneFind == false)
        //    {
        //        throw new Exception("Need to find the control before getting a value");
        //    }

        //    if (m_DoneQuery == false)
        //    {
        //        throw new Exception("Need to query a control before getting a value");
        //    }

        //    //add a result request
        //    Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

        //    PtrMessage->Action = MessageAction.GetResult;

        //    PtrMessage->SourceStore = SourceStore;
        //    PtrMessage->DestinationStore = SourceStore;
        //    PtrMessage->NameOffset = -1;
        //    PtrMessage->NameLength = -1;

        //    PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = TypeCode;
        //    PtrMessage->NumberOfParameters++;

        //    m_PtrMessageStore->NumberOfMessages++;
        //    m_DoneGet = true;
        //}

        unsafe public void AddFirstMessageControlExistsByProperty(ControlIdentifier Identifier)
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->Action = MessageAction.ControlExists;

            AddIdentifierMessage(Identifier);
        }

        unsafe public void AddFirstMessageFindByProperty(ControlIdentifier Identifier)
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->Action = MessageAction.Find;

            AddIdentifierMessage(Identifier);
        }

        unsafe public void AddMessageGetApeTypeFromType(DataStores SourceStore, DataStores DestinationStore)
        {
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->SourceStore = SourceStore;
            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.GetApeTypeFromType;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
        }

        unsafe public void AddMessageGetApeTypeFromObject(DataStores SourceStore, DataStores DestinationStore)
        {
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->SourceStore = SourceStore;
            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.GetApeTypeFromObject;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
        }

        unsafe public void AddMessageGetRecognisedType(DataStores SourceStore, DataStores DestinationStore)
        {
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->SourceStore = SourceStore;
            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.GetRecognisedType;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
        }

        unsafe private void AddFirstMessageSetTimeOuts()
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->Action = MessageAction.SetTimeOuts;

            //p1
            Parameter p = new Parameter(this, TimeOut);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddResultMessage(MessageResult Result)
        {
            Message* PtrMessage = (Message*)(this.m_IntPtrMemoryMappedFileViewMessageStore + (this.m_PtrMessageStore->NumberOfMessages * this.m_SizeOfMessage));
            DebugLogging.WriteLog("setting result message " + m_PtrMessageStore->NumberOfMessages.ToString() + " to " + Result.ToString() + " parameter " + PtrMessage->NumberOfParameters.ToString());

            // This should be the last message returned
            Parameter p;

            //p1
            p = new Parameter(this, (int)Result);

            m_PtrMessageStore->NumberOfMessages++;
        }

        unsafe public void AddResultMessage(MessageResult Result, string Message)
        {
            DebugLogging.WriteLog("setting result message " + m_PtrMessageStore->NumberOfMessages.ToString() + " to " + Result.ToString());

            // This should be the last message returned
            Parameter p;

            //p1
            p = new Parameter(this, (int)Result);

            //p2
            p = new Parameter(this, Message);

            m_PtrMessageStore->NumberOfMessages++;
        }

        unsafe public void AddIdentifierMessage(ControlIdentifier Identifier)
        {
            Parameter p;

            //p1
            p = new Parameter(this, Identifier.ParentHandle);

            //p2
            p = new Parameter(this, Identifier.Handle);

            //p3
            p = new Parameter(this, Identifier.Name);

            //p4
            p = new Parameter(this, Identifier.TechnologyType);

            //p5
            p = new Parameter(this, Identifier.TypeNameSpace);

            //p6
            p = new Parameter(this, Identifier.TypeName);

            //p7
            p = new Parameter(this, Identifier.ModuleName);

            //p8
            p = new Parameter(this, Identifier.AssemblyName);

            //p9
            p = new Parameter(this, Identifier.Index);

            //p10
            p = new Parameter(this, Identifier.Text);

            //p11
            p = new Parameter(this, Identifier.ChildOf);

            //p12
            p = new Parameter(this, Identifier.SiblingOf);

            //p13
            p = new Parameter(this, Identifier.ParentOf);

            //p14
            p = new Parameter(this, Identifier.Description);

            //p15
            p = new Parameter(this, Identifier.AccessibilityObjectName);

            //p16
            p = new Parameter(this, Identifier.UniqueId);

            m_PtrMessageStore->NumberOfMessages++;
        }

        unsafe public void AddMessageConvertType(DataStores SourceStore, DataStores DestinationStore, string typeFullName)
        {
            if (m_DoneFind == false)
            {
                throw new Exception("Need to find the control before querying it");
            }

            if (m_DoneGet == true)
            {
                throw new Exception("Can not query control after getting values from it");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->SourceStore = SourceStore;
            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.ConvertType;
            fixed (void* PtrName = typeFullName)
            {
                NM.CopyMemory(m_IntPtrMemoryMappedFileViewStringStore + m_StringStoreOffset, (IntPtr)PtrName, (UIntPtr)(typeFullName.Length * 2));    //UTF16 charcter: For a 4 byte surrogate pair, length actually returns 2 somewhat confusingly although its convenient for us here, so we can just use length * 2
            }

            PtrMessage->NameOffset = m_StringStoreOffset;
            PtrMessage->NameLength = typeFullName.Length;
            m_StringStoreOffset = m_StringStoreOffset + (typeFullName.Length * 2);
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        unsafe public void AddFirstMessageFindByHandle(DataStores DestinationStore, IntPtr ParentHandle, IntPtr ControlHandle)
        {
            FirstMessageInitialise();
            AddQueryMessageFindByHandle(DestinationStore, ParentHandle, ControlHandle);
            m_DoneFind = true;
            m_DoneQuery = false;
        }

        unsafe public void AddQueryMessageFindByHandle(DataStores DestinationStore, IntPtr ParentHandle, IntPtr ControlHandle)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.RefindByHandle;

            Parameter ParentHandleParam = new Parameter(this, ParentHandle);
            Parameter ControlHandleParam = new Parameter(this, ControlHandle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        unsafe public void AddFirstMessageFindByUniqueId(DataStores destinationStore, string uniqueId)
        {
            FirstMessageInitialise();

            Message* ptrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            ptrMessage->DestinationStore = destinationStore;
            ptrMessage->Action = MessageAction.RefindByUniqueId;

            Parameter uniqueIdParam = new Parameter(this, uniqueId);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
        }

        unsafe public void AddFirstMessageGetAppDomains()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetAppDomains;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddFirstMessageGetListViewItemRectangle(IntPtr ControlHandle, int ItemIndex)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetListViewItemRectangle;

            Parameter ControlHandleParam = new Parameter(this, ControlHandle);
            Parameter GroupIDParam = new Parameter(this, ItemIndex);
            m_MessageNumber = 0;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddFirstMessageGetListViewGroupRectangle(IntPtr ControlHandle, int GroupID)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetListViewGroupRectangle;

            Parameter ControlHandleParam = new Parameter(this, ControlHandle);
            Parameter GroupIDParam = new Parameter(this, GroupID);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddFirstMessageGetContextMenuStrip(IntPtr Handle)
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetContextMenuStrip;

            Parameter HandleParam = new Parameter(this, Handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe private void FirstMessageInitialise()
        {
            m_MessageNumber = 0;
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;
        }

        unsafe public void AddFirstMessageGarbageCollect(int generation)
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GarbageCollect;

            Parameter generationParam = new Parameter(this, generation);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddFirstMessageVisualStyleSupported()
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.VisualStyleSupported;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessagePollMember(DataStores SourceStore, string Name, MemberTypes MemberType, Parameter ItemToPollFor, Parameter unsafeCrossThread)
        {
            if (m_DoneFind == false)
            {
                throw new Exception("Need to find the control before querying it");
            }

            if (m_DoneGet == true)
            {
                throw new Exception("Can not query control after getting values from it");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->SourceStore = SourceStore;
            PtrMessage->Action = MessageAction.ReflectPoll;
            fixed (void* PtrName = Name)
            {
                NM.CopyMemory(m_IntPtrMemoryMappedFileViewStringStore + m_StringStoreOffset, (IntPtr)PtrName, (UIntPtr)(Name.Length * 2));    //UTF16 charcter: For a 4 byte surrogate pair, length actually returns 2 somewhat confusingly although its convenient for us here, so we can just use length * 2
            }

            PtrMessage->NameOffset = m_StringStoreOffset;
            PtrMessage->NameLength = Name.Length;
            m_StringStoreOffset = m_StringStoreOffset + (Name.Length * 2);
            PtrMessage->MemberType = MemberType;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        unsafe public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType)
        {
            if (m_DoneFind == false)
            {
                throw new Exception("Need to find the control before querying it");
            }

            if (m_DoneGet == true)
            {
                throw new Exception("Can not query control after getting values from it");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->SourceStore = SourceStore;
            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.ReflectGet;
            fixed (void* PtrName = Name)
            {
                NM.CopyMemory(m_IntPtrMemoryMappedFileViewStringStore + m_StringStoreOffset, (IntPtr)PtrName, (UIntPtr)(Name.Length * 2));    //UTF16 charcter: For a 4 byte surrogate pair, length actually returns 2 somewhat confusingly although its convenient for us here, so we can just use length * 2
            }

            PtrMessage->NameOffset = m_StringStoreOffset;
            PtrMessage->NameLength = Name.Length;
            m_StringStoreOffset = m_StringStoreOffset + (Name.Length * 2);
            PtrMessage->MemberType = MemberType;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;

            //Debug.Listeners[0].WriteLine("\t AddMessageQueryMember Message: " + (m_PtrMessageStore->NumberOfMessages - 1).ToString() + " Parameters: " + PtrMessage->NumberOfParameters.ToString());
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6, Parameter p7)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6, Parameter p7, Parameter p8)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddQueryMessageReflect(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6, Parameter p7, Parameter p8, Parameter p9)
        {
            AddQueryMessageReflect(SourceStore, DestinationStore, Name, MemberType);
        }

        unsafe public void DecodeControl(int messageNumber, out ControlIdentifier identifier)
        {
            Message* ptrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
            identifier = new ControlIdentifier();

            // p1 = Parent handle
            identifier.ParentHandle = GetParameterIntPtr(ptrMessage, 0);
            // p2 = Handle
            identifier.Handle = GetParameterIntPtr(ptrMessage, 1);

            if (ptrMessage->NumberOfParameters > 2)
            {
                // p3 = Name
                identifier.Name = GetParameterString(ptrMessage, 2);
                // p4 = Technology Type (Winforms, WPF, etc)
                identifier.TechnologyType = GetParameterString(ptrMessage, 3);
                // p5 = Type Namespace
                identifier.TypeNameSpace = GetParameterString(ptrMessage, 4);
                // p6 = Type Name
                identifier.TypeName = GetParameterString(ptrMessage, 5);
                // p7 = Module Name
                identifier.ModuleName = GetParameterString(ptrMessage, 6);
                // p8 = Assembly Name
                identifier.AssemblyName = GetParameterString(ptrMessage, 7);
                // p9 = Index
                identifier.Index = GetParameterInt32(ptrMessage, 8);
                // p10 = Text
                identifier.Text = GetParameterString(ptrMessage, 9);
                // p11 = ChildOf
                identifier.ChildOf = GetParameterIntPtr(ptrMessage, 10);
                // p12 = SiblingOf
                identifier.SiblingOf = GetParameterIntPtr(ptrMessage, 11);
                // p13 = ParentOf
                identifier.ParentOf = GetParameterIntPtr(ptrMessage, 12);
                // p14 = Description
                identifier.Description = GetParameterString(ptrMessage, 13);
                // p15 = AccessibilityObjectName
                identifier.AccessibilityObjectName = GetParameterString(ptrMessage, 14);
                // p16 = UniqueId
                identifier.UniqueId = GetParameterString(ptrMessage, 15);
            }

            CleanUpMessage(ptrMessage);
        }

        private void GetWPFHandleAndNameAndTitle(WPF.Window theWindow)
        {
            m_Handle = new WPF.Interop.WindowInteropHelper(theWindow).Handle;
            m_Name = theWindow.Name;
            m_Text = theWindow.Title;
        }

        private void GetHandleAndName(WF.Control theControl)
        {
            m_Handle = IntPtr.Zero;
            m_Name = "";

            try
            {
                if (theControl.IsHandleCreated)   //if it doesnt have a handle it can't be visible so ignore it
                {
                    m_Handle = (IntPtr)theControl.GetType().GetProperty("HandleInternal", BindingFlags.Instance | BindingFlags.NonPublic).GetValue(theControl, null);
                    m_Name = theControl.Name;
                }
            }
            catch
            {
            }
        }

        unsafe private void wpfFindSourceByHandle(IntPtr handle, ref string name, ref string theText, ref Type theType, ref bool foundControl)
        {
            foreach (WPF.Interop.HwndSource source in WPF.PresentationSource.CurrentSources)
            {
                if (source.Handle == handle)
                {
                    //WPF.PresentationSource presentationSource = (WPF.PresentationSource)source;
                    name = "";
                    theText = GetWindowTextViaWindowMessage(handle);
                    theType = source.GetType();
                    
                    foundControl = true;
                    break;
                }
            }
        }

        unsafe private void wpfFindSourceByIdentifier(ControlIdentifier identifier, ref IntPtr handle, ref string name, ref string theText, ref Type theType, ref int currentIndex, ref bool foundControl)
        {
            foreach (WPF.Interop.HwndSource source in WPF.PresentationSource.CurrentSources)
            {
                //WPF.PresentationSource presentationSource = (WPF.PresentationSource)source;
                if (source.Handle == handle)
                {
                    name = "";
                    theText = GetWindowTextViaWindowMessage(handle);
                    theType = source.GetType();

                    if (identifier.Name != null)
                    {
                        if (name != identifier.Name)
                        {
                            continue;
                        }
                    }

                    if (identifier.TypeNameSpace != null)
                    {
                        if (theType.Namespace == null)
                        {
                            continue;
                        }
                        else
                        {
                            if (!Regex.IsMatch(theType.Namespace, identifier.TypeNameSpace))
                            {
                                continue;
                            }
                        }
                    }

                    if (identifier.TypeName != null)
                    {
                        if (theType.Name == null)
                        {
                            continue;
                        }
                        else
                        {
                            if (!Regex.IsMatch(theType.Name, identifier.TypeName))
                            {
                                continue;
                            }
                        }
                    }

                    if (identifier.ModuleName != null)
                    {
                        if (theType.Module.Name != identifier.ModuleName)
                        {
                            continue;
                        }
                    }

                    if (identifier.AssemblyName != null)
                    {
                        if (theType.Assembly.GetName().Name != identifier.AssemblyName)
                        {
                            continue;
                        }
                    }

                    if (identifier.Text != null)
                    {
                        if (theText == null)
                        {
                            continue;
                        }
                        else
                        {
                            if (!Regex.IsMatch(theText, identifier.Text))
                            {
                                continue;
                            }
                        }
                    }

                    currentIndex++;

                    DebugLogging.WriteLog("found wpf form for " + name);

                    if (identifier.Index > 0)
                    {
                        if (currentIndex != identifier.Index)
                        {
                            continue;
                        }
                    }

                    //we have a match
                    if (NM.IsWindowVisible(handle))
                    {
                        foundControl = true;
                        break;
                    }
                }
            }
        }

        unsafe private string Find(int messageNumber, uint timeoutMs)
        {
            ControlIdentifier Identifier;
            DecodeControl(messageNumber, out Identifier);

            //Searching for the control by property
            Type theType = null;
            string uniqueId = null;
            string typeName = null;
            string typeNameSpace = null;
            string technologyType = null;
            IntPtr Handle = IntPtr.Zero;
            string Name = null;
            string accessibilityObjectName = null;
            string theText = null;
            bool FoundControl = false;

            if (Identifier.ParentHandle == IntPtr.Zero)     //Find toplevel (parent)
            {
                if (Identifier.Handle != IntPtr.Zero || Identifier.UniqueId != null)
                {
                    if (Identifier.Handle == IntPtr.Zero && Identifier.UniqueId != null)
                    {
                        switch (Identifier.UniqueId.Substring(0, 1))
                        {
                            case "A":
                                if (Identifier.TechnologyType == "Windows ActiveX" || Identifier.TechnologyType == null)
                                {
                                    uniqueId = Identifier.UniqueId;
                                    object controlActiveX = FindByUniqueIdActiveX(uniqueId, out Name, out typeNameSpace, out typeName, out Handle);
                                    if (controlActiveX != null)
                                    {
                                        theText = GetWindowTextViaWindowMessage(Handle);
                                        technologyType = "Windows ActiveX";
                                        FoundControl = true;
                                    }
                                }
                                break;
                            case "H":
                                Identifier.Handle = new IntPtr(long.Parse(Identifier.UniqueId.Substring(1)));
                                break;
                        }
                    }

                    if (!FoundControl)
                    {
                        if (Identifier.TechnologyType == "Windows Forms (WinForms)" || Identifier.TechnologyType == null)
                        {
                            //WinForms
                            WF.Control TheControl = WF.Control.FromHandle(Identifier.Handle);
                            if (TheControl != null)
                            {
                                GetHandleAndName(TheControl);
                                Handle = m_Handle;
                                Name = m_Name;
                                theType = TheControl.GetType();
                                object[] parameters = { TheControl };
                                theText = (string)TheControl.Invoke(m_GetTextDelegater, parameters);
                                accessibilityObjectName = (string)TheControl.Invoke(m_GetAccessibilityObjectNameDelegater, parameters);
                                FoundControl = true;
                            }
                        }
                    }

                    if (!FoundControl)
                    {
                        if (Identifier.TechnologyType == "Windows ActiveX" || Identifier.TechnologyType == null)
                        {
                            Handle = Identifier.Handle;
                            object controlActiveX = FindByHandleActiveX(Identifier.Handle, out Name, out typeNameSpace, out typeName, out uniqueId);
                            if (controlActiveX != null)
                            {
                                theText = GetWindowTextViaWindowMessage(Handle);
                                technologyType = "Windows ActiveX";
                                FoundControl = true;
                            }
                        }
                    }

                    if (!FoundControl)
                    {
                        if (m_WPF)  //Only look for WPF source if the application has WPF loaded
                        {
                            if (Identifier.TechnologyType == "Windows Presentation Foundation (WPF)" || Identifier.TechnologyType == null)
                            {
                                Handle = Identifier.Handle;
                                wpfFindSourceByHandle(Handle, ref Name, ref theText, ref theType, ref FoundControl);
                            }
                        }
                    }

                    if (!FoundControl)
                    {
                        if (Identifier.TechnologyType == "Windows Native" || Identifier.TechnologyType == null)
                        {
                            //Native
                            Handle = Identifier.Handle;
                            theText = GetWindowTextViaWindowMessage(Identifier.Handle);
                            FoundControl = true;
                        }
                    }
                }
                else    //Look form without having its handle
                {
                    //Start polling
                    Stopwatch timer = Stopwatch.StartNew();
                    do
                    {
                        //looking for a form
                        try
                        {
                            int CurrentIndex = 0;

                            m_AllControls = new List<IntPtr>();
                            //0 for the thread seems to enumerate all threads
                            NM.EnumThreadWindows(0, EnumThreadProcedue, IntPtr.Zero);

                            foreach (IntPtr hWnd in m_AllControls)
                            {
                                if (NM.IsWindowVisible(hWnd))
                                {
                                    if (Identifier.TechnologyType == "Windows Forms (WinForms)" || Identifier.TechnologyType == null)
                                    {
                                        //WinForms
                                        WF.Control form = WF.Control.FromHandle(hWnd);
                                        if (form != null)
                                        {
                                            theType = form.GetType();
                                            GetHandleAndName(form);
                                            Handle = m_Handle;
                                            Name = m_Name;

                                            if (Identifier.Name != null)
                                            {
                                                if (Name != Identifier.Name)
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.TechnologyType != null)
                                            {
                                                if ("Windows Forms (WinForms)" != Identifier.TechnologyType)
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.TypeNameSpace != null)
                                            {
                                                if (theType.Namespace == null)
                                                {
                                                    continue;
                                                }
                                                else
                                                {
                                                    if (!Regex.IsMatch(theType.Namespace, Identifier.TypeNameSpace))
                                                    {
                                                        continue;
                                                    }
                                                }
                                            }

                                            if (Identifier.TypeName != null)
                                            {
                                                if (theType.Name == null)
                                                {
                                                    continue;
                                                }
                                                else
                                                {
                                                    if (!Regex.IsMatch(theType.Name, Identifier.TypeName))
                                                    {
                                                        continue;
                                                    }
                                                }
                                            }

                                            if (Identifier.ModuleName != null)
                                            {
                                                if (theType.Module.Name != Identifier.ModuleName)
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.AssemblyName != null)
                                            {
                                                if (theType.Assembly.GetName().Name != Identifier.AssemblyName)
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.ChildOf != IntPtr.Zero)
                                            {
                                                if (!NM.IsChild(Identifier.ChildOf, Handle))
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.SiblingOf != IntPtr.Zero)
                                            {
                                                if (!NM.IsSibling(Identifier.SiblingOf, Handle))
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.ParentOf != IntPtr.Zero)
                                            {
                                                if (!NM.IsChild(Handle, Identifier.ParentOf))
                                                {
                                                    continue;
                                                }
                                            }

                                            object[] parameters = { form };

                                            theText = null;
                                            if (!form.IsDisposed && !form.Disposing)
                                            {
                                                try
                                                {
                                                    theText = (string)form.Invoke(m_GetTextDelegater, parameters);
                                                }
                                                catch (ObjectDisposedException ex)
                                                {
                                                }
                                            }
                                            if (Identifier.Text != null)
                                            {
                                                if (theText == null)
                                                {
                                                    continue;
                                                }
                                                else
                                                {
                                                    if (!Regex.IsMatch(theText, Identifier.Text))
                                                    {
                                                        continue;
                                                    }
                                                }
                                            }

                                            accessibilityObjectName = null;
                                            if (!form.IsDisposed && !form.Disposing)
                                            {
                                                try
                                                {
                                                    accessibilityObjectName = (string)form.Invoke(m_GetAccessibilityObjectNameDelegater, parameters);
                                                }
                                                catch (ObjectDisposedException ex)
                                                {
                                                }
                                            }
                                            if (Identifier.AccessibilityObjectName != null)
                                            {
                                                if (accessibilityObjectName != Identifier.AccessibilityObjectName)
                                                {
                                                    continue;
                                                }
                                            }

                                            DebugLogging.WriteLog("found form for " + Name);

                                            CurrentIndex++;
                                            if (Identifier.Index > 0)
                                            {
                                                if (CurrentIndex != Identifier.Index)
                                                {
                                                    continue;
                                                }
                                            }

                                            //we have a match
                                            FoundControl = true;
                                            break;
                                        }
                                    }

                                    if (Identifier.TechnologyType == "Windows ActiveX" || Identifier.TechnologyType == null)
                                    {
                                        Handle = hWnd;
                                        FindByIdentifierActiveX(Identifier, ref Handle, ref Name, ref theText, ref typeNameSpace, ref typeName, ref CurrentIndex, ref uniqueId, ref FoundControl);
                                        if (FoundControl)
                                        {
                                            technologyType = "Windows ActiveX";
                                            break;
                                        }
                                    }

                                    if (m_WPF)  //Only look for WPF source if the application has WPF loaded
                                    {
                                        if (Identifier.TechnologyType == "Windows Presentation Foundation (WPF)" || Identifier.TechnologyType == null)
                                        {
                                            Handle = hWnd;
                                            wpfFindSourceByIdentifier(Identifier, ref Handle, ref Name, ref theText, ref theType, ref CurrentIndex, ref FoundControl);
                                            if (FoundControl)
                                            {
                                                break;
                                            }
                                        }
                                    }

                                    if (Identifier.TechnologyType == "Windows Native" || Identifier.TechnologyType == null)
                                    {
                                        //Windows Native
                                        Handle = hWnd;

                                        if (Identifier.Name != null)
                                        {
                                            continue;
                                        }

                                        if (Identifier.TechnologyType != null)
                                        {
                                            if ("Windows Native" != Identifier.TechnologyType)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.TypeNameSpace != null)
                                        {
                                            continue;
                                        }

                                        if (Identifier.TypeName != null)
                                        {
                                            string theTypeName = NM.GetClassName(Handle);
                                            if (theTypeName == null)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (!Regex.IsMatch(theTypeName, Identifier.TypeName))
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        if (Identifier.ModuleName != null)
                                        {
                                            string theModuleName = Path.GetFileName(NM.GetWindowModuleFileName(Handle));
                                            if (theModuleName != Identifier.ModuleName)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.AssemblyName != null)
                                        {
                                            continue;
                                        }

                                        if (Identifier.ChildOf != IntPtr.Zero)
                                        {
                                            if (!NM.IsChild(Identifier.ChildOf, Handle))
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.SiblingOf != IntPtr.Zero)
                                        {
                                            if (!NM.IsSibling(Identifier.SiblingOf, Handle))
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.ParentOf != IntPtr.Zero)
                                        {
                                            if (!NM.IsChild(Handle, Identifier.ParentOf))
                                            {
                                                continue;
                                            }
                                        }

                                        theText = GetWindowTextViaWindowMessage(Handle);
                                        if (Identifier.Text != null)
                                        {
                                            if (theText == null)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (!Regex.IsMatch(theText, Identifier.Text))
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        if (Identifier.AccessibilityObjectName != null)
                                        {
                                            IAccessible obj = NM.IAccessibleObjectFromWindow(Handle);
                                            if (obj == null || obj.accName == null)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (obj.accName != Identifier.AccessibilityObjectName)
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        DebugLogging.WriteLog("found form for " + Name);

                                        CurrentIndex++;
                                        if (Identifier.Index > 0)
                                        {
                                            if (CurrentIndex != Identifier.Index)
                                            {
                                                continue;
                                            }
                                        }

                                        //Confirm its the correct process
                                        int pid;
                                        NM.GetWindowThreadProcessId(Handle, out pid);
                                        if (AUTProcess.Id != pid)
                                        {
                                            continue;
                                        }

                                        //we have a match
                                        FoundControl = true;
                                        break;
                                    }
                                }
                            }

                            if (!FoundControl)
                            {
                                FindByIdentifierRenderedActiveX(Identifier, ref Handle, ref Name, ref theText, ref typeNameSpace, ref typeName, ref technologyType, ref uniqueId, ref FoundControl);
                            }
                        }
                        catch (InvalidOperationException ex)
                        {
                            if (ex.Message == "Collection was modified; enumeration operation may not execute.")
                            {
                                //Ignore
                            }
                            else if (ex.Message.Contains("Cannot access a disposed object."))
                            {
                                //Ignore
                            }
                            else
                            {
                                throw;
                            }
                        }

                        DebugLogging.WriteLog("done find form loop for " + Name);

                        if (FoundControl == false)
                        {
                            DebugLogging.WriteLog("");
                            Thread.Sleep(15);
                        }
                    }
                    while (FoundControl == false && timer.ElapsedMilliseconds < timeoutMs);
                    timer.Stop();
                }
            }
            else    //find child
            {
                if (Identifier.Handle != IntPtr.Zero || Identifier.UniqueId != null)
                {
                    if (Identifier.Handle == IntPtr.Zero && Identifier.UniqueId != null)
                    {
                        switch (Identifier.UniqueId.Substring(0, 1))
                        {
                            case "A":
                                if (Identifier.TechnologyType == "Windows ActiveX" || Identifier.TechnologyType == null)
                                {
                                    uniqueId = Identifier.UniqueId;
                                    object controlActiveX = FindByUniqueIdActiveX(uniqueId, out Name, out typeNameSpace, out typeName, out Handle);
                                    if (controlActiveX != null)
                                    {
                                        theText = GetWindowTextViaWindowMessage(Handle);
                                        technologyType = "Windows ActiveX";
                                        FoundControl = true;
                                    }
                                }
                                break;
                            case "H":
                                Identifier.Handle = new IntPtr(long.Parse(Identifier.UniqueId.Substring(1)));
                                break;
                        }
                    }

                    //WinForms
                    if (!FoundControl)
                    {
                        if (Identifier.TechnologyType == "Windows Forms (WinForms)" || Identifier.TechnologyType == null)
                        {
                            WF.Control TheControl = WF.Control.FromHandle(Identifier.Handle);

                            if (TheControl != null)
                            {
                                GetHandleAndName(TheControl);
                                Handle = m_Handle;
                                Name = m_Name;
                                object[] parameters = { TheControl };
                                theText = (string)TheControl.Invoke(m_GetTextDelegater, parameters);
                                accessibilityObjectName = (string)TheControl.Invoke(m_GetAccessibilityObjectNameDelegater, parameters);
                                theType = TheControl.GetType();
                                FoundControl = true;
                            }
                        }
                    }

                    if (!FoundControl)
                    {
                        if (Identifier.TechnologyType == "Windows ActiveX" || Identifier.TechnologyType == null)
                        {
                            //TODO check parent handle
                            Handle = Identifier.Handle;
                            object controlActiveX = FindByHandleActiveX(Identifier.Handle, out Name, out typeNameSpace, out typeName, out uniqueId);
                            if (controlActiveX != null)
                            {
                                theText = GetWindowTextViaWindowMessage(Handle);
                                technologyType = "Windows ActiveX";
                                FoundControl = true;
                            }
                        }
                    }

                    if (!FoundControl)
                    {
                        if (m_WPF)  //Only look for WPF source if the application has WPF loaded
                        {
                            if (Identifier.TechnologyType == "Windows Presentation Foundation (WPF)" || Identifier.TechnologyType == null)
                            {
                                Handle = Identifier.Handle;
                                wpfFindSourceByHandle(Handle, ref Name, ref theText, ref theType, ref FoundControl);
                            }
                        }
                    }

                    if (!FoundControl)
                    {
                        if (Identifier.TechnologyType == "Windows Native" || Identifier.TechnologyType == null)
                        {
                            //Native
                            Handle = Identifier.Handle;
                            theText = GetWindowTextViaWindowMessage(Identifier.Handle);
                            FoundControl = true;
                        }
                    }
                }
                else    //Look for a child without having its handle
                {
                    //Start polling
                    Stopwatch timer = Stopwatch.StartNew();
                    do
                    {
                        int CurrentIndex = 0;

                        m_AllControls = new List<IntPtr>();
                        //Build a list of all the visible child windows of the parent form
                        BuildListOfAllControls(Identifier.ParentHandle);
                        //NM.EnumChildWindows(Identifier.ParentHandle, EnumThreadProcedue, IntPtr.Zero);

                        foreach (IntPtr hWnd in m_AllControls)
                        {
                            if (NM.IsWindowVisible(hWnd))
                            {
                                if (Identifier.TechnologyType == "Windows Forms (WinForms)" || Identifier.TechnologyType == null)
                                {
                                    //WinForms
                                    WF.Control control = WF.Control.FromHandle(hWnd);
                                    if (control != null)
                                    {
                                        GetHandleAndName(control);
                                        Handle = m_Handle;
                                        Name = m_Name;
                                        theType = control.GetType();

                                        if (Identifier.Name != null)
                                        {
                                            if (Name != Identifier.Name)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.TechnologyType != null)
                                        {
                                            if (GetTechnologyType(theType) != Identifier.TechnologyType)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.TypeNameSpace != null)
                                        {
                                            if (theType.Namespace == null)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (!Regex.IsMatch(theType.Namespace, Identifier.TypeNameSpace))
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        if (Identifier.TypeName != null)
                                        {
                                            if (theType.Name == null)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (!Regex.IsMatch(theType.Name, Identifier.TypeName))
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        if (Identifier.ModuleName != null)
                                        {
                                            if (theType.Module.Name != Identifier.ModuleName)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.AssemblyName != null)
                                        {
                                            if (theType.Assembly.GetName().Name != Identifier.AssemblyName)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.ChildOf != IntPtr.Zero)
                                        {
                                            if (!NM.IsChild(Identifier.ChildOf, Handle))
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.SiblingOf != IntPtr.Zero)
                                        {
                                            if (!NM.IsSibling(Identifier.SiblingOf, Handle))
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.ParentOf != IntPtr.Zero)
                                        {
                                            if (!NM.IsChild(Handle, Identifier.ParentOf))
                                            {
                                                continue;
                                            }
                                        }

                                        object[] parameters = { control };

                                        theText = null;
                                        if (!control.IsDisposed && !control.Disposing)
                                        {
                                            try
                                            {
                                                theText = (string)control.Invoke(m_GetTextDelegater, parameters);
                                            }
                                            catch (ObjectDisposedException ex)
                                            {
                                            }
                                        }
                                        if (Identifier.Text != null)
                                        {
                                            if (theText == null)
                                            {
                                                continue;
                                            }
                                            else
                                            {
                                                if (!Regex.IsMatch(theText, Identifier.Text))
                                                {
                                                    continue;
                                                }
                                            }
                                        }

                                        accessibilityObjectName = null;
                                        if (!control.IsDisposed && !control.Disposing)
                                        {
                                            try
                                            {
                                                accessibilityObjectName = (string)control.Invoke(m_GetAccessibilityObjectNameDelegater, parameters);
                                            }
                                            catch (ObjectDisposedException ex)
                                            {
                                            }
                                        }
                                        if (Identifier.AccessibilityObjectName != null)
                                        {
                                            if (accessibilityObjectName != Identifier.AccessibilityObjectName)
                                            {
                                                continue;
                                            }
                                        }

                                        CurrentIndex++;
                                        if (Identifier.Index > 0)
                                        {
                                            if (CurrentIndex != Identifier.Index)
                                            {
                                                continue;
                                            }
                                        }

                                        //we have a match
                                        FoundControl = true;
                                        break;
                                    }
                                }

                                if (Identifier.TechnologyType == "Windows ActiveX" || Identifier.TechnologyType == null)
                                {
                                    Handle = hWnd;
                                    FindByIdentifierActiveX(Identifier, ref Handle, ref Name, ref theText, ref typeNameSpace, ref typeName, ref CurrentIndex, ref uniqueId, ref FoundControl);
                                    if (FoundControl)
                                    {
                                        technologyType = "Windows ActiveX";
                                        break;
                                    }
                                }

                                if (m_WPF)  //Only look for WPF source if the application has WPF loaded
                                {
                                    if (Identifier.TechnologyType == "Windows Presentation Foundation (WPF)" || Identifier.TechnologyType == null)
                                    {
                                        Handle = hWnd;
                                        wpfFindSourceByIdentifier(Identifier, ref Handle, ref Name, ref theText, ref theType, ref CurrentIndex, ref FoundControl);
                                        if (FoundControl)
                                        {
                                            break;
                                        }
                                    }
                                }

                                if (Identifier.TechnologyType == "Windows Native" || Identifier.TechnologyType == null)
                                {
                                    //Windows Native
                                    Handle = hWnd;

                                    if (Identifier.Name != null)
                                    {
                                        continue;
                                    }

                                    if (Identifier.TechnologyType != null)
                                    {
                                        if ("Windows Native" != Identifier.TechnologyType)
                                        {
                                            continue;
                                        }
                                    }

                                    if (Identifier.TypeNameSpace != null)
                                    {
                                        continue;
                                    }

                                    if (Identifier.TypeName != null)
                                    {
                                        string theTypeName = NM.GetClassName(Handle);
                                        if (theTypeName == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            if (!Regex.IsMatch(theTypeName, Identifier.TypeName))
                                            {
                                                continue;
                                            }
                                        }
                                    }

                                    if (Identifier.ModuleName != null)
                                    {
                                        string theModuleName = Path.GetFileName(NM.GetWindowModuleFileName(Handle));
                                        if (theModuleName != Identifier.ModuleName)
                                        {
                                            continue;
                                        }
                                    }

                                    if (Identifier.AssemblyName != null)
                                    {
                                        continue;
                                    }

                                    if (Identifier.ChildOf != IntPtr.Zero)
                                    {
                                        if (!NM.IsChild(Identifier.ChildOf, Handle))
                                        {
                                            continue;
                                        }
                                    }

                                    if (Identifier.SiblingOf != IntPtr.Zero)
                                    {
                                        if (!NM.IsSibling(Identifier.SiblingOf, Handle))
                                        {
                                            continue;
                                        }
                                    }

                                    if (Identifier.ParentOf != IntPtr.Zero)
                                    {
                                        if (!NM.IsChild(Handle, Identifier.ParentOf))
                                        {
                                            continue;
                                        }
                                    }

                                    theText = GetWindowTextViaWindowMessage(Handle);
                                    if (Identifier.Text != null)
                                    {
                                        if (theText == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            if (!Regex.IsMatch(theText, Identifier.Text))
                                            {
                                                continue;
                                            }
                                        }
                                    }

                                    if (Identifier.AccessibilityObjectName != null)
                                    {
                                        IAccessible obj = NM.IAccessibleObjectFromWindow(Handle);
                                        if (obj == null || obj.accName == null)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            if (obj.accName != Identifier.AccessibilityObjectName)
                                            {
                                                continue;
                                            }
                                        }
                                    }

                                    CurrentIndex++;
                                    if (Identifier.Index > 0)
                                    {
                                        if (CurrentIndex != Identifier.Index)
                                        {
                                            continue;
                                        }
                                    }

                                    //we have a match
                                    FoundControl = true;
                                    break;
                                }
                            }
                        }

                        if (!FoundControl)
                        {
                            FindByIdentifierRenderedActiveX(Identifier, ref Handle, ref Name, ref theText, ref typeNameSpace, ref typeName, ref technologyType, ref uniqueId, ref FoundControl);
                        }

                        if (!FoundControl)
                        {
                            DebugLogging.WriteLog("");
                            Thread.Sleep(15);
                        }
                    }
                    while (FoundControl == false && timer.ElapsedMilliseconds < timeoutMs);
                    timer.Stop();
                }
            }

            //return the item
            if (FoundControl)
            {
                ControlIdentifier NewIdentifier = new ControlIdentifier();

                NewIdentifier.Description = Identifier.Description;
                NewIdentifier.ParentHandle = Identifier.ParentHandle;
                NewIdentifier.Handle = Handle;
                NewIdentifier.Name = Name;
                if (string.IsNullOrEmpty(technologyType))
                {
                    NewIdentifier.TechnologyType = GetTechnologyType(theType);
                }
                else
                {
                    NewIdentifier.TechnologyType = technologyType;
                }
                if (NewIdentifier.TechnologyType == "Windows Native")
                {
                    NewIdentifier.TypeName = NM.GetClassName(Identifier.Handle);
                    NewIdentifier.ModuleName = Path.GetFileName(NM.GetWindowModuleFileName(Identifier.Handle));

                    IAccessible obj = NM.IAccessibleObjectFromWindow(Handle);
                    if (obj != null && obj.accName != null)
                    {
                        NewIdentifier.AccessibilityObjectName = obj.accName;
                    }
                }
                else if(NewIdentifier.TechnologyType == "Windows ActiveX")
                {
                    NewIdentifier.TypeName = typeName;
                    NewIdentifier.TypeNameSpace = typeNameSpace;
                    NewIdentifier.ModuleName = Path.GetFileName(NM.GetWindowModuleFileName(Handle));
                }
                else
                {
                    NewIdentifier.TypeNameSpace = theType.Namespace;
                    NewIdentifier.TypeName = theType.Name;
                    NewIdentifier.ModuleName = theType.Module.Name;
                    NewIdentifier.AssemblyName = theType.Assembly.GetName().Name;
                    NewIdentifier.AccessibilityObjectName = accessibilityObjectName;
                }
                NewIdentifier.Index = Identifier.Index;
                NewIdentifier.Text = theText;
                if (uniqueId == null)
                {
                    uniqueId = "H" + Handle.ToString();
                }
                NewIdentifier.UniqueId = uniqueId;
                AddIdentifierMessage(NewIdentifier);
                return null;
            }
            else
            {
                return "Failed to find the " + Identifier.Description;
            }
        }

        unsafe private void RefindByUniqueId(int messageNumber)
        {
            object destinationObject = null;

            Message* ptrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));

            string uniqueId = GetParameterString(ptrMessage, 0);
            IntPtr handle;

            switch (uniqueId.Substring(0, 1))
            {
                case "H":
                    handle = new IntPtr(long.Parse(uniqueId.Substring(1)));
                    destinationObject = WF.Control.FromHandle(handle);
                    break;
                case "A":
                    destinationObject = FindByUniqueIdActiveX(uniqueId, out string name, out string typeNameSpace, out string typeName, out handle);
                    break;
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        unsafe private void RefindByHandle(int messageNumber)
        {
            object destinationObject = null;

            Message* ptrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));

            ControlIdentifier identifier;
            DecodeControl(messageNumber, out identifier);

            //WinForms
            if (identifier.TechnologyType == "Windows Forms (WinForms)" || identifier.TechnologyType == null)
            {
                destinationObject = WF.Control.FromHandle(identifier.Handle);
            }

            if (destinationObject == null)
            {
                if (identifier.TechnologyType == "Windows ActiveX" || identifier.TechnologyType == null)
                {
                    destinationObject = FindByHandleActiveX(identifier.Handle, out string name, out string typeNameSpace, out string typeName, out string uniqueId);
                }
            }

            if (destinationObject == null)
            {
                if (m_WPF)  //Only look for WPF source if the application has WPF loaded
                {
                    if (identifier.TechnologyType == "Windows Presentation Foundation (WPF)" || identifier.TechnologyType == null)
                    {
                        //WPF TODO
                    }
                }
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        unsafe private void GetResult(int MessageNumber)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));
            object SourceObject;

            switch (PtrMessage->SourceStore)
            {
                case DataStores.Store0:
                    SourceObject = tempStore0;
                    break;
                case DataStores.Store1:
                    SourceObject = tempStore1;
                    break;
                case DataStores.Store2:
                    SourceObject = tempStore2;
                    break;
                case DataStores.Store3:
                    SourceObject = tempStore3;
                    break;
                case DataStores.Store4:
                    SourceObject = tempStore4;
                    break;
                case DataStores.Store5:
                    SourceObject = tempStore5;
                    break;
                case DataStores.Store6:
                    SourceObject = tempStore6;
                    break;
                case DataStores.Store7:
                    SourceObject = tempStore7;
                    break;
                case DataStores.Store8:
                    SourceObject = tempStore8;
                    break;
                case DataStores.Store9:
                    SourceObject = tempStore9;
                    break;
                default:
                    throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
            }

            //store the important parts of the message so we can cleanup
            int Parameters = PtrMessage->NumberOfParameters;
            ApeTypeCode ParameterTypeCode = ApeTypeCode.Empty;
            if (Parameters == 1)
            {
                ParameterTypeCode = (ApeTypeCode)PtrMessage->Parameter.TypeCode[0];
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            if (SourceObject == null)
            {
                PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)ApeTypeCode.Empty;
                PtrMessage->TypeCodeKey += ((PtrMessage->NumberOfParameters * Parameter.OneLargerThanApeTypeCodeEnumMax * 2) + (int)ApeTypeCode.Empty) + ((PtrMessage->NumberOfParameters * Parameter.OneLargerThanApeTypeCodeEnumMax * 2) + Parameter.OneLargerThanApeTypeCodeEnumMax + (int)ParameterType.In);
                PtrMessage->NumberOfParameters++;
                m_PtrMessageStore->NumberOfMessages++;
            }
            else
            {
                //TODO is it worth doing to avoid dynamic for performance?
                //if (Parameters == 1)
                //{
                //    //fast (explicit casting) 
                //    //TO DO add rest of switch
                //    switch (ParameterTypeCode)
                //    {
                //        case ApeTypeCode.Int32:
                //            AddReturnValue(new Parameter(this, (int)SourceObject));
                //            break;
                //        case ApeTypeCode.String:
                //            AddReturnValue(new Parameter(this, (string)SourceObject));
                //            break;
                //        default:
                //            throw new Exception("Unsupported typecode: " + PtrMessage->Parameter.TypeCode[0].ToString());
                //    }
                //}
                //else
                //{
                //slow due to the use of dynamic
                AddReturnValue(new Parameter(this, (dynamic)SourceObject));
                //}
            }
        }

        unsafe private void GetApeTypeFromType(int MessageNumber)
        {
            GetApeTypeInternal(MessageNumber, true);
        }

        unsafe private void GetApeTypeFromObject(int MessageNumber)
        {
            GetApeTypeInternal(MessageNumber, false);
        }

        unsafe private void GetApeTypeInternal(int MessageNumber, bool isType)
        {
            object SourceObject;
            object DestinationObject;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            switch (PtrMessage->SourceStore)
            {
                case DataStores.Store0:
                    SourceObject = tempStore0;
                    break;
                case DataStores.Store1:
                    SourceObject = tempStore1;
                    break;
                case DataStores.Store2:
                    SourceObject = tempStore2;
                    break;
                case DataStores.Store3:
                    SourceObject = tempStore3;
                    break;
                case DataStores.Store4:
                    SourceObject = tempStore4;
                    break;
                case DataStores.Store5:
                    SourceObject = tempStore5;
                    break;
                case DataStores.Store6:
                    SourceObject = tempStore6;
                    break;
                case DataStores.Store7:
                    SourceObject = tempStore7;
                    break;
                case DataStores.Store8:
                    SourceObject = tempStore8;
                    break;
                case DataStores.Store9:
                    SourceObject = tempStore9;
                    break;
                default:
                    throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
            }

            if (SourceObject == null)
            {
                DestinationObject = null;
            }
            else
            {
                Type ControlType;

                if (isType)
                {
                    ControlType = (Type)SourceObject;
                }
                else
                {
                    ControlType = SourceObject.GetType();
                }

                DestinationObject = GetAPETypeFromTypeInternal(ControlType.Namespace, ControlType.Name);
            }

            switch (PtrMessage->DestinationStore)
            {
                case DataStores.Store0:
                    tempStore0 = DestinationObject;
                    break;
                case DataStores.Store1:
                    tempStore1 = DestinationObject;
                    break;
                case DataStores.Store2:
                    tempStore2 = DestinationObject;
                    break;
                case DataStores.Store3:
                    tempStore3 = DestinationObject;
                    break;
                case DataStores.Store4:
                    tempStore4 = DestinationObject;
                    break;
                case DataStores.Store5:
                    tempStore5 = DestinationObject;
                    break;
                case DataStores.Store6:
                    tempStore6 = DestinationObject;
                    break;
                case DataStores.Store7:
                    tempStore7 = DestinationObject;
                    break;
                case DataStores.Store8:
                    tempStore8 = DestinationObject;
                    break;
                case DataStores.Store9:
                    tempStore9 = DestinationObject;
                    break;
                default:
                    throw new Exception("Unsupported DestinationStore " + (PtrMessage->DestinationStore).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        private string GetAPETypeFromTypeInternal(string TypeNameSpace, string TypeName)
        {
            switch (TypeNameSpace)
            {
                case "System.Windows.Forms":
                    switch (TypeName)
                    {
                        case "ListView":
                            return "GUIListView";
                        case "TreeView":
                            return "GUITreeView";
                        case "ListBox":
                            return "GUIListBox";
                        case "ComboBox":
                            return "GUIComboBox";
                        case "RadioButton":
                            return "GUIRadioButton";
                        case "CheckBox":
                            return "GUICheckBox";
                        case "TextBox":
                            return "GUITextBox";
                        case "StatusStrip":
                            return "GUIStatusStrip";
                        case "StatusBar":
                            return "GUIStatusBar";
                        case "Label":
                            return "GUILabel";
                        case "Button":
                            return "GUIButton";
                        case "MenuStrip":
                            return "GUIMenuStrip";
                        case "NumericUpDown":
                            return "GUIUpDown";
                        case "DateTimePicker":
                            return "GUIDateTimePicker";
                        case "ProgressBar":
                            return "GUIProgressBar";
                        case "PictureBox":
                            return "GUIPictureBox";
                        case "ToolStrip":
                            return "GUIToolStrip";
                        case "TabControl":
                            return "GUITabControl";
                        case "CheckedListBox":
                            return "GUICheckedListBox";
                        case "HScrollBar":
                            return "GUIHScrollBar";
                        case "VScrollBar":
                            return "GUIVScrollBar";
                        case "SplitContainer":
                            return "GUISplitContainer";
                        case "DataGridView":
                            return "GUIDataGridView";
                        default:
                            return "";
                    }
                case "TD.SandDock":
                    switch (TypeName)
                    {
                        case "DocumentContainer":
                            return "GUIDocumentContainer";
                        case "DockableWindow":
                            return "GUIDockableWindow";
                        case "DockContainer":
                            return "GUIDockContainer";
                        case "AutoHideBar":
                            return "GUIAutoHideBar";
                        case "PopupContainer":
                            return "GUIPopupContainer";
                        default:
                            return "";
                    }
                case "C1.Win.C1FlexGrid":
                    switch (TypeName)
                    {
                        case "C1FlexGrid":
                            return "GUIFlexgrid";
                        default:
                            return "";
                    }
                case "LatentZero.Capstone.ComSupport.ResultsGrid":
                case "AxDRILLDOWNLib":
                    switch (TypeName)
                    {
                        case "AxLZResultsGrid":
                            return "GUIAxLZResultsGrid";
                        default:
                            return "";
                    }
                case "LatentZero.Interop.LzCashViewerCtl":
                case "LZCASHVIEWERLib":
                    switch (TypeName)
                    {
                        case "LzCashViewerCtrl":
                            return "GUIAxLzCashViewerCtrl";
                        default:
                            return "";
                    }
                case "LatentZero.Capstone.Controls.Walker":
                    switch (TypeName)
                    {
                        case "WalkerControl":
                            return "GUIGenericWalker";
                        default:
                            return "";
                    }
                case "LatentZero.Utility.Controls":
                    switch (TypeName)
                    {
                        case "TitleFrame":
                            return "GUITitleFrame";
                        case "LzComboBox":
                            return "GUIComboBox";
                        default:
                            return "";
                    }
                case "LatentZero.Common.ModelControls.ElementStrip":
                    switch (TypeName)
                    {
                        case "ElementStripGrid":
                            return "GUIElementStripGrid";
                        default:
                            return "";
                    }
                case "LatentZero.Common.ModelControls.ElementStrip.Grouping":
                    switch (TypeName)
                    {
                        case "GroupingBar":
                            return "GUIGroupingBar";
                        default:
                            return "";
                    }
                case "LzGenericWalker":
                    switch (TypeName)
                    {
                        case "GenericWalker":
                            return "GUIGenericWalker";
                        default:
                            return "";
                    }
                case "LzcDealingResults":
                    switch (TypeName)
                    {
                        case "lzDealingResults":
                            return "GUILzDealingResults";
                        default:
                            return "";
                    }
                case "XPExplorerBar":
                    switch (TypeName)
                    {
                        case "Expando":
                            return "GUIExpando";
                        default:
                            return "";
                    }
                case "LzStretchyCombo":
                    switch (TypeName)
                    {
                        case "StretchyCombo":
                            return "GUIStretchyCombo";
                        default:
                            return "";
                    }
                case "LzNavBarControls":
                    switch (TypeName)
                    {
                        case "LzNavBarGridControl":
                            return "GUILzNavBarGridControl";
                        default:
                            return "";
                    }
                case "AxLatentZero.Interop.LzFlexGrid":
                    switch (TypeName)
                    {
                        case "AxLzFlexGridCtrl":
                            return "GUIFlexgrid";
                        default:
                            return "";
                    }
                case "":
                    if (TypeName.Contains("tooltips_class"))
                    {
                        return "GUIToolTip";
                    }
                    else
                    {
                        return "";
                    }
                default:
                    return "";
            }
        }

        unsafe private void GetRecognisedType(int MessageNumber)
        {
            object SourceObject;
            object DestinationObject;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            switch (PtrMessage->SourceStore)
            {
                case DataStores.Store0:
                    SourceObject = tempStore0;
                    break;
                case DataStores.Store1:
                    SourceObject = tempStore1;
                    break;
                case DataStores.Store2:
                    SourceObject = tempStore2;
                    break;
                case DataStores.Store3:
                    SourceObject = tempStore3;
                    break;
                case DataStores.Store4:
                    SourceObject = tempStore4;
                    break;
                case DataStores.Store5:
                    SourceObject = tempStore5;
                    break;
                case DataStores.Store6:
                    SourceObject = tempStore6;
                    break;
                case DataStores.Store7:
                    SourceObject = tempStore7;
                    break;
                case DataStores.Store8:
                    SourceObject = tempStore8;
                    break;
                case DataStores.Store9:
                    SourceObject = tempStore9;
                    break;
                default:
                    throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
            }

            if (SourceObject == null)
            {
                DestinationObject = null;
            }
            else
            {
                Type ControlType = SourceObject.GetType();

                do
                {
                    if (ControlType == null)
                    {
                        break;
                    }

                    string nameSpace = ControlType.Namespace;
                    if (nameSpace != null)  //Some obfuscation tools set the namespace to null...
                    {
                        if (nameSpace.StartsWith("System.Windows."))
                        {
                            break;
                        }

                        string APEType = GetAPETypeFromTypeInternal(ControlType.Namespace, ControlType.Name);

                        if (APEType != "")
                        {
                            break;
                        }
                    }

                    ControlType = ControlType.BaseType;
                }
                while (true);

                DestinationObject = ControlType;
            }

            switch (PtrMessage->DestinationStore)
            {
                case DataStores.Store0:
                    tempStore0 = DestinationObject;
                    break;
                case DataStores.Store1:
                    tempStore1 = DestinationObject;
                    break;
                case DataStores.Store2:
                    tempStore2 = DestinationObject;
                    break;
                case DataStores.Store3:
                    tempStore3 = DestinationObject;
                    break;
                case DataStores.Store4:
                    tempStore4 = DestinationObject;
                    break;
                case DataStores.Store5:
                    tempStore5 = DestinationObject;
                    break;
                case DataStores.Store6:
                    tempStore6 = DestinationObject;
                    break;
                case DataStores.Store7:
                    tempStore7 = DestinationObject;
                    break;
                case DataStores.Store8:
                    tempStore8 = DestinationObject;
                    break;
                case DataStores.Store9:
                    tempStore9 = DestinationObject;
                    break;
                default:
                    throw new Exception("Unsupported DestinationStore " + (PtrMessage->DestinationStore).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        //unsafe private void GetSystemType(int MessageNumber)
        //{
        //    object SourceObject;
        //    object DestinationObject;

        //    Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

        //    switch (PtrMessage->SourceStore)
        //    {
        //        case DataStores.Store0:
        //            SourceObject = tempStore0;
        //            break;
        //        case DataStores.Store1:
        //            SourceObject = tempStore1;
        //            break;
        //        case DataStores.Store2:
        //            SourceObject = tempStore2;
        //            break;
        //        case DataStores.Store3:
        //            SourceObject = tempStore3;
        //            break;
        //        case DataStores.Store4:
        //            SourceObject = tempStore4;
        //            break;
        //        case DataStores.Store5:
        //            SourceObject = tempStore5;
        //            break;
        //        case DataStores.Store6:
        //            SourceObject = tempStore6;
        //            break;
        //        case DataStores.Store7:
        //            SourceObject = tempStore7;
        //            break;
        //        case DataStores.Store8:
        //            SourceObject = tempStore8;
        //            break;
        //        case DataStores.Store9:
        //            SourceObject = tempStore9;
        //            break;
        //        default:
        //            throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
        //    }

        //    Type ControlType = SourceObject.GetType();

        //    while (!ControlType.Namespace.StartsWith("System.Windows."))
        //    {
        //        ControlType = ControlType.BaseType;
        //    }

        //    DestinationObject = ControlType;

        //    switch (PtrMessage->DestinationStore)
        //    {
        //        case DataStores.Store0:
        //            tempStore0 = DestinationObject;
        //            break;
        //        case DataStores.Store1:
        //            tempStore1 = DestinationObject;
        //            break;
        //        case DataStores.Store2:
        //            tempStore2 = DestinationObject;
        //            break;
        //        case DataStores.Store3:
        //            tempStore3 = DestinationObject;
        //            break;
        //        case DataStores.Store4:
        //            tempStore4 = DestinationObject;
        //            break;
        //        case DataStores.Store5:
        //            tempStore5 = DestinationObject;
        //            break;
        //        case DataStores.Store6:
        //            tempStore6 = DestinationObject;
        //            break;
        //        case DataStores.Store7:
        //            tempStore7 = DestinationObject;
        //            break;
        //        case DataStores.Store8:
        //            tempStore8 = DestinationObject;
        //            break;
        //        case DataStores.Store9:
        //            tempStore9 = DestinationObject;
        //            break;
        //        default:
        //            throw new Exception("Unsupported DestinationStore " + (PtrMessage->DestinationStore).ToString());
        //    }

        //    //cleanup the message
        //    PtrMessage->TypeCodeKey = 0;
        //    PtrMessage->NumberOfParameters = 0;
        //    PtrMessage->NameOffset = 0;
        //    PtrMessage->NameLength = 0;
        //    PtrMessage->Action = MessageAction.None;
        //}

        unsafe private void ReflectPoll(int MessageNumber)
        {
            object SourceObject;
            dynamic ItemToPollFor;
            dynamic ItemFound;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            switch (PtrMessage->Parameter.TypeCode[0])
            {
                case (int)ApeTypeCode.Boolean:
                    ItemToPollFor = PtrMessage->Parameter.Boolean[0];
                    break;
                case (int)ApeTypeCode.Char:
                    ItemToPollFor = PtrMessage->Parameter.Char[0];
                    break;
                case (int)ApeTypeCode.SByte:
                    ItemToPollFor = PtrMessage->Parameter.SByte[0];
                    break;
                case (int)ApeTypeCode.Byte:
                    ItemToPollFor = PtrMessage->Parameter.Byte[0];
                    break;
                case (int)ApeTypeCode.Int16:
                    ItemToPollFor = PtrMessage->Parameter.Int16[0];
                    break;
                case (int)ApeTypeCode.UInt16:
                    ItemToPollFor = PtrMessage->Parameter.UInt16[0];
                    break;
                case (int)ApeTypeCode.Int32:
                    ItemToPollFor = PtrMessage->Parameter.Int32[0];
                    break;
                case (int)ApeTypeCode.UInt32:
                    ItemToPollFor = PtrMessage->Parameter.UInt32[0];
                    break;
                case (int)ApeTypeCode.Int64:
                    ItemToPollFor = PtrMessage->Parameter.Int64[0];
                    break;
                case (int)ApeTypeCode.UInt64:
                    ItemToPollFor = PtrMessage->Parameter.UInt64[0];
                    break;
                case (int)ApeTypeCode.Single:
                    ItemToPollFor = PtrMessage->Parameter.Single[0];
                    break;
                case (int)ApeTypeCode.Double:
                    ItemToPollFor = PtrMessage->Parameter.Double[0];
                    break;
                case (int)ApeTypeCode.Decimal:
                    int[] DecimalBits = new int[4];
                    DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[0];
                    DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[0];
                    DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[0];
                    DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[0];
                    ItemToPollFor = new decimal(DecimalBits);
                    break;
                case (int)ApeTypeCode.DateTime:
                    ItemToPollFor = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[0]);
                    break;
                case (int)ApeTypeCode.String:
                    ItemToPollFor = m_TypeString;
                    if (PtrMessage->Parameter.StringLength[0] == -1)
                    {
                        string Empty = null;
                        ItemToPollFor = Empty;
                    }
                    else
                    {
                        ItemToPollFor = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[0]), 0, PtrMessage->Parameter.StringLength[0]);
                    }
                    break;
                case (int)ApeTypeCode.IntPtr:
                    ItemToPollFor = new IntPtr(PtrMessage->Parameter.IntPtr[0]);
                    break;
                default:
                    throw new Exception("Unsupported ApeTypeCode: " + ((int)PtrMessage->Parameter.TypeCode).ToString());
            }

            bool unsafeCrossThread = GetParameterBoolean(PtrMessage, 1);

            Type SourceType;
            string Name;
            Fasterflect.MemberGetter MemberGetter;

            switch (PtrMessage->SourceStore)
            {
                case DataStores.Store0:
                    SourceObject = tempStore0;
                    break;
                case DataStores.Store1:
                    SourceObject = tempStore1;
                    break;
                case DataStores.Store2:
                    SourceObject = tempStore2;
                    break;
                case DataStores.Store3:
                    SourceObject = tempStore3;
                    break;
                case DataStores.Store4:
                    SourceObject = tempStore4;
                    break;
                case DataStores.Store5:
                    SourceObject = tempStore5;
                    break;
                case DataStores.Store6:
                    SourceObject = tempStore6;
                    break;
                case DataStores.Store7:
                    SourceObject = tempStore7;
                    break;
                case DataStores.Store8:
                    SourceObject = tempStore8;
                    break;
                case DataStores.Store9:
                    SourceObject = tempStore9;
                    break;
                default:
                    throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
            }

            if (SourceObject == null)
            {
                throw new Exception("Source object is null");
            }
            else
            {
                SourceType = SourceObject.GetType();
                if (PtrMessage->NameLength == -1)
                {
                    string Empty = null;
                    Name = Empty;
                }
                else
                {
                    Name = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->NameOffset), 0, PtrMessage->NameLength);
                }

                //Get the value
                switch (PtrMessage->MemberType)
                {
                    case MemberTypes.Field:
                        MemberGetterCache.GetFromList(SourceType.TypeHandle.Value, Name, out MemberGetter);
                        if (MemberGetter == null)
                        {
                            MemberGetter = SourceType.DelegateForGetFieldValue(Name);
                            MemberGetterCache.AddToList(SourceType.TypeHandle.Value, Name, MemberGetter);
                        }
                        break;
                    case MemberTypes.Property:
                        MemberGetterCache.GetFromList(SourceType.TypeHandle.Value, Name, out MemberGetter);
                        if (MemberGetter == null)
                        {
                            MemberGetter = SourceType.DelegateForGetPropertyValue(Name);
                            MemberGetterCache.AddToList(SourceType.TypeHandle.Value, Name, MemberGetter);
                        }
                        break;
                    default:
                        throw new Exception("Unsupported member type: " + (PtrMessage->MemberType).ToString());
                }

                Stopwatch timer = Stopwatch.StartNew();
                while (true)
                {
                    if (unsafeCrossThread)
                    {
                        ItemFound = MemberGetter(SourceObject.WrapIfValueType());
                    }
                    else
                    {
                        ItemFound = ((WF.Control)tempStore0).Invoke(MemberGetter, SourceObject.WrapIfValueType());
                    }
                    
                    if (ItemFound == ItemToPollFor)
                    {
                        timer.Stop();
                        break;
                    }

                    if (timer.ElapsedMilliseconds > m_TimeOut)
                    {
                        timer.Stop();
                        //TODO pass back a fail rather than throwing an exception
                        throw new Exception("Polling failed to match item");
                    }

                    Thread.Sleep(50);
                }

                MemberGetter = null;
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        unsafe private void ConvertType(int MessageNumber)
        {
            object SourceObject;
            object DestinationObject;
            IntPtr datastoreTypeHandle = IntPtr.Zero;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            string typeFullName;
            if (PtrMessage->NameLength == -1)
            {
                string Empty = null;
                typeFullName = Empty;
            }
            else
            {
                typeFullName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->NameOffset), 0, PtrMessage->NameLength);
            }

            switch (PtrMessage->SourceStore)
            {
                case DataStores.Store0:
                    SourceObject = tempStore0;
                    break;
                case DataStores.Store1:
                    SourceObject = tempStore1;
                    break;
                case DataStores.Store2:
                    SourceObject = tempStore2;
                    break;
                case DataStores.Store3:
                    SourceObject = tempStore3;
                    break;
                case DataStores.Store4:
                    SourceObject = tempStore4;
                    break;
                case DataStores.Store5:
                    SourceObject = tempStore5;
                    break;
                case DataStores.Store6:
                    SourceObject = tempStore6;
                    break;
                case DataStores.Store7:
                    SourceObject = tempStore7;
                    break;
                case DataStores.Store8:
                    SourceObject = tempStore8;
                    break;
                case DataStores.Store9:
                    SourceObject = tempStore9;
                    break;
                default:
                    throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
            }

            if (SourceObject == null)
            {
                DestinationObject = null;
            }
            else
            {
                //Find the type we want to convert the object to
                Type theType = AppDomain.CurrentDomain.GetAssemblies().SelectMany(x => x.GetTypes()).FirstOrDefault(x => x.FullName == typeFullName);
                //setup the delegate parameters
                object[] convertParameters = { theType, SourceObject };
                //call the delegate on the correct thread
                DestinationObject = ((WF.Control)tempStore0).Invoke(m_ConvertTypeDelegater, convertParameters);
            }

            switch (PtrMessage->DestinationStore)
            {
                case DataStores.Store0:
                    tempStore0 = DestinationObject;
                    break;
                case DataStores.Store1:
                    tempStore1 = DestinationObject;
                    break;
                case DataStores.Store2:
                    tempStore2 = DestinationObject;
                    break;
                case DataStores.Store3:
                    tempStore3 = DestinationObject;
                    break;
                case DataStores.Store4:
                    tempStore4 = DestinationObject;
                    break;
                case DataStores.Store5:
                    tempStore5 = DestinationObject;
                    break;
                case DataStores.Store6:
                    tempStore6 = DestinationObject;
                    break;
                case DataStores.Store7:
                    tempStore7 = DestinationObject;
                    break;
                case DataStores.Store8:
                    tempStore8 = DestinationObject;
                    break;
                case DataStores.Store9:
                    tempStore9 = DestinationObject;
                    break;
                default:
                    throw new Exception("Unsupported DestinationStore " + (PtrMessage->DestinationStore).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        private object Cast(Type Type, object data)
        {
            var DataParam = Expression.Parameter(typeof(object), "data");
            var Body = Expression.Block(Expression.Convert(Expression.Convert(DataParam, data.GetType()), Type));
            var Run = Expression.Lambda(Body, DataParam).Compile();
            var ret = Run.DynamicInvoke(data);
            return ret;
        }

        private string GetText(WF.Control theControl)
        {
            return theControl.Text;
        }

        private string GetAccessibilityObjectName(WF.Control theControl)
        {
            return theControl.AccessibilityObject.Name;
        }

        unsafe private void Reflect(int MessageNumber)
        {
            object SourceObject;
            object DestinationObject;
            StringBuilder datastoreTypes = new StringBuilder(1024);

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            Type SourceType;
            switch (PtrMessage->SourceStore)
            {
                case DataStores.Store0:
                    SourceObject = tempStore0;
                    break;
                case DataStores.Store1:
                    SourceObject = tempStore1;
                    break;
                case DataStores.Store2:
                    SourceObject = tempStore2;
                    break;
                case DataStores.Store3:
                    SourceObject = tempStore3;
                    break;
                case DataStores.Store4:
                    SourceObject = tempStore4;
                    break;
                case DataStores.Store5:
                    SourceObject = tempStore5;
                    break;
                case DataStores.Store6:
                    SourceObject = tempStore6;
                    break;
                case DataStores.Store7:
                    SourceObject = tempStore7;
                    break;
                case DataStores.Store8:
                    SourceObject = tempStore8;
                    break;
                case DataStores.Store9:
                    SourceObject = tempStore9;
                    break;
                default:
                    throw new Exception("Unsupported SourceStore " + (PtrMessage->SourceStore).ToString());
            }

            bool releaseComObject = false;
            if (SourceObject == null)
            {
                DestinationObject = null;
            }
            else
            {
                SourceType = SourceObject.GetType();
                Type[] ParametersType = null;
                object[] ParametersObject = new object[PtrMessage->NumberOfParameters];

                ParametersTypeCache.GetFromList(PtrMessage->TypeCodeKey, out ParametersType);

                if (ParametersType == null)
                {
                    ParametersType = new Type[PtrMessage->NumberOfParameters];

                    for (int i = 0; i < PtrMessage->NumberOfParameters; i++)
                    {
                        switch ((int)PtrMessage->Parameter.TypeCode[i])
                        {
                            case (int)ApeTypeCode.Boolean:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeBoolean;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeBooleanByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Boolean[i];
                                break;
                            case (int)ApeTypeCode.Char:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeChar;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeCharByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Char[i];
                                break;
                            case (int)ApeTypeCode.SByte:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeSByte;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeSByteByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.SByte[i];
                                break;
                            case (int)ApeTypeCode.Byte:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeByte;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeByteByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Byte[i];
                                break;
                            case (int)ApeTypeCode.Int16:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeInt16;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeInt16ByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Int16[i];
                                break;
                            case (int)ApeTypeCode.UInt16:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeUInt16;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeUInt16ByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.UInt16[i];
                                break;
                            case (int)ApeTypeCode.Int32:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeInt32;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeInt32ByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Int32[i];
                                break;
                            case (int)ApeTypeCode.UInt32:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeUInt32;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeUInt32ByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.UInt32[i];
                                break;
                            case (int)ApeTypeCode.Int64:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeInt64;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeInt64ByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Int64[i];
                                break;
                            case (int)ApeTypeCode.UInt64:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeUInt64;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeUInt64ByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.UInt64[i];
                                break;
                            case (int)ApeTypeCode.Single:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeSingle;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeSingleByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Single[i];
                                break;
                            case (int)ApeTypeCode.Double:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeDouble;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeDoubleByRef;
                                }
                                ParametersObject[i] = PtrMessage->Parameter.Double[i];
                                break;
                            case (int)ApeTypeCode.Decimal:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeDecimal;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeDecimalByRef;
                                }
                                int[] DecimalBits = new int[4];
                                DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[i];
                                DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[i];
                                DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[i];
                                DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[i];
                                ParametersObject[i] = new decimal(DecimalBits);
                                break;
                            case (int)ApeTypeCode.DateTime:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeDateTime;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeDateTimeByRef;
                                }
                                ParametersObject[i] = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[i]);
                                break;
                            case (int)ApeTypeCode.String:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeString;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeStringByRef;
                                }
                                if (PtrMessage->Parameter.StringLength[i] == -1)
                                {
                                    string Empty = null;
                                    ParametersObject[i] = Empty;
                                }
                                else
                                {
                                    ParametersObject[i] = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[i]), 0, PtrMessage->Parameter.StringLength[i]);
                                }
                                break;
                            case (int)ApeTypeCode.IntPtr:
                                if (PtrMessage->Parameter.ParameterType[i] == (int)ParameterType.In)
                                {
                                    ParametersType[i] = m_TypeIntPtr;
                                }
                                else
                                {
                                    ParametersType[i] = m_TypeIntPtrByRef;
                                }
                                ParametersObject[i] = new IntPtr(PtrMessage->Parameter.IntPtr[i]);
                                break;
                            case (int)ApeTypeCode.DataStore:
                                int datastoreNumber = PtrMessage->Parameter.Int32[i];
                                switch (datastoreNumber)
                                {
                                    case 0:
                                        ParametersObject[i] = tempStore0;
                                        break;
                                    case 1:
                                        ParametersObject[i] = tempStore1;
                                        break;
                                    case 2:
                                        ParametersObject[i] = tempStore2;
                                        break;
                                    case 3:
                                        ParametersObject[i] = tempStore3;
                                        break;
                                    case 4:
                                        ParametersObject[i] = tempStore4;
                                        break;
                                    case 5:
                                        ParametersObject[i] = tempStore5;
                                        break;
                                    case 6:
                                        ParametersObject[i] = tempStore6;
                                        break;
                                    case 7:
                                        ParametersObject[i] = tempStore7;
                                        break;
                                    case 8:
                                        ParametersObject[i] = tempStore8;
                                        break;
                                    case 9:
                                        ParametersObject[i] = tempStore9;
                                        break;
                                    default:
                                        throw new Exception("Unsupported SourceStore " + datastoreNumber.ToString());
                                }

                                string parameterTypeText;
                                string parameterType = null;
                                switch ((ParameterType)PtrMessage->Parameter.ParameterType[i])
                                {
                                    case ParameterType.In:
                                        parameterTypeText = " In";
                                        ParametersType[i] = ParametersObject[i].GetType();
                                        break;
                                    case ParameterType.Out:
                                        parameterTypeText = " Out";
                                        if (PtrMessage->Parameter.StringLength[i] != -1)
                                        {
                                            parameterType = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[i]), 0, PtrMessage->Parameter.StringLength[i]);
                                        }

                                        Type typeContainingParameterType;
                                        ParameterTypeCache.GetFromList(parameterType + parameterTypeText, out typeContainingParameterType);
                                        if (typeContainingParameterType == null)
                                        {
                                            typeContainingParameterType = SourceType.Assembly.GetTypes().FirstOrDefault(x => x.FullName == parameterType);
                                            if (typeContainingParameterType == null)
                                            {
                                                typeContainingParameterType = SourceType.Assembly.GetReferencedAssemblies().Select(x => Assembly.Load(x)).SelectMany(x => x.GetTypes()).FirstOrDefault(x => x.FullName == parameterType);
                                                if (typeContainingParameterType == null)
                                                {
                                                    throw new Exception("Failed to find type containing " + parameterType);
                                                }
                                            }
                                            typeContainingParameterType = typeContainingParameterType.MakeByRefType();
                                            ParameterTypeCache.AddToList(parameterType + parameterTypeText, typeContainingParameterType);
                                        }
                                        ParametersType[i] = typeContainingParameterType;
                                        break;
                                    case ParameterType.Ref:
                                        parameterTypeText = " Ref";
                                        ParametersType[i] = ParametersObject[i].GetType().MakeByRefType();
                                        break;
                                    default:
                                        throw new Exception("Unsupported ape parameter type: " + ((ParameterType)PtrMessage->Parameter.ParameterType[i]).ToString());
                                }

                                datastoreTypes.AppendLine(ParametersType[i].Namespace);
                                datastoreTypes.AppendLine(".");
                                datastoreTypes.AppendLine(ParametersType[i].Name);
                                datastoreTypes.AppendLine(parameterTypeText);
                                break;
                            default:
                                throw new Exception("Unsupported ApeTypeCode: " + PtrMessage->Parameter.TypeCode[i].ToString());
                        }
                    }
                    if (datastoreTypes.Length == 0)  //none of the parameters are a datastore type
                    {
                        ParametersTypeCache.AddToList(PtrMessage->TypeCodeKey, ParametersType);
                    }
                }
                else
                {
                    for (int i = 0; i < PtrMessage->NumberOfParameters; i++)
                    {
                        switch ((int)PtrMessage->Parameter.TypeCode[i])
                        {
                            case (int)ApeTypeCode.Boolean:
                                ParametersObject[i] = PtrMessage->Parameter.Boolean[i];
                                break;
                            case (int)ApeTypeCode.Char:
                                ParametersObject[i] = PtrMessage->Parameter.Char[i];
                                break;
                            case (int)ApeTypeCode.SByte:
                                ParametersObject[i] = PtrMessage->Parameter.SByte[i];
                                break;
                            case (int)ApeTypeCode.Byte:
                                ParametersObject[i] = PtrMessage->Parameter.Byte[i];
                                break;
                            case (int)ApeTypeCode.Int16:
                                ParametersObject[i] = PtrMessage->Parameter.Int16[i];
                                break;
                            case (int)ApeTypeCode.UInt16:
                                ParametersObject[i] = PtrMessage->Parameter.UInt16[i];
                                break;
                            case (int)ApeTypeCode.Int32:
                                ParametersObject[i] = PtrMessage->Parameter.Int32[i];
                                break;
                            case (int)ApeTypeCode.UInt32:
                                ParametersObject[i] = PtrMessage->Parameter.UInt32[i];
                                break;
                            case (int)ApeTypeCode.Int64:
                                ParametersObject[i] = PtrMessage->Parameter.Int64[i];
                                break;
                            case (int)ApeTypeCode.UInt64:
                                ParametersObject[i] = PtrMessage->Parameter.UInt64[i];
                                break;
                            case (int)ApeTypeCode.Single:
                                ParametersObject[i] = PtrMessage->Parameter.Single[i];
                                break;
                            case (int)ApeTypeCode.Double:
                                ParametersObject[i] = PtrMessage->Parameter.Double[i];
                                break;
                            case (int)ApeTypeCode.Decimal:
                                int[] DecimalBits = new int[4];
                                DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[i];
                                DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[i];
                                DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[i];
                                DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[i];
                                ParametersObject[i] = new decimal(DecimalBits);
                                break;
                            case (int)ApeTypeCode.DateTime:
                                ParametersObject[i] = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[i]);
                                break;
                            case (int)ApeTypeCode.String:
                                ParametersType[i] = m_TypeString;
                                if (PtrMessage->Parameter.StringLength[i] == -1)
                                {
                                    string Empty = null;
                                    ParametersObject[i] = Empty;
                                }
                                else
                                {
                                    ParametersObject[i] = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[i]), 0, PtrMessage->Parameter.StringLength[i]);
                                }
                                break;
                            case (int)ApeTypeCode.IntPtr:
                                ParametersObject[i] = new IntPtr(PtrMessage->Parameter.IntPtr[i]);
                                break;
                            default:
                                throw new Exception("Unsupported ApeTypeCode: " + PtrMessage->Parameter.TypeCode[i].ToString());
                        }
                    }
                }

                string Name;
                Fasterflect.ConstructorInvoker ConstructorInvoker;
                Fasterflect.MethodInvoker MethodInvoker;
                Fasterflect.MemberGetter MemberGetter;
                //Fasterflect.ArrayElementGetter ArrayElementGetter;

                if (PtrMessage->NameLength == -1)
                {
                    string Empty = null;
                    Name = Empty;
                }
                else
                {
                    Name = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->NameOffset), 0, PtrMessage->NameLength);
                }

                if (Marshal.IsComObject(SourceObject))
                {
                    WF.Control invokeControl;
                    if (Marshal.IsComObject(tempStore0))
                    {
                        invokeControl = Ax.InvokeForm;
                    }
                    else
                    {
                        invokeControl = (WF.Control)tempStore0;
                    }

                    object[] comParameters = new object[3];
                    comParameters[0] = Name;
                    comParameters[1] = SourceObject;
                    comParameters[2] = ParametersObject;

                    DestinationObject = invokeControl.Invoke(m_ComReflectDelegater, comParameters);

                    if (DestinationObject != null && Marshal.IsComObject(DestinationObject))
                    {
                        releaseComObject = true;
                    }
                }
                else
                {
                    //Get the value
                    switch (PtrMessage->MemberType)
                    {
                        case MemberTypes.Constructor:
                            ConstructorInvokerCache.GetFromList(Name, PtrMessage->TypeCodeKey, datastoreTypes.ToString(), out ConstructorInvoker);
                            if (ConstructorInvoker == null)
                            {
                                Type typeContainingConstructor;
                                typeContainingConstructor = SourceType.Assembly.GetTypes().FirstOrDefault(x => x.FullName == Name);
                                if (typeContainingConstructor == null)
                                {
                                    typeContainingConstructor = SourceType.Assembly.GetReferencedAssemblies().Select(x => Assembly.Load(x)).SelectMany(x => x.GetTypes()).FirstOrDefault(x => x.FullName == Name);
                                    if (typeContainingConstructor == null)
                                    {
                                        throw new Exception("Failed to find type containing " + Name);
                                    }
                                }
                                ConstructorInvoker = typeContainingConstructor.DelegateForCreateInstance(Flags.Default, ParametersType);
                                ConstructorInvokerCache.AddToList(Name, PtrMessage->TypeCodeKey, datastoreTypes.ToString(), ConstructorInvoker);
                            }
                            DestinationObject = ConstructorInvoker.Invoke(ParametersObject);
                            //DestinationObject = ((Control)tempStore0).Invoke(ConstructorInvoker, SourceObject.WrapIfValueType(), ParametersObject);
                            ConstructorInvoker = null;
                            break;
                        case MemberTypes.Field:
                            MemberGetterCache.GetFromList(SourceType.TypeHandle.Value, Name, out MemberGetter);
                            if (MemberGetter == null)
                            {
                                MemberGetter = SourceType.DelegateForGetFieldValue(Name, Flags.Default);
                                MemberGetterCache.AddToList(SourceType.TypeHandle.Value, Name, MemberGetter);
                            }
                            DestinationObject = ((WF.Control)tempStore0).Invoke(MemberGetter, SourceObject.WrapIfValueType());
                            MemberGetter = null;
                            break;
                        case MemberTypes.Property:
                            if (ParametersType.Length == 0)
                            {
                                MemberGetterCache.GetFromList(SourceType.TypeHandle.Value, Name, out MemberGetter);
                                if (MemberGetter == null)
                                {
                                    MemberGetter = SourceType.DelegateForGetPropertyValue(Name, Flags.Default);
                                    MemberGetterCache.AddToList(SourceType.TypeHandle.Value, Name, MemberGetter);
                                }
                                DestinationObject = ((WF.Control)tempStore0).Invoke((Delegate)MemberGetter, SourceObject.WrapIfValueType());
                                MemberGetter = null;
                            }
                            else
                            {
                                MethodInvokerCache.GetFromList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypes.ToString(), out MethodInvoker);
                                if (MethodInvoker == null)
                                {
                                    MethodInvoker = SourceType.DelegateForGetIndexer(Flags.Default, ParametersType);
                                    MethodInvokerCache.AddToList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypes.ToString(), MethodInvoker);
                                }
                                DestinationObject = ((WF.Control)tempStore0).Invoke(MethodInvoker, SourceObject.WrapIfValueType(), ParametersObject);
                                MethodInvoker = null;
                            }
                            break;
                        case MemberTypes.Method:
                            //Reflection doesn't seem to work on Enums so access it directly
                            if (SourceType.IsEnum && Name == "ToString")
                            {
                                DestinationObject = SourceObject.ToString();
                            }
                            else if (Name == "GetType")
                            {
                                DestinationObject = SourceType;
                            }
                            else
                            {
                                MethodInvokerCache.GetFromList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypes.ToString(), out MethodInvoker);
                                if (MethodInvoker == null)
                                {
                                    MethodInvoker = SourceType.DelegateForCallMethod(Name, Flags.Default, ParametersType);
                                    MethodInvokerCache.AddToList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypes.ToString(), MethodInvoker);
                                }
                                DestinationObject = ((WF.Control)tempStore0).Invoke(MethodInvoker, SourceObject.WrapIfValueType(), ParametersObject);
                            }
                            MethodInvoker = null;
                            break;
                        default:
                            throw new Exception("Unsupported member type: " + (PtrMessage->MemberType).ToString());
                    }

                    // Copy back the out parameters to the tempstore
                    if (datastoreTypes != null)
                    {
                        for (int i = 0; i < PtrMessage->NumberOfParameters; i++)
                        {
                            switch ((int)PtrMessage->Parameter.TypeCode[i])
                            {
                                case (int)ApeTypeCode.DataStore:
                                    int datastoreNumber = PtrMessage->Parameter.Int32[i];
                                    if (PtrMessage->Parameter.ParameterType[i] != (int)ParameterType.In)
                                    {
                                        switch (datastoreNumber)
                                        {
                                            case 0:
                                                tempStore0 = ParametersObject[i];
                                                break;
                                            case 1:
                                                tempStore1 = ParametersObject[i];
                                                break;
                                            case 2:
                                                tempStore2 = ParametersObject[i];
                                                break;
                                            case 3:
                                                tempStore3 = ParametersObject[i];
                                                break;
                                            case 4:
                                                tempStore4 = ParametersObject[i];
                                                break;
                                            case 5:
                                                tempStore5 = ParametersObject[i];
                                                break;
                                            case 6:
                                                tempStore6 = ParametersObject[i];
                                                break;
                                            case 7:
                                                tempStore7 = ParametersObject[i];
                                                break;
                                            case 8:
                                                tempStore8 = ParametersObject[i];
                                                break;
                                            case 9:
                                                tempStore9 = ParametersObject[i];
                                                break;
                                        }
                                    }
                                    break;
                            }
                        }
                    }
                }
            }
           
            switch (PtrMessage->DestinationStore)
            {
                case DataStores.Store0:
                    tempStore0 = DestinationObject;
                    tempStore0ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store1:
                    tempStore1 = DestinationObject;
                    tempStore1ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store2:
                    tempStore2 = DestinationObject;
                    tempStore2ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store3:
                    tempStore3 = DestinationObject;
                    tempStore3ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store4:
                    tempStore4 = DestinationObject;
                    tempStore4ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store5:
                    tempStore5 = DestinationObject;
                    tempStore5ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store6:
                    tempStore6 = DestinationObject;
                    tempStore6ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store7:
                    tempStore7 = DestinationObject;
                    tempStore7ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store8:
                    tempStore8 = DestinationObject;
                    tempStore8ReleaseComObject = releaseComObject;
                    break;
                case DataStores.Store9:
                    tempStore9 = DestinationObject;
                    tempStore9ReleaseComObject = releaseComObject;
                    break;
                default:
                    throw new Exception("Unsupported DestinationStore " + (PtrMessage->DestinationStore).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        public string GetWindowTextViaWindowMessage(IntPtr Handle)
        {
            IntPtr MessageResult;
            IntPtr SendResult;

            SendResult = NM.SendMessageTimeout(Handle, NM.WM_GETTEXTLENGTH, IntPtr.Zero, IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out MessageResult);
            if (SendResult != IntPtr.Zero)  //Succeeded
            {
                int BufferSize = (int)(MessageResult.ToInt32() * 2) + 1;
                IntPtr StringBuffer = Marshal.AllocHGlobal(BufferSize);
                SendResult = NM.SendMessageTimeout(Handle, NM.WM_GETTEXT, MessageResult + 1, StringBuffer, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out MessageResult);

                string WindowText = null;
                if (SendResult != IntPtr.Zero)  //Succeeded
                {
                    WindowText = Marshal.PtrToStringAuto(StringBuffer);
                }

                Marshal.FreeHGlobal(StringBuffer);
                return WindowText;
            }

            return null;
        }

        private string GetTechnologyType(Type theType)
        {
            while (theType != null)
            {
                switch (theType.Namespace)
                {
                    case "System.Windows.Forms":
                        return "Windows Forms (WinForms)";
                    case "System.Windows":
                        return "Windows Presentation Foundation (WPF)";
                }

                theType = theType.BaseType;
            }

            return "Windows Native";
        }

        unsafe private void AddReturnValue(Parameter p0)
        {
            m_PtrMessageStore->NumberOfMessages++;
        }

        public uint TimeOut
        {
            get
            {
                return m_TimeOut;
            }
            set
            {
                m_TimeOut = value;

                AddFirstMessageSetTimeOuts();
                SendMessages(EventSet.APE);
                WaitForMessages(EventSet.APE);
            }
        }

        public uint HangTimeOut
        {
            get
            {
                return m_HangTimeOut;
            }
            set
            {
                m_HangTimeOut = value;
            }
        }
    }
}

