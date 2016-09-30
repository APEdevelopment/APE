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
        public string Description;
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

    public partial class APEIPC
    {
        private const int ERROR_ALREADY_EXISTS = 183;
        private const int StringSpaceBytes = 1048576;

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
        private MemberGetterCircularList MemberGetterCache = new MemberGetterCircularList(10);
        private MethodInvokerCircularList MethodInvokerCache = new MethodInvokerCircularList(10);
        private ConstructorInvokerCircularList ConstructorInvokerCache = new ConstructorInvokerCircularList(10);

        //Cache all the primitive types (for performance)
        private Type m_TypeBoolean = typeof(Boolean);
        private Type m_TypeChar = typeof(Char);
        private Type m_TypeSByte = typeof(SByte);
        private Type m_TypeByte = typeof(Byte);
        private Type m_TypeInt16 = typeof(Int16);
        private Type m_TypeUInt16 = typeof(UInt16);
        private Type m_TypeInt32 = typeof(Int32);
        private Type m_TypeUInt32 = typeof(UInt32);
        private Type m_TypeInt64 = typeof(Int64);
        private Type m_TypeUInt64 = typeof(UInt64);
        private Type m_TypeSingle = typeof(Single);
        private Type m_TypeDouble = typeof(Double);
        private Type m_TypeDecimal = typeof(Decimal);
        private Type m_TypeDateTime = typeof(DateTime);
        private Type m_TypeString = typeof(String);
        private Type m_TypeIntPtr = typeof(IntPtr);

        private object tempStore0;
        private object tempStore1;
        private object tempStore2;
        private object tempStore3;
        private object tempStore4;
        private object tempStore5;
        private object tempStore6;
        private object tempStore7;
        private object tempStore8;
        private object tempStore9;

        private Boolean m_DoneFind;
        private Boolean m_DoneQuery;
        private Boolean m_DoneGet;

        EventWaitHandle m_eventIPC;
        IntPtr m_Handle;
        string m_Name;
        string m_Text;

        private bool m_WPF = false;

        private delegate string GetTextDelegate(WF.Control theControl);
        private GetTextDelegate m_GetTextDelegater;
        private delegate object ConvertTypeDelegate(Type theTyoe, object theObject);
        private ConvertTypeDelegate m_ConvertTypeDelegater;
        private delegate void GetWPFHandleAndNameAndTitleDelegate(WPF.Window theWindow);
        private GetWPFHandleAndNameAndTitleDelegate m_GetWPFHandleAndNameAndTitleDelegater;
        private NM.EnumWindow EnumThreadProcedue;

        List<IntPtr> m_AllControls;
        EventSet Side;
        Process ApeProcess = null;
        Process AUTProcess = null;
        bool m_Abort = false;
        uint m_TimeOut = 0;
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

        private unsafe void InjectAPEIPC(Process AUTProcess, string AppDomain)
        {
            string AUTProcessId = AUTProcess.Id.ToString();
            string Path = Directory.GetCurrentDirectory();
            string Method = "LoadAPEIPC";
            string APEProcessId = Process.GetCurrentProcess().Id.ToString();
            
            this.AUTProcess = AUTProcess;

            m_HandleMemoryMappedFileStringStore = NM.CreateFileMapping((IntPtr)(NM.INVALID_HANDLE_VALUE), (IntPtr)0, NM.FileMapProtection.PageReadWrite, 0, StringSpaceBytes, APEProcessId + "_String_" + AppDomain + "_" + AUTProcessId);

            if (m_HandleMemoryMappedFileStringStore == null)
            {
                throw new Exception("Failed to create memory mapped file error: " + Marshal.GetLastWin32Error().ToString());
            }

            if (Marshal.GetLastWin32Error() != ERROR_ALREADY_EXISTS)
            {
                Process Injector = new Process();
                RegistryKey key = Registry.CurrentUser.CreateSubKey("Software").CreateSubKey("APE");
                try
                {
                    key.SetValue(APEProcessId + "_Path_" + AUTProcessId, Path);
                    key.SetValue(APEProcessId + "_AppDomain_" + AUTProcessId, AppDomain);

                    if (NM.Is32BitProcess(AUTProcess))
                    {
                        string Assembly = "APE.Loader_x86.dll";
                        Injector.StartInfo = new ProcessStartInfo("APE.Injector_x86.exe", AUTProcessId + " " + Assembly + " " + Method + " " + APEProcessId);
                    }
                    else
                    {
                        string Assembly = "APE.Loader_x64.dll";
                        Injector.StartInfo = new ProcessStartInfo("APE.Injector_x64.exe", AUTProcessId + " " + Assembly + " " + Method + " " + APEProcessId);
                    }
                    Injector.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    Injector.Start();

                    Injector.WaitForExit();
                }
                finally
                {
                    key.DeleteValue(APEProcessId + "_Path_" + AUTProcessId, false);
                    key.DeleteValue(APEProcessId + "_AppDomain_" + AUTProcessId, false);
                    key.Close();
                }

                if (Injector.ExitCode != 0)
                {
                    throw new Exception("Injector exited with status: " + Injector.ExitCode.ToString());
                }
            }

            m_IntPtrMemoryMappedFileViewStringStore = NM.MapViewOfFile(m_HandleMemoryMappedFileStringStore, NM.FileMapAccess.FileMapAllAccess, 0, 0, (UIntPtr)StringSpaceBytes);
            m_HandleMemoryMappedFileMessageStore = NM.CreateFileMapping((IntPtr)(NM.INVALID_HANDLE_VALUE), (IntPtr)0, NM.FileMapProtection.PageReadWrite, 0, (uint)sizeof(MessageStore), APEProcessId + "_Message_" + AppDomain + "_" + AUTProcessId);
            m_IntPtrMemoryMappedFileViewMessageStore = NM.MapViewOfFile(m_HandleMemoryMappedFileMessageStore, NM.FileMapAccess.FileMapAllAccess, 0, 0, (UIntPtr)sizeof(MessageStore));
            m_PtrMessageStore = (MessageStore*)m_IntPtrMemoryMappedFileViewMessageStore.ToPointer();

            m_eventIPC = new EventWaitHandle(false, EventResetMode.AutoReset, APEProcessId + "_EventIPC_" + AppDomain + "_" + AUTProcessId);
            Side = EventSet.APE;
        }

        [EditorBrowsable(EditorBrowsableState.Never)]
        public unsafe APEIPC(int APEPID, string AppDomainToLoadInto, bool WPF)
        {
            try
            {
                Thread myThread = new Thread(() => ProcessMessages(APEPID, AppDomainToLoadInto, WPF));
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
            //MessageBox.Show("unloading " + Side.ToString());
            if (Side == EventSet.APE)
            {
                if (m_IntPtrMemoryMappedFileViewStringStore != IntPtr.Zero)
                {
                    if (!AUTProcess.HasExited)
                    {
                        this.AddFirstMessageRemoveFileMapping();
                        this.SendMessages(EventSet.APE);
                    }
                }
            }
            RemoveFileMapping();
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

        private unsafe void GarbageCollect(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GarbageCollect must be first message");
            }

            int generation = GetParameterInt32(ptrMessage, 0);
            CleanUpMessage(ptrMessage);

            //// Make sure the the logs are all flushed.
            //Assembly[] allAssemblies = AppDomain.CurrentDomain.GetAssemblies();
            //for (int i = 0; i < allAssemblies.Count(); i++)
            //{
            //    if (allAssemblies[i].FullName.StartsWith("Lz.Utility.Diagnostics.LogManager, "))
            //    {
            //        allAssemblies[i].CallMethod("FlushAllLogs", null);
            //        break;
            //    }
            //}

            // Do a single GC (more than this can have undesirable effects on performance)
            if (generation == GC.MaxGeneration)
            {
                GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce;
            }
            GC.Collect(generation);
            GC.WaitForPendingFinalizers();

            Thread.Sleep(150);  //A small sleep after GC seems to make the performance timings more accurate
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

        public unsafe void GetTitleBarItemRectangle(int MessageNumber)
        {
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

            // p2  = TitleBarStateElement
            NM.TitleBarStateElement Item;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)ApeTypeCode.Int32)
            {
                Item = (NM.TitleBarStateElement)PtrMessage->Parameter.Int32[1];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.Int32 got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            NM.TITLEBARINFOEX tbi = new NM.TITLEBARINFOEX();
            tbi.cbSize = Marshal.SizeOf(typeof(NM.TITLEBARINFOEX));

            // Send the WM_GETTITLEBARINFOEX message
            IntPtr Result;
            IntPtr Return;
            Return = NM.SendMessageTimeout(Handle, NM.WM_GETTITLEBARINFOEX, IntPtr.Zero, ref tbi, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out Result);

            uint State;
            int Left;
            int Top;
            int Right;
            int Bottom;

            switch (Item)
            {
                case NM.TitleBarStateElement.TitleBar:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.TitleBar];
                    Top = tbi.rcTitleBar.top;
                    Left = tbi.rcTitleBar.left;
                    Bottom = tbi.rcTitleBar.bottom;
                    Right = tbi.rcTitleBar.right;
                    break;
                case NM.TitleBarStateElement.Minimize:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Minimize];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].right;
                    break;
                case NM.TitleBarStateElement.Maximize:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Maximize];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].right;
                    break;
                case NM.TitleBarStateElement.Help:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Help];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Help].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Help].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Help].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Help].right;
                    break;
                case NM.TitleBarStateElement.Close:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Close];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Close].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Close].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Close].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Close].right;
                    break;
                default:
                    throw new Exception("Unknown titlebar element: " + Item.ToString());
            }

            AddReturnValue(new Parameter(this, State));
            AddReturnValue(new Parameter(this, Top));
            AddReturnValue(new Parameter(this, Left));
            AddReturnValue(new Parameter(this, Bottom));
            AddReturnValue(new Parameter(this, Right));
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

            if (m_eventIPC != null)
            {
                m_eventIPC.Dispose();
            }
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
            if (WhoIsSending == EventSet.APE)
            {
                if (m_DoneQuery == true)
                {
                    //TODO remove this?
                    //if (m_DoneGet == false)
                    //{
                    //    throw new Exception("Must get atleast one value after performing a query");
                    //}
                }

                // Check AUT is still running
                if (AUTProcess.HasExited)
                {
                    throw new Exception("AUT has exited!");
                }
            }

            //signal the other process
            m_eventIPC.Set();

            //yield the current process until the other process wakes up and notifies us via the mmf
            //if we didn't do this but instead waited we can end up in a situation where we immediatly wake up
            //instead of the remote process due to kernal mode stuff preventing the other process from taking
            //the signal
            Thread.CurrentThread.Priority = ThreadPriority.BelowNormal;
            while (m_PtrMessageStore->LastWake == WhoIsSending)
            {
                Thread.Sleep(0);    //better performance than Thread.Yield();
            }
            Thread.CurrentThread.Priority = ThreadPriority.Normal;
        }

        unsafe public void WaitForMessages(EventSet WhoIsWaiting)
        {
            //we know the other process is awake so wait for it to signal us via the event
            if (Side == EventSet.APE)
            {
                //m_eventIPC.WaitOne();
                while (true)
                {
                    if (m_eventIPC.WaitOne(100))
                    {
                        m_PtrMessageStore->LastWake = WhoIsWaiting;
                        GetResultMessage();
                        break;
                    }
                    else
                    {
                        //check if the AUT is still running (if it crashed it wont ask us to unload)
                        if (AUTProcess == null || AUTProcess.HasExited)
                        {
                            m_Abort = true;
                            break;
                        }
                    }
                }
            }
            else
            {
                while (true)
                {
                    if (m_eventIPC.WaitOne(30000))
                    {
                        break;
                    }
                    else
                    {
                        //check if the APE is still running (if it crashed it wont ask us to unload)
                        //if its being debugged HasExited will throw a permission denined exception
                        try
                        {
                            if (ApeProcess == null || ApeProcess.HasExited)
                            {
                                m_Abort = true;
                                break;
                            }
                        }
                        catch
                        {
                        }
                    }
                }
            }

            //writes to the mmf to let the other process know we have woken up
            m_PtrMessageStore->LastWake = WhoIsWaiting;
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

        unsafe public void AddFirstMessageRemoveFileMapping()
        {
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->Action = MessageAction.RemoveFileMapping;
            ControlIdentifier Identifier = new ControlIdentifier();
            AddIdentifierMessage(Identifier);
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
            p = new Parameter(this, Identifier.Description);

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
            ////debug
            //Message* PtrMessage;

            //for (int messageNumber = 1; messageNumber <= 10; messageNumber++)
            //{
            //    PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
            //    Debug.Listeners[0].WriteLine("\t AddMessageFindByHandle Message: " + messageNumber.ToString() + " Parameters: " + PtrMessage->NumberOfParameters.ToString());
            //}
            //Debug.Listeners[0].WriteLine("");
            ////end dbug
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.Refind;

            Parameter ParentHandleParam = new Parameter(this, ParentHandle);
            Parameter ControlHandleParam = new Parameter(this, ControlHandle);

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

        unsafe public void AddFirstMessageGetTitleBarItemRectangle(IntPtr Handle, NM.TitleBarStateElement Item)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetTitleBarItemRectangle;

            Parameter HandleParam = new Parameter(this, Handle);
            Parameter ItemParam = new Parameter(this, (int)Item);
            m_MessageNumber = 0;

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

        unsafe public void DecodeControl(int messageNumber, out ControlIdentifier Identifier)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
            Identifier = new ControlIdentifier();

            // p1  = Parent handle
            if ((PtrMessage->Parameter.TypeCode[0]) == (int)ApeTypeCode.IntPtr)
            {
                Identifier.ParentHandle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = Handle
            if ((PtrMessage->Parameter.TypeCode[1]) == (int)ApeTypeCode.IntPtr)
            {
                Identifier.Handle = (IntPtr)PtrMessage->Parameter.IntPtr[1];
            }
            else
            {
                throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
            }

            if (PtrMessage->NumberOfParameters > 2)
            {
                // p3  = Name
                if (PtrMessage->Parameter.StringLength[2] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[2]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.Name = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[2]), 0, PtrMessage->Parameter.StringLength[2]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[2])).ToString());
                    }
                }

                // p4  = Technology Type (Winforms, WPF, etc)
                if (PtrMessage->Parameter.StringLength[3] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[3]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.TechnologyType = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[3]), 0, PtrMessage->Parameter.StringLength[3]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[3])).ToString());
                    }
                }

                // p5  = Type Namespace
                if (PtrMessage->Parameter.StringLength[4] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[4]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.TypeNameSpace = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[4]), 0, PtrMessage->Parameter.StringLength[4]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[4])).ToString());
                    }
                }

                // p6  = Type Name
                if (PtrMessage->Parameter.StringLength[5] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[5]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.TypeName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[5]), 0, PtrMessage->Parameter.StringLength[5]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[5])).ToString());
                    }
                }

                // p7  = Module Name
                if (PtrMessage->Parameter.StringLength[6] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[6]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.ModuleName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[6]), 0, PtrMessage->Parameter.StringLength[6]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[6])).ToString());
                    }
                }

                // p8  = Assembly Name
                if (PtrMessage->Parameter.StringLength[7] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[7]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.AssemblyName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[7]), 0, PtrMessage->Parameter.StringLength[7]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[7])).ToString());
                    }
                }

                // p9  = Index
                if ((PtrMessage->Parameter.TypeCode[8]) == (Int32)ApeTypeCode.Int32)
                {
                    Identifier.Index = PtrMessage->Parameter.Int32[8];
                }
                else
                {
                    throw new Exception("Expected ApeTypeCode." + ApeTypeCode.Int32.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[8])).ToString());
                }

                // p10 = Text
                if (PtrMessage->Parameter.StringLength[9] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[9]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.Text = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[9]), 0, PtrMessage->Parameter.StringLength[9]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + TypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[9])).ToString());
                    }
                }

                // p11 = ChildOf
                if ((PtrMessage->Parameter.TypeCode[10]) == (int)ApeTypeCode.IntPtr)
                {
                    Identifier.ChildOf = (IntPtr)PtrMessage->Parameter.IntPtr[10];
                }
                else
                {
                    throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[10])).ToString());
                }

                // p12 = SiblingOf
                if ((PtrMessage->Parameter.TypeCode[11]) == (int)ApeTypeCode.IntPtr)
                {
                    Identifier.SiblingOf = (IntPtr)PtrMessage->Parameter.IntPtr[11];
                }
                else
                {
                    throw new Exception("Expected ApeTypeCode.IntPtr got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[11])).ToString());
                }

                // p13  = Description
                if (PtrMessage->Parameter.StringLength[12] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[12]) == (Int32)ApeTypeCode.String)
                    {
                        Identifier.Description = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[12]), 0, PtrMessage->Parameter.StringLength[12]);
                    }
                    else
                    {
                        throw new Exception("Expected ApeTypeCode." + ApeTypeCode.String.ToString() + " got ApeTypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[12])).ToString());
                    }
                }
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
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

        unsafe private void wpfFindFormByHandle(IntPtr Handle, ref string Name, ref string theText, ref Type theType, ref bool foundControl)
        {
            WPF.Application wpfApp = WPF.Application.Current;
            if (wpfApp != null)
            {
                WPF.WindowCollection wpfWindows = (WPF.WindowCollection)wpfApp.GetType().GetProperty("WindowsInternal", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(wpfApp, null);

                foreach (WPF.Window wpfWindow in wpfWindows)
                {
                    wpfWindow.Dispatcher.Invoke(m_GetWPFHandleAndNameAndTitleDelegater, wpfWindow);
                    if (Handle == m_Handle)
                    {
                        Name = m_Name;
                        theText = m_Text;
                        theType = wpfWindow.GetType();
                        foundControl = true;
                        break;
                    }
                }
            }
        }

        unsafe private void wpfFindFormByIdentifier(ControlIdentifier Identifier, ref IntPtr Handle, ref string Name, ref string theText, ref Type theType, ref int CurrentIndex, ref bool foundControl)
        {
            WPF.Application wpfApp = WPF.Application.Current;
            if (wpfApp != null)
            {
                WPF.WindowCollection wpfWindows = (WPF.WindowCollection)wpfApp.GetType().GetProperty("WindowsInternal", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(wpfApp, null);

                foreach (WPF.Window wpfWindow in wpfWindows)
                {
                    wpfWindow.Dispatcher.Invoke(m_GetWPFHandleAndNameAndTitleDelegater, wpfWindow);

                    theType = wpfWindow.GetType();
                    Handle = m_Handle;
                    Name = m_Name;
                    theText = m_Text;

                    if (Identifier.Name != null)
                    {
                        if (Name != Identifier.Name)
                        {
                            continue;
                        }
                    }

                    if (Identifier.TypeNameSpace != null)
                    {
                        if (theType.Namespace != Identifier.TypeNameSpace)
                        {
                            continue;
                        }
                    }

                    if (Identifier.TypeName != null)
                    {
                        if (theType.Name != Identifier.TypeName)
                        {
                            DebugLogging.WriteLog(theType.Name + " != " + Identifier.TypeName);
                            continue;
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

                    CurrentIndex++;

                    DebugLogging.WriteLog("found wpf form for " + Name);

                    if (Identifier.Index > 0)
                    {
                        if (CurrentIndex != Identifier.Index)
                        {
                            continue;
                        }
                    }

                    //we have a match
                    if (NM.IsWindowVisible(Handle))
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
            IntPtr Handle = IntPtr.Zero;
            string Name = null;
            string theText = null;
            bool FoundControl = false;

            if (Identifier.ParentHandle == IntPtr.Zero)     //Find toplevel (parent)
            {
                if (Identifier.Handle != IntPtr.Zero)
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
                            FoundControl = true;
                        }
                    }

                    if (!FoundControl)
                    {
                        if (m_WPF)  //Only look for WPF forms if the application has WPF loaded
                        {
                            if (Identifier.TechnologyType == "Windows NativeWindows Presentation Foundation (WPF)" || Identifier.TechnologyType == null)
                            {
                                Handle = Identifier.Handle;
                                wpfFindFormByHandle(Handle, ref Name, ref theText, ref theType, ref FoundControl);
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
                                    if (Identifier.TechnologyType == "Windows Forms (WinForms)")
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

                                            if (Identifier.TypeNameSpace != null)
                                            {
                                                if (theType.Namespace != Identifier.TypeNameSpace)
                                                {
                                                    continue;
                                                }
                                            }

                                            if (Identifier.TypeName != null)
                                            {
                                                if (theType.Name != Identifier.TypeName)
                                                {
                                                    DebugLogging.WriteLog(theType.Name + " != " + Identifier.TypeName);
                                                    continue;
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

                                            object[] parameters = { form };
                                            theText = (string)form.Invoke(m_GetTextDelegater, parameters);
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

                                            CurrentIndex++;

                                            DebugLogging.WriteLog("found form for " + Name);

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

                                    if (Identifier.TechnologyType == "Windows Native")
                                    {
                                        //Windows Native
                                        Handle = hWnd;

                                        if (Identifier.TypeName != null)
                                        {
                                            string theTypeName = NM.GetClassName(Handle);
                                            if (theTypeName != Identifier.TypeName)
                                            {
                                                DebugLogging.WriteLog(theTypeName + " != " + Identifier.TypeName);
                                                continue;
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

                                        CurrentIndex++;

                                        DebugLogging.WriteLog("found form for " + Name);

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
                                if (m_WPF)  //Only look for WPF forms if the application has WPF loaded
                                {
                                    if (Identifier.TechnologyType == "Windows NativeWindows Presentation Foundation (WPF)")
                                    {
                                        wpfFindFormByIdentifier(Identifier, ref Handle, ref Name, ref theText, ref theType, ref CurrentIndex, ref FoundControl);
                                    }
                                }
                            }
                        }
                        catch (InvalidOperationException ex)
                        {
                            if (ex.Message == "Collection was modified; enumeration operation may not execute.")
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
                if (Identifier.Handle != IntPtr.Zero)
                {
                    //WinForms
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
                            theType = TheControl.GetType();
                            FoundControl = true;
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
                        NM.EnumChildWindows(Identifier.ParentHandle, EnumThreadProcedue, IntPtr.Zero);

                        foreach (IntPtr hWnd in m_AllControls)
                        {
                            if (NM.IsWindowVisible(hWnd))
                            {
                                if (Identifier.TechnologyType == "Windows Forms (WinForms)")
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
                                            if (theType.Namespace != Identifier.TypeNameSpace)
                                            {
                                                continue;
                                            }
                                        }

                                        if (Identifier.TypeName != null)
                                        {
                                            if (theType.Name != Identifier.TypeName)
                                            {
                                                DebugLogging.WriteLog(theType.Name + " != " + Identifier.TypeName);
                                                continue;
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

                                        object[] parameters = { control };
                                        theText = (string)control.Invoke(m_GetTextDelegater, parameters);
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

                                if (Identifier.TechnologyType == "Windows Native")
                                {
                                    //Windows Native
                                    Handle = hWnd;

                                    if (Identifier.TypeName != null)
                                    {
                                        string theTypeName = NM.GetClassName(Handle);
                                        if (theTypeName != Identifier.TypeName)
                                        {
                                            DebugLogging.WriteLog(theTypeName + " != " + Identifier.TypeName);
                                            continue;
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

                                    CurrentIndex++;

                                    DebugLogging.WriteLog("found native child control " + Name);

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
                            if (Identifier.TechnologyType == "Windows NativeWindows Presentation Foundation (WPF)")
                            {
                            }
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
                NewIdentifier.TechnologyType = GetTechnologyType(theType);
                if (NewIdentifier.TechnologyType == "Windows Native")
                {
                    NewIdentifier.TypeName = NM.GetClassName(Identifier.Handle);
                    NewIdentifier.ModuleName = Path.GetFileName(NM.GetWindowModuleFileName(Identifier.Handle));
                }
                else
                {
                    NewIdentifier.TypeNameSpace = theType.Namespace;
                    NewIdentifier.TypeName = theType.Name;
                    NewIdentifier.ModuleName = theType.Module.Name;
                    NewIdentifier.AssemblyName = theType.Assembly.GetName().Name;
                }
                NewIdentifier.Index = Identifier.Index;
                NewIdentifier.Text = theText;
                AddIdentifierMessage(NewIdentifier);
                return null;
            }
            else
            {
                return "Failed to find the " + Identifier.Description;
            }
        }

        unsafe private void Refind(int messageNumber)
        {
            object DestinationObject = null;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));

            ControlIdentifier Identifier;
            DecodeControl(messageNumber, out Identifier);

            //WinForms
            if (Identifier.TechnologyType == "Windows Forms (WinForms)" || Identifier.TechnologyType == null)
            {
                DestinationObject = WF.Control.FromHandle(Identifier.Handle);
            }

            if (DestinationObject == null)
            {
                if (Identifier.TechnologyType == "Windows Presentation Foundation (WPF)" || Identifier.TechnologyType == null)
                {
                    //WPF TODO
                }
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
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * Parameter.OneLargerThanApeTypeCodeEnumMax) + (int)ApeTypeCode.Empty);
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
                case "LzGenericWalker":
                    switch (TypeName)
                    {
                        case "GenericWalker":
                            return "GUIGenericWalker";
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

        unsafe private void Reflect(int MessageNumber)
        {
            object SourceObject;
            object DestinationObject;
            IntPtr datastoreTypeHandle = IntPtr.Zero;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

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
                            ParametersType[i] = m_TypeBoolean;
                            ParametersObject[i] = PtrMessage->Parameter.Boolean[i];
                            break;
                        case (int)ApeTypeCode.Char:
                            ParametersType[i] = m_TypeChar;
                            ParametersObject[i] = PtrMessage->Parameter.Char[i];
                            break;
                        case (int)ApeTypeCode.SByte:
                            ParametersType[i] = m_TypeSByte;
                            ParametersObject[i] = PtrMessage->Parameter.SByte[i];
                            break;
                        case (int)ApeTypeCode.Byte:
                            ParametersType[i] = m_TypeByte;
                            ParametersObject[i] = PtrMessage->Parameter.Byte[i];
                            break;
                        case (int)ApeTypeCode.Int16:
                            ParametersType[i] = m_TypeInt16;
                            ParametersObject[i] = PtrMessage->Parameter.Int16[i];
                            break;
                        case (int)ApeTypeCode.UInt16:
                            ParametersType[i] = m_TypeUInt16;
                            ParametersObject[i] = PtrMessage->Parameter.UInt16[i];
                            break;
                        case (int)ApeTypeCode.Int32:
                            ParametersType[i] = m_TypeInt32;
                            ParametersObject[i] = PtrMessage->Parameter.Int32[i];
                            break;
                        case (int)ApeTypeCode.UInt32:
                            ParametersType[i] = m_TypeUInt32;
                            ParametersObject[i] = PtrMessage->Parameter.UInt32[i];
                            break;
                        case (int)ApeTypeCode.Int64:
                            ParametersType[i] = m_TypeInt64;
                            ParametersObject[i] = PtrMessage->Parameter.Int64[i];
                            break;
                        case (int)ApeTypeCode.UInt64:
                            ParametersType[i] = m_TypeUInt64;
                            ParametersObject[i] = PtrMessage->Parameter.UInt64[i];
                            break;
                        case (int)ApeTypeCode.Single:
                            ParametersType[i] = m_TypeSingle;
                            ParametersObject[i] = PtrMessage->Parameter.Single[i];
                            break;
                        case (int)ApeTypeCode.Double:
                            ParametersType[i] = m_TypeDouble;
                            ParametersObject[i] = PtrMessage->Parameter.Double[i];
                            break;
                        case (int)ApeTypeCode.Decimal:
                            ParametersType[i] = m_TypeDecimal;
                            int[] DecimalBits = new int[4];
                            DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[i];
                            DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[i];
                            DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[i];
                            DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[i];
                            ParametersObject[i] = new decimal(DecimalBits);
                            break;
                        case (int)ApeTypeCode.DateTime:
                            ParametersType[i] = m_TypeDateTime;
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
                            ParametersType[i] = m_TypeIntPtr;
                            ParametersObject[i] = new IntPtr(PtrMessage->Parameter.IntPtr[i]);
                            break;
                        case (int)ApeTypeCode.DataStore:
                            if (datastoreTypeHandle != IntPtr.Zero)
                            {
                                throw new Exception("Only passing one datastore item as a parameter is supported");
                            }

                            int datastoreNumber = PtrMessage->Parameter.Int32[i];
                            switch (datastoreNumber)
                            {
                                case 0:
                                    ParametersType[i] = tempStore0.GetType();
                                    ParametersObject[i] = tempStore0;
                                    break;
                                case 1:
                                    ParametersType[i] = tempStore1.GetType();
                                    ParametersObject[i] = tempStore1;
                                    break;
                                case 2:
                                    ParametersType[i] = tempStore2.GetType();
                                    ParametersObject[i] = tempStore2;
                                    break;
                                case 3:
                                    ParametersType[i] = tempStore3.GetType();
                                    ParametersObject[i] = tempStore3;
                                    break;
                                case 4:
                                    ParametersType[i] = tempStore4.GetType();
                                    ParametersObject[i] = tempStore4;
                                    break;
                                case 5:
                                    ParametersType[i] = tempStore5.GetType();
                                    ParametersObject[i] = tempStore5;
                                    break;
                                case 6:
                                    ParametersType[i] = tempStore6.GetType();
                                    ParametersObject[i] = tempStore6;
                                    break;
                                case 7:
                                    ParametersType[i] = tempStore7.GetType();
                                    ParametersObject[i] = tempStore7;
                                    break;
                                case 8:
                                    ParametersType[i] = tempStore8.GetType();
                                    ParametersObject[i] = tempStore8;
                                    break;
                                case 9:
                                    ParametersType[i] = tempStore9.GetType();
                                    ParametersObject[i] = tempStore9;
                                    break;
                                default:
                                    throw new Exception("Unsupported SourceStore " + datastoreNumber.ToString());
                            }
                            datastoreTypeHandle = ParametersType[i].TypeHandle.Value;
                            break;
                        default:
                            throw new Exception("Unsupported ApeTypeCode: " + PtrMessage->Parameter.TypeCode[i].ToString());
                    }
                }
                if (datastoreTypeHandle == IntPtr.Zero)  //none of the parameters are a datastore type
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

            Type SourceType;
            string Name;
            Fasterflect.ConstructorInvoker ConstructorInvoker;
            Fasterflect.MethodInvoker MethodInvoker;
            Fasterflect.MemberGetter MemberGetter;
            Fasterflect.ArrayElementGetter ArrayElementGetter;

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

                if (Marshal.IsComObject(SourceObject))
                {
                    DestinationObject = SourceObject.GetType().InvokeMember(Name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, SourceObject, ParametersObject);
                }
                else
                {
                    //Get the value
                    if (Name == "<Array>")
                    {
                        ArrayElementGetter = SourceType.DelegateForGetElement();
                        DestinationObject = ArrayElementGetter(SourceObject, (int)ParametersObject[0]);
                        ArrayElementGetter = null;
                    }
                    else
                    {
                        switch (PtrMessage->MemberType)
                        {
                            case MemberTypes.Constructor:
                                ConstructorInvokerCache.GetFromList(Name, PtrMessage->TypeCodeKey, datastoreTypeHandle, out ConstructorInvoker);
                                if (ConstructorInvoker == null)
                                {
                                    Type typeContainingConstructor;
                                    typeContainingConstructor = SourceType.Assembly.GetTypes().FirstOrDefault(x => x.FullName == Name);
                                    if (typeContainingConstructor == null)
                                    {
                                        typeContainingConstructor = SourceType.Assembly.GetReferencedAssemblies().Select(x => Assembly.Load(x)).SelectMany(x => x.GetTypes()).FirstOrDefault(x => x.FullName == Name);
                                    }
                                    ConstructorInvoker = typeContainingConstructor.DelegateForCreateInstance(ParametersType);
                                    ConstructorInvokerCache.AddToList(Name, PtrMessage->TypeCodeKey, datastoreTypeHandle, ConstructorInvoker);
                                }
                                DestinationObject = ConstructorInvoker.Invoke(ParametersObject);
                                //DestinationObject = ((Control)tempStore0).Invoke(ConstructorInvoker, SourceObject.WrapIfValueType(), ParametersObject);
                                ConstructorInvoker = null;
                                break;
                            case MemberTypes.Field:
                                MemberGetterCache.GetFromList(SourceType.TypeHandle.Value, Name, out MemberGetter);
                                if (MemberGetter == null)
                                {
                                    MemberGetter = SourceType.DelegateForGetFieldValue(Name);
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
                                        MemberGetter = SourceType.DelegateForGetPropertyValue(Name);
                                        MemberGetterCache.AddToList(SourceType.TypeHandle.Value, Name, MemberGetter);
                                    }
                                    DestinationObject = ((WF.Control)tempStore0).Invoke((Delegate)MemberGetter, SourceObject.WrapIfValueType());
                                    MemberGetter = null;
                                }
                                else
                                {
                                    MethodInvokerCache.GetFromList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypeHandle, out MethodInvoker);
                                    if (MethodInvoker == null)
                                    {
                                        MethodInvoker = SourceType.DelegateForGetIndexer(ParametersType);
                                        MethodInvokerCache.AddToList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypeHandle, MethodInvoker);
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
                                else
                                {
                                    MethodInvokerCache.GetFromList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypeHandle, out MethodInvoker);
                                    if (MethodInvoker == null)
                                    {
                                        MethodInvoker = SourceType.DelegateForCallMethod(Name, ParametersType);
                                        MethodInvokerCache.AddToList(SourceType.TypeHandle.Value, Name, PtrMessage->TypeCodeKey, datastoreTypeHandle, MethodInvoker);
                                    }
                                    DestinationObject = ((WF.Control)tempStore0).Invoke(MethodInvoker, SourceObject.WrapIfValueType(), ParametersObject);
                                }
                                MethodInvoker = null;
                                break;
                            default:
                                throw new Exception("Unsupported member type: " + (PtrMessage->MemberType).ToString());
                        }
                    }
                }
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
    }
}

