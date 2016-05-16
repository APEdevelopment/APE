using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//
using System.IO;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Fasterflect;
using System.Threading;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.ComponentModel;
using System.Management;
using APE.Domain;
using NM = APE.Native.NativeMethods;
using Microsoft.Win32;
using System.Runtime;

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
    }

    public class APEIPC
    {
        private const int MaxParameters = 10;
        private const int MaxMessages = 20;
        private const int ERROR_ALREADY_EXISTS = 183;
        private const int OneLargerThanTypeCodeEnumMax = 19;
        private const int StringSpaceBytes = 1048576;

        //Memory Map File for IPC for message strings
        private IntPtr m_HandleMemoryMappedFileStringStore;
        private IntPtr m_IntPtrMemoryMappedFileViewStringStore;
        private int m_StringStoreOffset;

        private NM.HookProc MouseHookProcedure;
        private NM.EnumThread EnumThreadProcedue;
        private int m_hMouseHook = 0;
        private IntPtr m_HookWindow;
        private IntPtr ContextMenuStrip;

        private bool m_WM_LBUTTONDOWN = false;
        private bool m_WM_LBUTTONUP = false;
        private bool m_WM_LBUTTONDBLCLK = false;
        private bool m_WM_RBUTTONDOWN = false;
        private bool m_WM_RBUTTONUP = false;
        private bool m_WM_RBUTTONDBLCLK = false;
        private bool m_WM_MBUTTONDOWN = false;
        private bool m_WM_MBUTTONUP = false;
        private bool m_WM_MBUTTONDBLCLK = false;

        //Memory Map File for IPC for message value types
        private IntPtr m_HandleMemoryMappedFileMessageStore;
        private IntPtr m_IntPtrMemoryMappedFileViewMessageStore;
        private unsafe MessageStore* m_PtrMessageStore;

        private unsafe int m_SizeOfMessage = sizeof(Message);
        private ParametersTypeCircularList ParametersTypeCache = new ParametersTypeCircularList(3);    //3 is optimal here as there are only a couple of very frequently used types
        private MemberGetterCircularList MemberGetterCache = new MemberGetterCircularList(10);
        private MethodInvokerCircularList MethodInvokerCache = new MethodInvokerCircularList(10);

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
        Control[] m_AllControlsOnForm;
        EventSet Side;
        Process ApeProcess = null;
        Process AUTProcess = null;
        bool m_Abort = false;
        uint m_TimeOut = 30000;         //timeout default 30 seconds
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
        public enum MessageAction : int
        {
            RemoveFileMapping = -1,
            None = 0,
            Find = 1,
            Refind = 2,
            ReflectGet = 3,
            GetResult = 4,
            GetListViewGroupRectangle = 5,
            GetListViewItemRectangle = 6,
            SetTimeOuts = 7,
            GetTitleBarItemRectangle = 8,
            AddMouseHook = 9,
            RemoveMouseHook = 10,
            WaitForMouseState = 11,
            GarbageCollect = 12,
            GetContextMenuStrip = 13,
            GetAppDomains = 14,
            GetRecognisedType = 15,
            GetApeTypeFromType = 16,
            GetApeTypeFromObject = 17,
            ReflectPoll = 18,
        }

        [Flags]
        public enum MessageResult : int
        {
            Success = 1,
            Failure = 2,
        }

        [Flags]
        public enum MouseKeyModifier : int
        {
            Control = 2,
            Shift = 4,
        }

        public enum MouseButton : int
        {
            Left = 0,
            Right = 1,
            Middle = 2,
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

        public enum EventSet : byte
        {
            APE = 0,
            AUT = 1,
        }

        // if you add more Messages here remember to update MaxMessages as well
        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        struct MessageStore
        {
            public Message Message0;
            public Message Message1;
            public Message Message2;
            public Message Message3;
            public Message Message4;
            public Message Message5;
            public Message Message6;
            public Message Message7;
            public Message Message8;
            public Message Message9;
            public Message Message10;
            public Message Message11;
            public Message Message12;
            public Message Message13;
            public Message Message14;
            public Message Message15;
            public Message Message16;
            public Message Message17;
            public Message Message18;
            public Message Message19;
            public byte NumberOfMessages;
            public EventSet LastWake;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        struct Message
        {
            public int MoreStringData;
            public int TotalStringDataLength;
            public MessageAction Action;
            public DataStores SourceStore;
            public DataStores DestinationStore;
            public int NameOffset;
            public int NameLength;
            public int NumberOfParameters;
            public MemberTypes MemberType;
            public long TypeCodeKey;
            public Parameters Parameter;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        unsafe struct Parameters
        {
            public fixed Int32 TypeCode[MaxParameters];
            public fixed Boolean Boolean[MaxParameters];
            public fixed Byte Byte[MaxParameters];
            public fixed Char Char[MaxParameters];
            public fixed Int64 DateTimeBinary[MaxParameters];
            public fixed Int32 DecimalBits0[MaxParameters];
            public fixed Int32 DecimalBits1[MaxParameters];
            public fixed Int32 DecimalBits2[MaxParameters];
            public fixed Int32 DecimalBits3[MaxParameters];
            public fixed Double Double[MaxParameters];
            public fixed Int16 Int16[MaxParameters];
            public fixed Int32 Int32[MaxParameters];
            public fixed Int64 Int64[MaxParameters];
            public fixed SByte SByte[MaxParameters];
            public fixed Single Single[MaxParameters];
            public fixed Int32 StringOffset[MaxParameters];
            public fixed Int32 StringLength[MaxParameters];
            public fixed UInt16 UInt16[MaxParameters];
            public fixed UInt32 UInt32[MaxParameters];
            public fixed UInt64 UInt64[MaxParameters];
            public fixed Int64 IntPtr[MaxParameters];
        }

        private bool EnumThreadCallback(IntPtr hWnd, IntPtr lParam)
        {
            Control FoundControl = Form.FromHandle(hWnd);
            if (FoundControl != null)
            {
                if (FoundControl.Visible)
                {
                    if (FoundControl.GetType().Name == "ContextMenuStrip")
                    {
                        ContextMenuStrip = hWnd;
                        return false;
                    }
                }
            }
            return true;
        }

        public int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return NM.CallNextHookEx(m_hMouseHook, nCode, wParam, lParam);
            }
            else
            {
                NM.MouseHookStruct MyMouseHookStruct = (NM.MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(NM.MouseHookStruct));

                if (MyMouseHookStruct.hwnd == m_HookWindow)
                {
                    if (nCode == NM.HC_ACTION)
                    {
                        switch (wParam.ToInt32())
                        {
                            case NM.WM_LBUTTONDOWN:
                                WriteLog("Left Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDOWN = true;
                                break;
                            case NM.WM_NCLBUTTONDOWN:
                                WriteLog("NCLeft Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDOWN = true;
                                break;
                            case NM.WM_LBUTTONUP:
                                WriteLog("Left Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONUP = true;
                                break;
                            case NM.WM_NCLBUTTONUP:
                                WriteLog("NCLeft Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONUP = true;
                                break;
                            case NM.WM_LBUTTONDBLCLK:
                                WriteLog("Left Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDBLCLK = true;
                                break;
                            case NM.WM_NCLBUTTONDBLCLK:
                                WriteLog("NCLeft Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDBLCLK = true;
                                break;
                            case NM.WM_RBUTTONDOWN:
                                WriteLog("Right Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDOWN = true;
                                break;
                            case NM.WM_NCRBUTTONDOWN:
                                WriteLog("NCRight Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDOWN = true;
                                break;
                            case NM.WM_RBUTTONUP:
                                WriteLog("Right Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONUP = true;
                                break;
                            case NM.WM_NCRBUTTONUP:
                                WriteLog("NCRight Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONUP = true;
                                break;
                            case NM.WM_RBUTTONDBLCLK:
                                WriteLog("Right Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDBLCLK = true;
                                break;
                            case NM.WM_NCRBUTTONDBLCLK:
                                WriteLog("NCRight Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDBLCLK = true;
                                break;
                            case NM.WM_MBUTTONDOWN:
                                WriteLog("Middle Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDOWN = true;
                                break;
                            case NM.WM_NCMBUTTONDOWN:
                                WriteLog("NCMiddle Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDOWN = true;
                                break;
                            case NM.WM_MBUTTONUP:
                                WriteLog("Middle Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                break;
                            case NM.WM_NCMBUTTONUP:
                                WriteLog("NCMiddle Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONUP = true;
                                break;
                            case NM.WM_MBUTTONDBLCLK:
                                WriteLog("Middle Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDBLCLK = true;
                                break;
                            case NM.WM_NCMBUTTONDBLCLK:
                                WriteLog("NCMiddle Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDBLCLK = true;
                                break;
                        }
                    }
                }
                return NM.CallNextHookEx(m_hMouseHook, nCode, wParam, lParam);
            }
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
            string Path = Directory.GetCurrentDirectory() + @"\APE";
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
                        string Assembly = @"APE\APE.Loader_x86.dll";
                        Injector.StartInfo = new ProcessStartInfo(@"APE\APE.Injector_x86.exe", AUTProcessId + " " + Assembly + " " + Method + " " + APEProcessId);
                    }
                    else
                    {
                        string Assembly = @"APE\APE.Loader_x64.dll";
                        Injector.StartInfo = new ProcessStartInfo(@"APE\APE.Injector_x64.exe", AUTProcessId + " " + Assembly + " " + Method + " " + APEProcessId);
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
        public unsafe APEIPC(int APEPID, string AppDomainToLoadInto)
        {
            try
            {
                Thread myThread = new Thread(() => DecodeMessage(APEPID, AppDomainToLoadInto));
                myThread.SetApartmentState(ApartmentState.STA);
                myThread.IsBackground = true;
                myThread.Start();
            }
            catch (Exception ex)
            {
                TextWriter log = File.AppendText(@"C:\critical.log");
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
                        this.AddMessageRemoveFileMapping();
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
                throw new Exception("GetListViewItemRectangle must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = TimeOut
            if ((PtrMessage->Parameter.TypeCode[0]) == (Int32)TypeCode.UInt32)
            {
                m_TimeOut = PtrMessage->Parameter.UInt32[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.UInt32 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        private unsafe void AddMouseHook(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("AddMouseHook must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = handle
            IntPtr Handle;
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            int ThreadID = NM.GetWindowThreadProcessId(Handle, IntPtr.Zero);
            m_HookWindow = Handle;

            // Add the mouse hook
            WriteLog("Adding Mouse hook");
            m_hMouseHook = NM.SetWindowsHookEx(NM.WH_MOUSE, MouseHookProcedure, IntPtr.Zero, ThreadID);
            if (m_hMouseHook == 0)
            {
                throw new Exception("SetWindowsHookEx Failed");
            }
            WriteLog("Added Mouse hook");

            ClearMouseState();
        }

        private unsafe void RemoveMouseHook(int MessageNumber)
        {
            bool Return = false;

            //must be first message
            if (MessageNumber == 1)
            {
                throw new Exception("RemoveMouseHook must not be the first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = handle
            IntPtr Handle;
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            // remove the hook
            WriteLog("Removing Mouse hook");
            Return = NM.UnhookWindowsHookEx(m_hMouseHook);
            if (!Return)
            {
                throw new Exception("UnhookWindowsHookEx Failed for Mouse hook");
            }
            m_hMouseHook = 0;
            WriteLog("Removed Mouse hook");

            ClearMouseState();
        }

        private void ClearMouseState()
        {
            m_WM_LBUTTONDBLCLK = false;
            m_WM_LBUTTONDOWN = false;
            m_WM_LBUTTONUP = false;
            m_WM_RBUTTONDBLCLK = false;
            m_WM_RBUTTONDOWN = false;
            m_WM_RBUTTONUP = false;
            m_WM_MBUTTONDBLCLK = false;
            m_WM_MBUTTONDOWN = false;
            m_WM_MBUTTONUP = false;
        }

        private unsafe void GarbageCollect(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("GarbageCollect must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1 = Generation
            int generation;
            if ((PtrMessage->Parameter.TypeCode[0]) == (Int32)TypeCode.Int32)
            {
                generation = PtrMessage->Parameter.Int32[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Int32 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            //// Make sure the the logs are all flushed
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

        private unsafe void WaitForMouseState(int MessageNumber)
        {
            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("WaitForMouseState must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1 = MouseButton
            MouseButton Button;
            if ((PtrMessage->Parameter.TypeCode[0]) == (Int32)TypeCode.Int32)
            {
                Button = (MouseButton)PtrMessage->Parameter.Int32[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Int32 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2 = MouseDown
            bool MouseDown;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)TypeCode.Boolean)
            {
                MouseDown = PtrMessage->Parameter.Boolean[1];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Boolean got System.TypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
            }

            // p3 = FirstClick
            bool FirstClick;
            if ((PtrMessage->Parameter.TypeCode[2]) == (Int32)TypeCode.Boolean)
            {
                FirstClick = PtrMessage->Parameter.Boolean[2];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Boolean got System.TypeCode." + (PtrMessage->Parameter.TypeCode[2]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            Stopwatch timer = Stopwatch.StartNew();

            if (MouseDown)
            {
                WriteLog("Waiting on " + Button.ToString() + " mouse down");
            }
            else
            {
                WriteLog("Waiting on " + Button.ToString() + " mouse up");
            }

            bool done = false;
            while (!done)
            {
                switch (Button)
                {
                    case MouseButton.Left:
                        if (MouseDown)
                        {
                            if (FirstClick)
                            {
                                if (m_WM_LBUTTONDOWN)
                                {
                                    timer.Stop();
                                    return;
                                }
                            }
                            else
                            {
                                if (m_WM_LBUTTONDOWN || m_WM_LBUTTONDBLCLK)
                                {
                                    timer.Stop();
                                    return;
                                }
                            }
                        }
                        else
                        {
                            if (m_WM_LBUTTONUP)
                            {
                                timer.Stop();
                                return;
                            }
                        }
                        break;
                    case MouseButton.Right:
                        if (MouseDown)
                        {
                            if (FirstClick)
                            {
                                if (m_WM_RBUTTONDOWN)
                                {
                                    timer.Stop();
                                    return;
                                }
                            }
                            else
                            {
                                if (m_WM_RBUTTONDOWN || m_WM_RBUTTONDBLCLK)
                                {
                                    timer.Stop();
                                    return;
                                }
                            }
                        }
                        else
                        {
                            if (m_WM_RBUTTONUP)
                            {
                                timer.Stop();
                                return;
                            }
                        }
                        break;
                    case MouseButton.Middle:
                        if (MouseDown)
                        {
                            if (FirstClick)
                            {
                                if (m_WM_MBUTTONDOWN)
                                {
                                    timer.Stop();
                                    return;
                                }
                            }
                            else
                            {
                                if (m_WM_MBUTTONDOWN || m_WM_MBUTTONDBLCLK)
                                {
                                    timer.Stop();
                                    return;
                                }
                            }
                        }
                        else
                        {
                            if (m_WM_MBUTTONUP)
                            {
                                timer.Stop();
                                return;
                            }
                        }
                        break;
                }

                if (!done)
                {
                    Thread.Yield();
                }

                if (timer.ElapsedMilliseconds > m_TimeOut)
                {
                    timer.Stop();
                    if (MouseDown)
                    {
                        throw new Exception("Failed to find " + Button.ToString() + " mouse down");
                    }
                    else
                    {
                        throw new Exception("Failed to find " + Button.ToString() + " mouse up");
                    }

                }
            }

            ClearMouseState();

            WriteLog("Mouse State done");
        }

        public unsafe void GetContextMenuStrip(int MessageNumber)
        {

            //must be first message
            if (MessageNumber != 1)
            {
                throw new Exception("GetListViewItemRectangle must be first message");
            }

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

            // p1  = handle
            IntPtr Handle;
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            //Find it
            ContextMenuStrip = IntPtr.Zero;
            NM.EnumThreadWindows((uint)NM.GetWindowThreadProcessId(Handle, IntPtr.Zero), EnumThreadProcedue, IntPtr.Zero);

            AddReturnValue(new Parameter(this, ContextMenuStrip));
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
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = TitleBarStateElement
            NM.TitleBarStateElement Item;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)TypeCode.Int32)
            {
                Item = (NM.TitleBarStateElement)PtrMessage->Parameter.Int32[1];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Int32 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
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
                    Top = tbi.rcTitleBar.Top;
                    Left = tbi.rcTitleBar.Left;
                    Bottom = tbi.rcTitleBar.Bottom;
                    Right = tbi.rcTitleBar.Right;
                    break;
                case NM.TitleBarStateElement.Minimize:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Minimize];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].Top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].Left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].Bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].Right;
                    break;
                case NM.TitleBarStateElement.Maximize:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Maximize];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].Top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].Left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].Bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].Right;
                    break;
                case NM.TitleBarStateElement.Help:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Help];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Help].Top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Help].Left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Help].Bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Help].Right;
                    break;
                case NM.TitleBarStateElement.Close:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Close];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Close].Top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Close].Left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Close].Bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Close].Right;
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
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = Index
            int Index;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)TypeCode.Int32)
            {
                Index = PtrMessage->Parameter.Int32[1];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Int32 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
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
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)
            {
                Handle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = Group
            int GroupID;
            if ((PtrMessage->Parameter.TypeCode[1]) == (Int32)TypeCode.Int32)
            {
                GroupID = PtrMessage->Parameter.Int32[1];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.Int32 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
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

        [MethodImplAttribute(MethodImplOptions.AggressiveInlining)]
        unsafe public dynamic GetValueFromMessage()
        {
            m_MessageNumber++;
            return GetValueFromMessage(m_MessageNumber);
        }

        [MethodImplAttribute(MethodImplOptions.AggressiveInlining)]
        unsafe public dynamic GetValueFromMessage(int messageNumber)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
            int parameter = 0;
            dynamic returnValue;

            switch ((int)PtrMessage->Parameter.TypeCode[parameter])
            {
                case (int)TypeCode.Boolean:
                    returnValue = PtrMessage->Parameter.Boolean[parameter];
                    break;
                case (int)TypeCode.Char:
                    returnValue = PtrMessage->Parameter.Char[parameter];
                    break;
                case (int)TypeCode.SByte:
                    returnValue = PtrMessage->Parameter.SByte[parameter];
                    break;
                case (int)TypeCode.Byte:
                    returnValue = PtrMessage->Parameter.Byte[parameter];
                    break;
                case (int)TypeCode.Int16:
                    returnValue = PtrMessage->Parameter.Int16[parameter];
                    break;
                case (int)TypeCode.UInt16:
                    returnValue = PtrMessage->Parameter.UInt16[parameter];
                    break;
                case (int)TypeCode.Int32:
                    returnValue = PtrMessage->Parameter.Int32[parameter];
                    break;
                case (int)TypeCode.UInt32:
                    returnValue = PtrMessage->Parameter.UInt32[parameter];
                    break;
                case (int)TypeCode.Int64:
                    returnValue = PtrMessage->Parameter.Int64[parameter];
                    break;
                case (int)TypeCode.UInt64:
                    returnValue = PtrMessage->Parameter.UInt64[parameter];
                    break;
                case (int)TypeCode.Single:
                    returnValue = PtrMessage->Parameter.Single[parameter];
                    break;
                case (int)TypeCode.Double:
                    returnValue = PtrMessage->Parameter.Double[parameter];
                    break;
                case (int)TypeCode.Decimal:
                    int[] DecimalBits = new int[4];
                    DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[parameter];
                    DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[parameter];
                    DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[parameter];
                    DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[parameter];
                    returnValue = new decimal(DecimalBits);
                    break;
                case (int)TypeCode.DateTime:
                    returnValue = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[parameter]);
                    break;
                case (int)TypeCode.String:
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
                case 17:
                    returnValue = new IntPtr(PtrMessage->Parameter.IntPtr[parameter]);
                    break;
                case (int)TypeCode.Empty:
                    returnValue = null;
                    break;
                default:
                    throw new Exception("Unsupported System.TypeCode: " + ((int)PtrMessage->Parameter.TypeCode).ToString());
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            return returnValue;
        }

        //TODO is worth doing this 
        [MethodImplAttribute(MethodImplOptions.AggressiveInlining)]
        unsafe public void GetValueFromMessage(int messageNumber, out string value)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
            int parameter = 0;

            if (PtrMessage->Parameter.StringLength[parameter] == -1)
            {
                string Empty = null;
                value = Empty;
            }
            else
            {
                value = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[parameter]), 0, PtrMessage->Parameter.StringLength[parameter]);
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;
        }

        [MethodImplAttribute(MethodImplOptions.AggressiveInlining)]
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

        unsafe public void AddMessageGetValue(DataStores SourceStore)
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

        unsafe public void AddMessageGetValue(DataStores SourceStore, int TypeCode)
        {

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

            PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = TypeCode;
            PtrMessage->NumberOfParameters++;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneGet = true;
        }

        unsafe public void AddMessageFindByProperty(ControlIdentifier Identifier)
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

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

        unsafe private void AddMessageSetTimeOuts()
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
            PtrMessage->Action = MessageAction.SetTimeOuts;

            //p1
            Parameter p = new Parameter(this, TimeOut);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageRemoveFileMapping()
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

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
            WriteLog("setting result message " + m_PtrMessageStore->NumberOfMessages.ToString() + " to " + Result.ToString());

            // This should be the last message returned
            Parameter p;

            //p1
            p = new Parameter(this, (int)Result);

            m_PtrMessageStore->NumberOfMessages++;
        }

        unsafe public void AddResultMessage(MessageResult Result, string Message)
        {
            WriteLog("setting result message " + m_PtrMessageStore->NumberOfMessages.ToString() + " to " + Result.ToString());

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

            m_PtrMessageStore->NumberOfMessages++;
        }

        unsafe public void AddMessageFindByHandle(DataStores DestinationStore, IntPtr ParentHandle, IntPtr ControlHandle)
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

            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->DestinationStore = DestinationStore;
            PtrMessage->Action = MessageAction.Refind;

            Parameter ParentHandleParam = new Parameter(this, ParentHandle);
            Parameter ControlHandleParam = new Parameter(this, ControlHandle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
        }

        unsafe public void AddMessageGetAppDomains()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetAppDomains;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageGetTitleBarItemRectangle(IntPtr Handle, NM.TitleBarStateElement Item)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetTitleBarItemRectangle;

            Parameter HandleParam = new Parameter(this, Handle);
            Parameter ItemParam = new Parameter(this, (int)Item);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageGetListViewItemRectangle(IntPtr ControlHandle, int ItemIndex)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetListViewItemRectangle;

            Parameter ControlHandleParam = new Parameter(this, ControlHandle);
            Parameter GroupIDParam = new Parameter(this, ItemIndex);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageGetListViewGroupRectangle(IntPtr ControlHandle, int GroupID)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetListViewGroupRectangle;

            Parameter ControlHandleParam = new Parameter(this, ControlHandle);
            Parameter GroupIDParam = new Parameter(this, GroupID);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        //unsafe public void AddMessageMouseClick(IntPtr Handle, MouseButton Button, Boolean Down, Boolean Up, int Clicks, bool ControlKey, bool ShiftKey)
        //{
        //    m_StringStoreOffset = 0;
        //    m_PtrMessageStore->NumberOfMessages = 0;
        //    m_DoneFind = false;
        //    m_DoneQuery = false;
        //    m_DoneGet = false;

        //    Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

        //    PtrMessage->Action = MessageAction.MouseClick;

        //    param HandleParam = new param(this, Handle);
        //    param ButtonParam = new param(this, (int)Button);
        //    param DownParam = new param(this, Down);
        //    param UpParam = new param(this, Up);
        //    param ClicksParam = new param(this, Clicks);
        //    param ControlKeyParam = new param(this, ControlKey);
        //    param ShiftKeyParam = new param(this, ShiftKey);

        //    m_PtrMessageStore->NumberOfMessages++;
        //    m_DoneFind = true;
        //    m_DoneQuery = true;
        //    m_DoneGet = true;
        //}

        unsafe public void AddMessageAddMouseHook(IntPtr Handle)
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.AddMouseHook;

            Parameter HandleParam = new Parameter(this, Handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageRemoveMouseHook(IntPtr Handle)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.RemoveMouseHook;

            Parameter HandleParam = new Parameter(this, Handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageGetContextMenuStrip(IntPtr Handle)
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GetContextMenuStrip;

            Parameter HandleParam = new Parameter(this, Handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageGarbageCollect(int generation)
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.GarbageCollect;

            Parameter generationParam = new Parameter(this, generation);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessageWaitForMouseState(MouseButton Button, Boolean MouseDown, Boolean FirstClick)
        {
            m_StringStoreOffset = 0;
            m_PtrMessageStore->NumberOfMessages = 0;
            m_DoneFind = false;
            m_DoneQuery = false;
            m_DoneGet = false;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));

            PtrMessage->Action = MessageAction.WaitForMouseState;

            Parameter ButtonParam = new Parameter(this, (int)Button);
            Parameter MouseDownParam = new Parameter(this, MouseDown);
            Parameter FirstClickParam = new Parameter(this, FirstClick);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        unsafe public void AddMessagePollMember(DataStores SourceStore, string Name, MemberTypes MemberType, Parameter ItemToPollFor)
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
                NM.CopyMemory(m_IntPtrMemoryMappedFileViewStringStore + m_StringStoreOffset, (IntPtr)PtrName, (UIntPtr)(Name.Length * 2));    //UTF16 charcter = 2 bytes
            }

            PtrMessage->NameOffset = m_StringStoreOffset;
            PtrMessage->NameLength = Name.Length;
            m_StringStoreOffset = m_StringStoreOffset + (Name.Length * 2);
            PtrMessage->MemberType = MemberType;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        unsafe public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType)
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
                NM.CopyMemory(m_IntPtrMemoryMappedFileViewStringStore + m_StringStoreOffset, (IntPtr)PtrName, (UIntPtr)(Name.Length * 2));    //UTF16 charcter = 2 bytes
            }

            PtrMessage->NameOffset = m_StringStoreOffset;
            PtrMessage->NameLength = Name.Length;
            m_StringStoreOffset = m_StringStoreOffset + (Name.Length * 2);
            PtrMessage->MemberType = MemberType;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;

            //Debug.Listeners[0].WriteLine("\t AddMessageQueryMember Message: " + (m_PtrMessageStore->NumberOfMessages - 1).ToString() + " Parameters: " + PtrMessage->NumberOfParameters.ToString());
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6, Parameter p7)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6, Parameter p7, Parameter p8)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        public void AddMessageQueryMember(DataStores SourceStore, DataStores DestinationStore, string Name, MemberTypes MemberType, Parameter p0, Parameter p1, Parameter p2, Parameter p3, Parameter p4, Parameter p5, Parameter p6, Parameter p7, Parameter p8, Parameter p9)
        {
            AddMessageQueryMember(SourceStore, DestinationStore, Name, MemberType);
        }

        unsafe public void DecodeControl(int messageNumber, out ControlIdentifier Identifier)
        {
            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));
            Identifier = new ControlIdentifier();

            // p1  = Parent handle
            if ((PtrMessage->Parameter.TypeCode[0]) == 17)//(Int32)TypeCode.Int32)
            {
                Identifier.ParentHandle = (IntPtr)PtrMessage->Parameter.IntPtr[0];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[0]).ToString());
            }

            // p2  = Handle
            if ((PtrMessage->Parameter.TypeCode[1]) == 17)//(Int32)TypeCode.Int32)
            {
                Identifier.Handle = (IntPtr)PtrMessage->Parameter.IntPtr[1];
            }
            else
            {
                throw new Exception("Expected System.TypeCode.17 got System.TypeCode." + (PtrMessage->Parameter.TypeCode[1]).ToString());
            }

            if (PtrMessage->NumberOfParameters > 2)
            {
                // p3  = Name
                if (PtrMessage->Parameter.StringLength[2] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[2]) == (Int32)TypeCode.String)
                    {
                        Identifier.Name = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[2]), 0, PtrMessage->Parameter.StringLength[2]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[2])).ToString());
                    }
                }

                // p4  = Technology Type (Winforms, WPF, etc)
                if (PtrMessage->Parameter.StringLength[3] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[3]) == (Int32)TypeCode.String)
                    {
                        Identifier.TechnologyType = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[3]), 0, PtrMessage->Parameter.StringLength[3]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[3])).ToString());
                    }
                }

                // p5  = Type Namespace
                if (PtrMessage->Parameter.StringLength[4] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[4]) == (Int32)TypeCode.String)
                    {
                        Identifier.TypeNameSpace = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[4]), 0, PtrMessage->Parameter.StringLength[4]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[4])).ToString());
                    }
                }

                // p6  = Type Name
                if (PtrMessage->Parameter.StringLength[5] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[5]) == (Int32)TypeCode.String)
                    {
                        Identifier.TypeName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[5]), 0, PtrMessage->Parameter.StringLength[5]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[5])).ToString());
                    }
                }

                // p7  = Module Name
                if (PtrMessage->Parameter.StringLength[6] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[6]) == (Int32)TypeCode.String)
                    {
                        Identifier.ModuleName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[6]), 0, PtrMessage->Parameter.StringLength[6]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[6])).ToString());
                    }
                }

                // p8  = Assembly Name
                if (PtrMessage->Parameter.StringLength[7] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[7]) == (Int32)TypeCode.String)
                    {
                        Identifier.AssemblyName = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[7]), 0, PtrMessage->Parameter.StringLength[7]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[7])).ToString());
                    }
                }

                // p9  = Index
                if ((PtrMessage->Parameter.TypeCode[8]) == (Int32)TypeCode.Int32)
                {
                    Identifier.Index = PtrMessage->Parameter.Int32[8];
                }
                else
                {
                    throw new Exception("Expected System.TypeCode." + TypeCode.Int32.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[8])).ToString());
                }

                // p10 = Text
                if (PtrMessage->Parameter.StringLength[9] > 0)
                {
                    if ((PtrMessage->Parameter.TypeCode[9]) == (Int32)TypeCode.String)
                    {
                        Identifier.Text = new string((char*)(m_IntPtrMemoryMappedFileViewStringStore + PtrMessage->Parameter.StringOffset[9]), 0, PtrMessage->Parameter.StringLength[9]);
                    }
                    else
                    {
                        throw new Exception("Expected System.TypeCode." + TypeCode.String.ToString() + " got System.TypeCode." + ((TypeCode)(PtrMessage->Parameter.TypeCode[9])).ToString());
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

        private void GetHandleAndName(Control theControl)
        {
            m_Handle = IntPtr.Zero;
            m_Name = "";

            try
            {
                if (theControl.IsHandleCreated)   //if it doesnt have a handle it can't be visible so ignore it
                {
                    m_Handle = (IntPtr)theControl.GetType().GetProperty("HandleInternal", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(theControl, null);
                    m_Name = theControl.Name;
                }
            }
            catch
            {
            }
        }

        //TODO might want to skip this and just get the control directly via Form.FromHandle?
        private Form FindFormViaHandle(IntPtr FormHandle)
        {
            //Find the form
            foreach (Form Form in Application.OpenForms)
            {
                try
                {
                    GetHandleAndName(Form);
                    if (m_Handle == FormHandle)
                    {
                        return Form;
                    }
                }
                catch
                {
                }
            }

            return null;
        }

        private void BuildControlArray(Control theControl)
        {
            if (m_AllControlsOnForm == null)
            {
                Array.Resize(ref m_AllControlsOnForm, 1);
            }
            else
            {
                Array.Resize(ref m_AllControlsOnForm, m_AllControlsOnForm.Length + 1);
            }

            m_AllControlsOnForm[m_AllControlsOnForm.GetUpperBound(0)] = theControl;

            if (theControl.IsHandleCreated)
            {
                if (theControl.HasChildren)
                {
                    foreach (Control childControl in theControl.Controls)
                    {
                        BuildControlArray(childControl);
                    }
                }
            }
        }

        private static object theLock = new object();

        private static void WriteLog(string Line)
        {
            /*
            lock (theLock)    // Needed as WriteLog is used in callbacks which are async
            {
                TextWriter log = File.AppendText(@"C:\debug.log");
                log.WriteLine(DateTime.Now.ToString() + "\t" + Line);
                log.Close();
            }
            */
        }

        unsafe private bool Find(int messageNumber)
        {
            ControlIdentifier Identifier;
            DecodeControl(messageNumber, out Identifier);

            //Searching for the control by property
            Type theType = null;
            IntPtr Handle = IntPtr.Zero;
            string Name = null;
            bool FoundControl = false;

            if (Identifier.ParentHandle == IntPtr.Zero)     //Find toplevel (parent)
            {
                if (Identifier.Handle != IntPtr.Zero)
                {
                    Control TheControl = Form.FromHandle(Identifier.Handle);
                    if (TheControl != null)
                    {
                        GetHandleAndName(TheControl);
                        Handle = m_Handle;
                        Name = m_Name;
                        theType = TheControl.GetType();

                        IntPtr TopLevel = (IntPtr)TheControl.TopLevelControl.GetType().GetProperty("HandleInternal", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic).GetValue(TheControl.TopLevelControl, null);
                        if (TopLevel != Identifier.Handle)
                        {
                            Identifier.ParentHandle = TopLevel;
                        }                    

                        FoundControl = true;
                    }
                }
                else
                {
                    //Start polling
                    Stopwatch timer = Stopwatch.StartNew();
                    do
                    {
                        //looking for a form
                        try
                        {
                            int CurrentIndex = 0;
                            foreach (Form Form in Application.OpenForms)
                            {
                                theType = Form.GetType();
                                GetHandleAndName(Form);
                                Handle = m_Handle;
                                Name = Form.Name;

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
                                        WriteLog(theType.Name + " != " + Identifier.TypeName);
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
                                    if (!System.Text.RegularExpressions.Regex.IsMatch(GetWindowText(Handle), Identifier.Text))
                                    {
                                        continue;
                                    }
                                }

                                CurrentIndex++;

                                WriteLog("found form for " + Name);

                                if (Identifier.Index > 0)
                                {
                                    //TO DO implement index support!
                                    if (CurrentIndex != Identifier.Index)
                                    {
                                        continue;
                                    }
                                }

                                //we have a match
                                if (NM.IsWindowVisible(Handle))
                                {
                                    FoundControl = true;
                                    break;
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

                        WriteLog("done find form loop for " + Name);

                        if (FoundControl == false)
                        {
                            WriteLog("");
                            //Thread.Sleep(m_Sleep);
                            Thread.Sleep(15);
                        }
                    }
                    while (FoundControl == false && timer.ElapsedMilliseconds < m_TimeOut);
                    timer.Stop();
                }
            }
            else    //find child
            {
                if (Identifier.Handle != IntPtr.Zero)
                {
                    Control TheControl = Form.FromHandle(Identifier.Handle);

                    if (TheControl != null)
                    {
                        GetHandleAndName(TheControl);
                        Handle = m_Handle;
                        Name = m_Name;
                        theType = TheControl.GetType();
                        FoundControl = true;
                    }
                }
                else
                {
                    //looking for a control on a form
                    Form Form = FindFormViaHandle(Identifier.ParentHandle);
                    Control[] Controls;

                    //Start polling
                    Stopwatch timer = Stopwatch.StartNew();
                    do
                    {
                        if (Identifier.Name == null)
                        {
                            foreach (Control Control in Form.Controls)
                            {
                                BuildControlArray(Control);
                            }
                            Controls = m_AllControlsOnForm;
                            m_AllControlsOnForm = null;
                        }
                        else
                        {
                            Controls = Form.Controls.Find(Identifier.Name, true);
                        }

                        int CurrentIndex = 0;
                        foreach (Control Control in Controls)
                        {
                            theType = Control.GetType();

                            GetHandleAndName(Control);
                            Handle = m_Handle;
                            Name = m_Name;

                            //TODO do we want only visible ones? yes we probably do
                            if (NM.IsWindowVisible(Handle))
                            {
                                //if (Identifier.Name != null)
                                //{
                                //    if (Name != Identifier.Name)
                                //    {
                                //        continue;
                                //    }
                                //}

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
                                        WriteLog(theType.Name + " != " + Identifier.TypeName);
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
                                    string WindowText = GetWindowText(Handle);
                                    if (WindowText == null)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        if (!System.Text.RegularExpressions.Regex.IsMatch(WindowText, Identifier.Text))
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
                                if (NM.IsWindowVisible(Handle))
                                {
                                    FoundControl = true;
                                    break;
                                }
                            }
                        }

                        if (FoundControl == false)
                        {
                            WriteLog("");
                            //Thread.Sleep(m_Sleep);
                            Thread.Sleep(15);
                        }
                    }
                    while (FoundControl == false && timer.ElapsedMilliseconds < m_TimeOut);
                    timer.Stop();
                }
            }

            //return the item
            if (FoundControl)
            {
                ControlIdentifier NewIdentifier = new ControlIdentifier();

                NewIdentifier.ParentHandle = Identifier.ParentHandle;
                NewIdentifier.Handle = Handle;
                NewIdentifier.Name = Name;
                NewIdentifier.TechnologyType = GetTechnologyType(theType);
                NewIdentifier.TypeNameSpace = theType.Namespace;
                NewIdentifier.TypeName = theType.Name;
                NewIdentifier.ModuleName = theType.Module.Name;
                NewIdentifier.AssemblyName = theType.Assembly.GetName().Name;
                NewIdentifier.Index = 1;
                NewIdentifier.Text = GetWindowText(Handle);

                AddIdentifierMessage(NewIdentifier);
            }
            //else
            //{
            //    //TODO remove this and just return an error, prevent spamming first chance exceptions
            //    throw new Exception("Failed to find control");
            //}
            return FoundControl;
        }

        unsafe private void Refind(int messageNumber)
        {
            object DestinationObject;

            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((messageNumber - 1) * m_SizeOfMessage));

            ControlIdentifier Identifier;
            DecodeControl(messageNumber, out Identifier);

            IntPtr FormHandle;
            //Refind the control by handle
            if (Identifier.ParentHandle == IntPtr.Zero)
            {
                //We are searching for a form
                FormHandle = Identifier.Handle;
            }
            else
            {
                //We are searching for a contol on a form
                FormHandle = Identifier.ParentHandle;
            }

            //Find the form
            Form Form = FindFormViaHandle(FormHandle);

            if (Identifier.ParentHandle == IntPtr.Zero)
            {
                DestinationObject = Form;
            }
            else
            {
                //Find the control on the form
                DestinationObject = Form.FromHandle(Identifier.Handle);
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
            TypeCode ParameterTypeCode = TypeCode.Empty;
            if (Parameters == 1)
            {
                ParameterTypeCode = (TypeCode)PtrMessage->Parameter.TypeCode[0];
            }

            //cleanup the message
            PtrMessage->TypeCodeKey = 0;
            PtrMessage->NumberOfParameters = 0;
            PtrMessage->NameOffset = 0;
            PtrMessage->NameLength = 0;
            PtrMessage->Action = MessageAction.None;

            //TODO is it worth doing this?
            if (Parameters == 1)
            {
                //fast (explicit casting) 
                //TO DO add rest of switch
                switch (ParameterTypeCode)
                {
                    case TypeCode.Int32:
                        AddReturnValue(new Parameter(this, (Int32)SourceObject));
                        break;
                    case TypeCode.String:
                        AddReturnValue(new Parameter(this, (string)SourceObject));
                        break;
                    default:
                        throw new Exception("Unsupported typecode: " + PtrMessage->Parameter.TypeCode[0].ToString());
                }
            }
            else
            {
                //slow
                if (SourceObject == null)
                {
                    PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + (m_PtrMessageStore->NumberOfMessages * m_SizeOfMessage));
                    PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Empty;
                    PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Empty);
                    PtrMessage->NumberOfParameters++;
                    m_PtrMessageStore->NumberOfMessages++;
                }
                else
                {
                    AddReturnValue(new Parameter(this, (dynamic)SourceObject));
                }
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
                        default:
                            return "";
                    }
                case "TD.SandDock":
                    switch (TypeName)
                    {
                        case "DocumentContainer":
                            return "GUIDocumentContainer";
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
                            return "GUIWalkerControl";
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
                    if (ControlType.Namespace.StartsWith("System.Windows."))
                    {
                        break;
                    }

                    string APEType = GetAPETypeFromTypeInternal(ControlType.Namespace, ControlType.Name);

                    if (APEType != "")
                    {
                        break;
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

            switch ((int)PtrMessage->Parameter.TypeCode[0])
            {
                case (int)TypeCode.Boolean:
                    ItemToPollFor = PtrMessage->Parameter.Boolean[0];
                    break;
                case (int)TypeCode.Char:
                    ItemToPollFor = PtrMessage->Parameter.Char[0];
                    break;
                case (int)TypeCode.SByte:
                    ItemToPollFor = PtrMessage->Parameter.SByte[0];
                    break;
                case (int)TypeCode.Byte:
                    ItemToPollFor = PtrMessage->Parameter.Byte[0];
                    break;
                case (int)TypeCode.Int16:
                    ItemToPollFor = PtrMessage->Parameter.Int16[0];
                    break;
                case (int)TypeCode.UInt16:
                    ItemToPollFor = PtrMessage->Parameter.UInt16[0];
                    break;
                case (int)TypeCode.Int32:
                    ItemToPollFor = PtrMessage->Parameter.Int32[0];
                    break;
                case (int)TypeCode.UInt32:
                    ItemToPollFor = PtrMessage->Parameter.UInt32[0];
                    break;
                case (int)TypeCode.Int64:
                    ItemToPollFor = PtrMessage->Parameter.Int64[0];
                    break;
                case (int)TypeCode.UInt64:
                    ItemToPollFor = PtrMessage->Parameter.UInt64[0];
                    break;
                case (int)TypeCode.Single:
                    ItemToPollFor = PtrMessage->Parameter.Single[0];
                    break;
                case (int)TypeCode.Double:
                    ItemToPollFor = PtrMessage->Parameter.Double[0];
                    break;
                case (int)TypeCode.Decimal:
                    int[] DecimalBits = new int[4];
                    DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[0];
                    DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[0];
                    DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[0];
                    DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[0];
                    ItemToPollFor = new decimal(DecimalBits);
                    break;
                case (int)TypeCode.DateTime:
                    ItemToPollFor = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[0]);
                    break;
                case (int)TypeCode.String:
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
                case 17:
                    ItemToPollFor = new IntPtr(PtrMessage->Parameter.IntPtr[0]);
                    break;
                default:
                    throw new Exception("Unsupported System.TypeCode: " + ((int)PtrMessage->Parameter.TypeCode).ToString());
            }

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
                        MemberGetterCache.GetFromList(SourceType.TypeHandle, Name, out MemberGetter);
                        if (MemberGetter == null)
                        {
                            MemberGetter = SourceType.DelegateForGetFieldValue(Name);
                            MemberGetterCache.AddToList(SourceType.TypeHandle, Name, MemberGetter);
                        }
                        break;
                    case MemberTypes.Property:
                        MemberGetterCache.GetFromList(SourceType.TypeHandle, Name, out MemberGetter);
                        if (MemberGetter == null)
                        {
                            MemberGetter = SourceType.DelegateForGetPropertyValue(Name);
                            MemberGetterCache.AddToList(SourceType.TypeHandle, Name, MemberGetter);
                        }
                        break;
                    default:
                        throw new Exception("Unsupported member type: " + (PtrMessage->MemberType).ToString());
                }

                Stopwatch timer = Stopwatch.StartNew();
                while (true)
                {
                    ItemFound = ((Control)tempStore0).Invoke((Delegate)MemberGetter, SourceObject.WrapIfValueType());
                    
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

        unsafe private void Reflect(int MessageNumber)
        {
            object SourceObject;
            object DestinationObject;

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
                        case (int)TypeCode.Boolean:
                            ParametersType[i] = m_TypeBoolean;
                            ParametersObject[i] = PtrMessage->Parameter.Boolean[i];
                            break;
                        case (int)TypeCode.Char:
                            ParametersType[i] = m_TypeChar;
                            ParametersObject[i] = PtrMessage->Parameter.Char[i];
                            break;
                        case (int)TypeCode.SByte:
                            ParametersType[i] = m_TypeSByte;
                            ParametersObject[i] = PtrMessage->Parameter.SByte[i];
                            break;
                        case (int)TypeCode.Byte:
                            ParametersType[i] = m_TypeByte;
                            ParametersObject[i] = PtrMessage->Parameter.Byte[i];
                            break;
                        case (int)TypeCode.Int16:
                            ParametersType[i] = m_TypeInt16;
                            ParametersObject[i] = PtrMessage->Parameter.Int16[i];
                            break;
                        case (int)TypeCode.UInt16:
                            ParametersType[i] = m_TypeUInt16;
                            ParametersObject[i] = PtrMessage->Parameter.UInt16[i];
                            break;
                        case (int)TypeCode.Int32:
                            ParametersType[i] = m_TypeInt32;
                            ParametersObject[i] = PtrMessage->Parameter.Int32[i];
                            break;
                        case (int)TypeCode.UInt32:
                            ParametersType[i] = m_TypeUInt32;
                            ParametersObject[i] = PtrMessage->Parameter.UInt32[i];
                            break;
                        case (int)TypeCode.Int64:
                            ParametersType[i] = m_TypeInt64;
                            ParametersObject[i] = PtrMessage->Parameter.Int64[i];
                            break;
                        case (int)TypeCode.UInt64:
                            ParametersType[i] = m_TypeUInt64;
                            ParametersObject[i] = PtrMessage->Parameter.UInt64[i];
                            break;
                        case (int)TypeCode.Single:
                            ParametersType[i] = m_TypeSingle;
                            ParametersObject[i] = PtrMessage->Parameter.Single[i];
                            break;
                        case (int)TypeCode.Double:
                            ParametersType[i] = m_TypeDouble;
                            ParametersObject[i] = PtrMessage->Parameter.Double[i];
                            break;
                        case (int)TypeCode.Decimal:
                            ParametersType[i] = m_TypeDecimal;
                            int[] DecimalBits = new int[4];
                            DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[i];
                            DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[i];
                            DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[i];
                            DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[i];
                            ParametersObject[i] = new decimal(DecimalBits);
                            break;
                        case (int)TypeCode.DateTime:
                            ParametersType[i] = m_TypeDateTime;
                            ParametersObject[i] = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[i]);
                            break;
                        case (int)TypeCode.String:
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
                        case 17:
                            ParametersType[i] = m_TypeIntPtr;
                            ParametersObject[i] = new IntPtr(PtrMessage->Parameter.IntPtr[i]);
                            break;
                        default:
                            throw new Exception("Unsupported System.TypeCode: " + ((int)(PtrMessage->Parameter.TypeCode[i])).ToString());
                    }
                }
                ParametersTypeCache.AddToList(PtrMessage->TypeCodeKey, ParametersType);
            }
            else
            {
                for (int i = 0; i < PtrMessage->NumberOfParameters; i++)
                {
                    switch ((int)PtrMessage->Parameter.TypeCode[i])
                    {
                        case (int)TypeCode.Boolean:
                            ParametersObject[i] = PtrMessage->Parameter.Boolean[i];
                            break;
                        case (int)TypeCode.Char:
                            ParametersObject[i] = PtrMessage->Parameter.Char[i];
                            break;
                        case (int)TypeCode.SByte:
                            ParametersObject[i] = PtrMessage->Parameter.SByte[i];
                            break;
                        case (int)TypeCode.Byte:
                            ParametersObject[i] = PtrMessage->Parameter.Byte[i];
                            break;
                        case (int)TypeCode.Int16:
                            ParametersObject[i] = PtrMessage->Parameter.Int16[i];
                            break;
                        case (int)TypeCode.UInt16:
                            ParametersObject[i] = PtrMessage->Parameter.UInt16[i];
                            break;
                        case (int)TypeCode.Int32:
                            ParametersObject[i] = PtrMessage->Parameter.Int32[i];
                            break;
                        case (int)TypeCode.UInt32:
                            ParametersObject[i] = PtrMessage->Parameter.UInt32[i];
                            break;
                        case (int)TypeCode.Int64:
                            ParametersObject[i] = PtrMessage->Parameter.Int64[i];
                            break;
                        case (int)TypeCode.UInt64:
                            ParametersObject[i] = PtrMessage->Parameter.UInt64[i];
                            break;
                        case (int)TypeCode.Single:
                            ParametersObject[i] = PtrMessage->Parameter.Single[i];
                            break;
                        case (int)TypeCode.Double:
                            ParametersObject[i] = PtrMessage->Parameter.Double[i];
                            break;
                        case (int)TypeCode.Decimal:
                            int[] DecimalBits = new int[4];
                            DecimalBits[0] = PtrMessage->Parameter.DecimalBits0[i];
                            DecimalBits[1] = PtrMessage->Parameter.DecimalBits1[i];
                            DecimalBits[2] = PtrMessage->Parameter.DecimalBits2[i];
                            DecimalBits[3] = PtrMessage->Parameter.DecimalBits3[i];
                            ParametersObject[i] = new decimal(DecimalBits);
                            break;
                        case (int)TypeCode.DateTime:
                            ParametersObject[i] = DateTime.FromBinary(PtrMessage->Parameter.DateTimeBinary[i]);
                            break;
                        case (int)TypeCode.String:
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
                        case 17:
                            ParametersObject[i] = new IntPtr(PtrMessage->Parameter.IntPtr[i]);
                            break;
                        default:
                            throw new Exception("Unsupported System.TypeCode: " + ((int)PtrMessage->Parameter.TypeCode).ToString());
                    }
                }
            }

            Type SourceType;
            string Name;
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
                        case MemberTypes.Field:
                            MemberGetterCache.GetFromList(SourceType.TypeHandle, Name, out MemberGetter);
                            if (MemberGetter == null)
                            {
                                MemberGetter = SourceType.DelegateForGetFieldValue(Name);
                                MemberGetterCache.AddToList(SourceType.TypeHandle, Name, MemberGetter);
                            }
                            DestinationObject = ((Control)tempStore0).Invoke((Delegate)MemberGetter, SourceObject.WrapIfValueType());
                            MemberGetter = null;
                            break;
                        case MemberTypes.Property:
                            if (ParametersType.Length == 0)
                            {
                                MemberGetterCache.GetFromList(SourceType.TypeHandle, Name, out MemberGetter);
                                if (MemberGetter == null)
                                {
                                    MemberGetter = SourceType.DelegateForGetPropertyValue(Name);
                                    MemberGetterCache.AddToList(SourceType.TypeHandle, Name, MemberGetter);
                                }
                                DestinationObject = ((Control)tempStore0).Invoke((Delegate)MemberGetter, SourceObject.WrapIfValueType());
                                MemberGetter = null;
                            }
                            else
                            {
                                MethodInvokerCache.GetFromList(SourceType.TypeHandle, Name, PtrMessage->TypeCodeKey, out MethodInvoker);
                                if (MethodInvoker == null)
                                {
                                    MethodInvoker = SourceType.DelegateForGetIndexer(ParametersType);
                                    MethodInvokerCache.AddToList(SourceType.TypeHandle, Name, PtrMessage->TypeCodeKey, MethodInvoker);
                                }
                                DestinationObject = ((Control)tempStore0).Invoke((Delegate)MethodInvoker, SourceObject.WrapIfValueType(), ParametersObject);
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
                                MethodInvokerCache.GetFromList(SourceType.TypeHandle, Name, PtrMessage->TypeCodeKey, out MethodInvoker);
                                if (MethodInvoker == null)
                                {
                                    MethodInvoker = SourceType.DelegateForCallMethod(Name, ParametersType);
                                    MethodInvokerCache.AddToList(SourceType.TypeHandle, Name, PtrMessage->TypeCodeKey, MethodInvoker);
                                }
                                DestinationObject = ((Control)tempStore0).Invoke((Delegate)MethodInvoker, SourceObject.WrapIfValueType(), ParametersObject);
                            }
                            MethodInvoker = null;
                            break;
                        default:
                            throw new Exception("Unsupported member type: " + (PtrMessage->MemberType).ToString());
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

        [EditorBrowsable(EditorBrowsableState.Never)]
        unsafe public void DecodeMessage(int APEPID, string AppDomainToLoadInto)
        {
            try
            {
                string AUTProcessId = Process.GetCurrentProcess().Id.ToString();
                string APEProcessId = APEPID.ToString();

                m_HandleMemoryMappedFileStringStore = NM.CreateFileMapping((IntPtr)(NM.INVALID_HANDLE_VALUE), (IntPtr)0, NM.FileMapProtection.PageReadWrite, 0, StringSpaceBytes, APEProcessId + "_String_" + AppDomainToLoadInto + "_" + AUTProcessId);
                m_IntPtrMemoryMappedFileViewStringStore = NM.MapViewOfFile(m_HandleMemoryMappedFileStringStore, NM.FileMapAccess.FileMapAllAccess, 0, 0, (UIntPtr)StringSpaceBytes);
                m_HandleMemoryMappedFileMessageStore = NM.CreateFileMapping((IntPtr)(NM.INVALID_HANDLE_VALUE), (IntPtr)0, NM.FileMapProtection.PageReadWrite, 0, (uint)sizeof(MessageStore), APEProcessId + "_Message_" + AppDomainToLoadInto + "_" + AUTProcessId);
                m_IntPtrMemoryMappedFileViewMessageStore = NM.MapViewOfFile(m_HandleMemoryMappedFileMessageStore, NM.FileMapAccess.FileMapAllAccess, 0, 0, (UIntPtr)sizeof(MessageStore));
                m_PtrMessageStore = (MessageStore*)m_IntPtrMemoryMappedFileViewMessageStore.ToPointer();

                m_eventIPC = new EventWaitHandle(false, EventResetMode.AutoReset, APEProcessId + "_EventIPC_" + AppDomainToLoadInto + "_" + AUTProcessId);

                Side = EventSet.AUT;
                try
                {
                    ApeProcess = Process.GetProcessById(APEPID);
                }
                catch
                {
                }

                // Create an instance of hook procedures
                MouseHookProcedure = new NM.HookProc(MouseHookProc);
                EnumThreadProcedue = new NM.EnumThread(EnumThreadCallback);

                while (true)
                {
                    WaitForMessages(EventSet.AUT);
                    if (m_Abort)
                    {
                        RemoveFileMapping();
                        break;
                    }

                    bool Result = true;
                    string ResultText = "";

                    try
                    {
                        int NumberOfMessages = m_PtrMessageStore->NumberOfMessages;
                        m_PtrMessageStore->NumberOfMessages = 0;
                        m_StringStoreOffset = 0;

                        for (int MessageNumber = 1; MessageNumber <= NumberOfMessages; MessageNumber++)
                        {
                            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

                            WriteLog("Processing message " + PtrMessage->Action.ToString());
                            //get the message action:
                            switch (PtrMessage->Action)
                            {
                                case MessageAction.RemoveFileMapping:
                                    RemoveFileMapping();
                                    return;
                                case MessageAction.GetListViewGroupRectangle:
                                    GetListViewGroupRectangle(MessageNumber);
                                    break;
                                case MessageAction.GetListViewItemRectangle:
                                    GetListViewItemRectangle(MessageNumber);
                                    break;
                                case MessageAction.Find:
                                    Result = Find(MessageNumber);
                                    if (!Result)
                                    {
                                        ResultText = "Failed to find control";
                                    }
                                    break;
                                case MessageAction.Refind:
                                    Refind(MessageNumber);
                                    break;
                                case MessageAction.ReflectGet:
                                    Reflect(MessageNumber);
                                    break;
                                case MessageAction.ReflectPoll:
                                    ReflectPoll(MessageNumber);
                                    break;
                                case MessageAction.GetResult:
                                    GetResult(MessageNumber);
                                    break;
                                case MessageAction.AddMouseHook:
                                    AddMouseHook(MessageNumber);
                                    break;
                                case MessageAction.RemoveMouseHook:
                                    RemoveMouseHook(MessageNumber);
                                    break;
                                case MessageAction.WaitForMouseState:
                                    WaitForMouseState(MessageNumber);
                                    break;
                                case MessageAction.SetTimeOuts:
                                    SetTimeOuts(MessageNumber);
                                    break;
                                case MessageAction.GetTitleBarItemRectangle:
                                    GetTitleBarItemRectangle(MessageNumber);
                                    break;
                                case MessageAction.GarbageCollect:
                                    GarbageCollect(MessageNumber);
                                    break;
                                case MessageAction.GetContextMenuStrip:
                                    GetContextMenuStrip(MessageNumber);
                                    break;
                                case MessageAction.GetAppDomains:
                                    GetAppDomains(MessageNumber);
                                    break;
                                case MessageAction.GetRecognisedType:
                                    GetRecognisedType(MessageNumber);
                                    break;
                                case MessageAction.GetApeTypeFromType:
                                    GetApeTypeFromType(MessageNumber);
                                    break;
                                case MessageAction.GetApeTypeFromObject:
                                    GetApeTypeFromObject(MessageNumber);
                                    break;
                                default:
                                    throw new Exception("Unknown action for message " + MessageNumber.ToString() + " : " + PtrMessage->Action.ToString());
                            }

                            if (!Result)
                            {
                                break;
                            }
                        }

                        if (Result)
                        {
                            AddResultMessage(MessageResult.Success);
                        }
                    }
                    catch (Exception ex)
                    {
                        Result = false;
                        ResultText = ex.GetType().Name + " " + ex.Message + "\r\n" + ex.StackTrace;
                    }

                    if (!Result)
                    {
                        //clean up all the messages
                        for (int MessageNumber = 1; MessageNumber <= MaxMessages; MessageNumber++)
                        {
                            Message* PtrMessage = (Message*)(m_IntPtrMemoryMappedFileViewMessageStore + ((MessageNumber - 1) * m_SizeOfMessage));

                            PtrMessage->TypeCodeKey = 0;
                            PtrMessage->NumberOfParameters = 0;
                            PtrMessage->NameOffset = 0;
                            PtrMessage->NameLength = 0;
                            PtrMessage->Action = MessageAction.None;
                        }

                        m_PtrMessageStore->NumberOfMessages = 0;
                        m_StringStoreOffset = 0;

                        AddResultMessage(MessageResult.Failure, ResultText);
                    }

                    //clear the data stores so we don't hold any references to objects in the AUT
                    //which would stop them being garbage collected
                    tempStore0 = null;
                    tempStore1 = null;
                    tempStore2 = null;
                    tempStore3 = null;
                    tempStore4 = null;
                    tempStore5 = null;
                    tempStore6 = null;
                    tempStore7 = null;
                    tempStore8 = null;
                    tempStore9 = null;

                    //send back our response
                    SendMessages(EventSet.AUT);
                }
            }
            catch (Exception ex)
            {
                TextWriter log = File.AppendText(@"C:\critical.log");
                log.WriteLine(DateTime.Now.ToString() + "\t" + ex.Message);
                log.WriteLine(DateTime.Now.ToString() + "\t" + ex.StackTrace);
                log.Close();

                throw;
            }
        }


        public string GetWindowText(IntPtr Handle)
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

            return null;
        }

        unsafe private void AddReturnValue(Parameter p0)
        {
            m_PtrMessageStore->NumberOfMessages++;
        }

        private NM.KEYBDINPUT CreateKeyboardInput(NM.VirtualKeyShort wVK, NM.ScanCodeShort wScan, NM.KEYEVENTF dwFlags, uint time, UIntPtr dwExtraInfo)
        {
            NM.KEYBDINPUT Result = new NM.KEYBDINPUT();

            Result.wVk = wVK;
            Result.wScan = wScan;
            Result.dwFlags = dwFlags;
            Result.time = time;
            Result.dwExtraInfo = dwExtraInfo;

            return Result;
        }

        private NM.MOUSEINPUT CreateMouseInput(int x, int y, uint data, uint time, NM.MOUSEEVENTF flag)
        {
            NM.MOUSEINPUT Result = new NM.MOUSEINPUT();
            double dx = ((double)x * GetNormaliseFactor(Direction.HORIZONTAL));
            double dy = ((double)y * GetNormaliseFactor(Direction.VERTICAL));

            Result.dx = Convert.ToInt32(dx);
            Result.dy = Convert.ToInt32(dy);
            Result.mouseData = data;
            Result.time = time;
            Result.dwFlags = flag;

            //MessageBox.Show(dx.ToString() + " " + Result.dx.ToString() + " " + dy.ToString() + " " + Result.dy.ToString());
            return Result;
        }

        private enum Direction : int
        {
            VERTICAL = 1,
            HORIZONTAL = 0
        }

        private double GetNormaliseFactor(Direction Axis)
        {
            switch (Axis)
            {
                case Direction.HORIZONTAL:
                    return (double)65535 / NM.GetSystemMetrics(NM.SystemMetric.SM_CXSCREEN);
                case Direction.VERTICAL:
                    return (double)65535 / NM.GetSystemMetrics(NM.SystemMetric.SM_CYSCREEN);
            }
            return 0;
        }

        public void MouseClick(MouseButton Button, Boolean Down, Boolean Up, int Clicks, bool ControlKey, bool ShiftKey)
        {
            NM.INPUT[] inputEvent = null;
            int Events = 0;

            if (Clicks != 1)
            {
                if (!Up)
                {
                    throw new Exception("Can only single click when using MouseDown");
                }

                if (!Down)
                {
                    throw new Exception("Can only single click when using MouseUp");
                }
            }

            if (Down)
            {
                if (ControlKey)
                {
                    Events++;
                    Array.Resize(ref inputEvent, Events);
                    inputEvent[inputEvent.GetUpperBound(0)].type = NM.INPUT_TYPE.INPUT_KEYBOARD;
                    inputEvent[inputEvent.GetUpperBound(0)].U.ki = CreateKeyboardInput(NM.VirtualKeyShort.CONTROL, NM.ScanCodeShort.CONTROL, NM.KEYEVENTF.NONE, 0, UIntPtr.Zero);
                }

                if (ShiftKey)
                {
                    Events++;
                    Array.Resize(ref inputEvent, Events);
                    inputEvent[inputEvent.GetUpperBound(0)].type = NM.INPUT_TYPE.INPUT_KEYBOARD;
                    inputEvent[inputEvent.GetUpperBound(0)].U.ki = CreateKeyboardInput(NM.VirtualKeyShort.SHIFT, NM.ScanCodeShort.SHIFT, NM.KEYEVENTF.NONE, 0, UIntPtr.Zero);
                }
            }

            //TODO pass in the x and y and set the absolute flag so we are 100% clicking in the right place
            for (int i = 0; i < Clicks; i++)
            {
                if (Down)
                {
                    Events++;
                    Array.Resize(ref inputEvent, Events);
                    inputEvent[inputEvent.Length - 1].type = NM.INPUT_TYPE.INPUT_MOUSE;

                    switch (Button)
                    {
                        case MouseButton.Left:
                            inputEvent[inputEvent.GetUpperBound(0)].U.mi = CreateMouseInput(0, 0, 0, 0, NM.MOUSEEVENTF.LEFTDOWN);
                            break;
                        case MouseButton.Right:
                            inputEvent[inputEvent.GetUpperBound(0)].U.mi = CreateMouseInput(0, 0, 0, 0, NM.MOUSEEVENTF.RIGHTDOWN);
                            break;
                        case MouseButton.Middle:
                            inputEvent[inputEvent.GetUpperBound(0)].U.mi = CreateMouseInput(0, 0, 0, 0, NM.MOUSEEVENTF.MIDDLEDOWN);
                            break;
                    }
                }

                if (Up)
                {
                    Events++;
                    Array.Resize(ref inputEvent, Events);
                    inputEvent[inputEvent.GetUpperBound(0)].type = NM.INPUT_TYPE.INPUT_MOUSE;

                    switch (Button)
                    {
                        case MouseButton.Left:
                            inputEvent[inputEvent.GetUpperBound(0)].U.mi = CreateMouseInput(0, 0, 0, 0, NM.MOUSEEVENTF.LEFTUP);
                            break;
                        case MouseButton.Right:
                            inputEvent[inputEvent.GetUpperBound(0)].U.mi = CreateMouseInput(0, 0, 0, 0, NM.MOUSEEVENTF.RIGHTUP);
                            break;
                        case MouseButton.Middle:
                            inputEvent[inputEvent.GetUpperBound(0)].U.mi = CreateMouseInput(0, 0, 0, 0, NM.MOUSEEVENTF.MIDDLEUP);
                            break;
                    }
                }
            }

            if (Up)
            {
                if (ControlKey)
                {
                    Events++;
                    Array.Resize(ref inputEvent, Events);
                    inputEvent[inputEvent.GetUpperBound(0)].type = NM.INPUT_TYPE.INPUT_KEYBOARD;
                    inputEvent[inputEvent.GetUpperBound(0)].U.ki = CreateKeyboardInput(NM.VirtualKeyShort.CONTROL, NM.ScanCodeShort.CONTROL, NM.KEYEVENTF.KEYUP, 0, UIntPtr.Zero);
                }

                if (ShiftKey)
                {
                    Events++;
                    Array.Resize(ref inputEvent, Events);
                    inputEvent[inputEvent.GetUpperBound(0)].type = NM.INPUT_TYPE.INPUT_KEYBOARD;
                    inputEvent[inputEvent.GetUpperBound(0)].U.ki = CreateKeyboardInput(NM.VirtualKeyShort.SHIFT, NM.ScanCodeShort.SHIFT, NM.KEYEVENTF.KEYUP, 0, UIntPtr.Zero);
                }
            }

            NM.SendInput((uint)inputEvent.Length, inputEvent, NM.INPUT.Size);
        }

        private const int MoveSize = 4;

        public NM.tagPoint MouseMove(IntPtr Handle, int x, int y, bool PerformCheck = true)
        {
            NM.tagRect WindowRect;
            NM.GetWindowRect(Handle, out WindowRect);

            NM.tagRect ClientRect;
            NM.GetClientRect(Handle, out ClientRect);

            //TODO fix this as -1 might be a valid move,,, maybe 0 instead or...

            int xOffset;
            if (x == -1)
            {
                xOffset = ClientRect.right / 2;
            }
            else
            {
                xOffset = x;
            }

            int yOffset;
            if (y == -1)
            {
                yOffset = ClientRect.bottom / 2;
            }
            else
            {
                yOffset = y;
            }

            //Convert the window area to screen point
            NM.tagPoint thePoint;
            thePoint.x = WindowRect.left + xOffset;
            thePoint.y = WindowRect.top + yOffset;

            if (NM.MonitorFromPoint(thePoint, NM.MonitorOptions.MONITOR_DEFAULTTONULL) == null)
            {
                throw new Exception("coordinate appears to be offscreen");
            }

            if (PerformCheck)
            {
                IntPtr ChildHandle;

                thePoint.x = xOffset + WindowRect.left;
                thePoint.y = yOffset + WindowRect.top;

                ChildHandle = NM.WindowFromPoint(thePoint);

                //Make sure we are inside the controls window area
                if (Handle != ChildHandle)
                {
                    throw new Exception("Coordinates are not inside the controls area");
                }
            }

            //Get the current mouse location
            NM.tagPoint currentPoint;
            NM.GetCursorPos(out currentPoint);

            //X direction
            int DirectionX;
            if (currentPoint.x <= WindowRect.left + xOffset)
            {
                DirectionX = 1;     //right
            }
            else
            {
                DirectionX = -1;    //left
            }

            //Y direction
            int DirectionY;
            if (currentPoint.y <= WindowRect.top + yOffset)
            {
                DirectionY = 1;     //down
            }
            else
            {
                DirectionY = -1;    //up
            }

            int MoveX = currentPoint.x;
            int MoveY = currentPoint.y;

            while (MoveX != WindowRect.left + xOffset || MoveY != WindowRect.top + yOffset)
            {
                if (MoveX != WindowRect.left + xOffset)
                {
                    if (DirectionX == 1)
                    {
                        if (MoveX + MoveSize > WindowRect.left + xOffset)
                        {
                            MoveX = MoveX + 1;
                        }
                        else
                        {
                            MoveX = MoveX + MoveSize;
                        }
                    }
                    else
                    {
                        if (MoveX - MoveSize < WindowRect.left + xOffset)
                        {
                            MoveX = MoveX - 1;
                        }
                        else
                        {
                            MoveX = MoveX - MoveSize;
                        }
                    }
                }

                if (MoveY != WindowRect.top + yOffset)
                {
                    if (DirectionY == 1)
                    {
                        if (MoveY + MoveSize > WindowRect.top + yOffset)
                        {
                            MoveY = MoveY + 1;
                        }
                        else
                        {
                            MoveY = MoveY + MoveSize;
                        }
                    }
                    else
                    {
                        if (MoveY - MoveSize < WindowRect.top + yOffset)
                        {
                            MoveY = MoveY - 1;
                        }
                        else
                        {
                            MoveY = MoveY - MoveSize;
                        }
                    }
                }

                MoveMouse(MoveX, MoveY);
            }

            return thePoint;
        }

        private void MoveMouse(int x, int y)
        {
            NM.INPUT[] MouseEvent = new NM.INPUT[1];
            MouseEvent[0].type = NM.INPUT_TYPE.INPUT_MOUSE;
            MouseEvent[0].U.mi = CreateMouseInput(x, y, 0, 0, NM.MOUSEEVENTF.ABSOLUTE | NM.MOUSEEVENTF.MOVE);
            NM.SendInput((uint)MouseEvent.Length, MouseEvent, Marshal.SizeOf(MouseEvent[0].GetType()));
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

                AddMessageSetTimeOuts();
                SendMessages(APEIPC.EventSet.APE);
                WaitForMessages(APEIPC.EventSet.APE);
            }
        }

        //TODO set the sleep in the remote apeipc thread!
        //public int Sleep
        //{
        //    get
        //    {
        //        return m_Sleep;
        //    }
        //    set
        //    {
        //        m_Sleep = value;
        //    }
        //}

        unsafe public class Parameter
        {
            public Parameter(APEIPC instance, Boolean param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Boolean[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Boolean;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Boolean);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Byte param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Byte[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Byte;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Byte);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Char param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Char[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Char;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Char);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, DateTime param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.DateTimeBinary[PtrMessage->NumberOfParameters] = param.ToBinary();
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.DateTime;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.DateTime);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Decimal param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));
                int[] DecimalBits = decimal.GetBits(param);

                PtrMessage->Parameter.DecimalBits0[PtrMessage->NumberOfParameters] = DecimalBits[0];
                PtrMessage->Parameter.DecimalBits1[PtrMessage->NumberOfParameters] = DecimalBits[1];
                PtrMessage->Parameter.DecimalBits2[PtrMessage->NumberOfParameters] = DecimalBits[2];
                PtrMessage->Parameter.DecimalBits3[PtrMessage->NumberOfParameters] = DecimalBits[3];

                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Decimal;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Decimal);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Double param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Double[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Double;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Double);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Int16 param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Int16[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Int16;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Int16);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Int32 param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Int32[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Int32;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Int32);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, IntPtr param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.IntPtr[PtrMessage->NumberOfParameters] = param.ToInt64();
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = 17;    //17 is unused so we steal it
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + 17);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Int64 param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Int64[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Int64;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Int64);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, SByte param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.SByte[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.SByte;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.SByte);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, Single param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.Single[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.Single;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.Single);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, String[] param)
            {
                Parameter p = new Parameter(instance, ConvertString1dArrayToString(param));
                //TODO update the type so we know its a 1d array for the other side to decode
            }

            public Parameter(APEIPC instance, String[,] param)
            {
                Parameter p = new Parameter(instance, ConvertString2dArrayToString(param));
                //TODO update the type so we know its a 2d array for the other side to decode
            }

            private string ConvertString1dArrayToString(string[] StringArray)
            {
                StringBuilder sb = new StringBuilder(StringArray.Length * 4);

                for (int y = 0; y < StringArray.GetUpperBound(0) + 1; y++)
                {
                    if (y < StringArray.GetUpperBound(0))
                    {
                        sb.Append("\n");
                    }
                }
                return sb.ToString();
            }

            private string ConvertString2dArrayToString(string[,] StringArray)
            {
                StringBuilder sb = new StringBuilder(StringArray.Length * 4);

                for (int y = 0; y < StringArray.GetUpperBound(0) + 1; y++)
                {
                    for (int x = 0; x < StringArray.GetUpperBound(1) + 1; x++)
                    {
                        sb.Append(StringArray[y, x]);
                        if (x < StringArray.GetUpperBound(1))
                        {
                            sb.Append("\t");
                        }
                        else
                        {
                            if (y < StringArray.GetUpperBound(0))
                            {
                                sb.Append("\n");
                            }
                        }
                    }
                }
                return sb.ToString();
            }

            public Parameter(APEIPC instance, String param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                if (param != null)
                {
                    fixed (void* PtrString = param)
                    {
                        NM.CopyMemory(instance.m_IntPtrMemoryMappedFileViewStringStore + instance.m_StringStoreOffset, (IntPtr)PtrString, (UIntPtr)(param.Length * 2));    //UTF16 charcter = 2 bytes
                    }
                    PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = instance.m_StringStoreOffset;
                    PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = param.Length;
                    instance.m_StringStoreOffset = instance.m_StringStoreOffset + (param.Length * 2);
                }
                else
                {
                    PtrMessage->Parameter.StringOffset[PtrMessage->NumberOfParameters] = -1;
                    PtrMessage->Parameter.StringLength[PtrMessage->NumberOfParameters] = -1;
                }

                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.String;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.String);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, UInt16 param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.UInt16[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.UInt16;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.UInt16);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, UInt32 param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.UInt32[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.UInt32;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.UInt32);
                PtrMessage->NumberOfParameters++;
            }

            public Parameter(APEIPC instance, UInt64 param)
            {
                Message* PtrMessage = (Message*)(instance.m_IntPtrMemoryMappedFileViewMessageStore + (instance.m_PtrMessageStore->NumberOfMessages * instance.m_SizeOfMessage));

                PtrMessage->Parameter.UInt64[PtrMessage->NumberOfParameters] = param;
                PtrMessage->Parameter.TypeCode[PtrMessage->NumberOfParameters] = (int)TypeCode.UInt64;
                PtrMessage->TypeCodeKey = PtrMessage->TypeCodeKey + ((PtrMessage->NumberOfParameters * OneLargerThanTypeCodeEnumMax) + (int)TypeCode.UInt64);
                PtrMessage->NumberOfParameters++;
            }
        }

        private class ParametersTypeCircularList
        {
            Type[][] m_ParametersType;
            Int64[] m_TypeCodeKey;
            int m_OldestItemInList;
            int m_ListSize;

            public void AddToList(Int64 TypeCodeKey, Type[] ParametersType)
            {
                //Add an item to the list replacing the eldest item
                m_TypeCodeKey[m_OldestItemInList] = TypeCodeKey;
                m_ParametersType[m_OldestItemInList] = ParametersType;
                m_OldestItemInList++;
                if (m_OldestItemInList > m_ListSize - 1)
                {
                    m_OldestItemInList = 0;
                }
            }

            public void GetFromList(Int64 TypeCodeKey, out Type[] ParametersType)
            {
                //Search for the item starting at the youngest item in the list
                for (int i = m_OldestItemInList - 1; i >= 0; i--)
                {
                    if (m_TypeCodeKey[i] == TypeCodeKey)
                    {
                        ParametersType = m_ParametersType[i];
                        return;
                    }
                }

                for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
                {
                    if (m_TypeCodeKey[i] == TypeCodeKey)
                    {
                        ParametersType = m_ParametersType[i];
                        return;
                    }
                }

                ParametersType = null;
            }

            public ParametersTypeCircularList(int Length)
            {
                if (Length < 1)
                {
                    throw new Exception("Length must be positive");
                }

                m_ListSize = Length;
                m_ParametersType = new Type[Length][];
                m_TypeCodeKey = new Int64[Length];
            }
        }

        private class MethodInvokerCircularList
        {
            RuntimeTypeHandle[] m_RunTimeTypeHandle;
            string[] m_Name;
            Fasterflect.MethodInvoker[] m_MethodInvoker;
            Int64[] m_TypeCodeKey;
            int m_OldestItemInList;
            int m_ListSize;

            public void AddToList(RuntimeTypeHandle RuntimeTypeHandle, string Name, Int64 TypeCodeKey, Fasterflect.MethodInvoker MethodInvoker)
            {
                //Add an item to the list replacing the eldest item
                m_RunTimeTypeHandle[m_OldestItemInList] = RuntimeTypeHandle;
                m_Name[m_OldestItemInList] = Name;
                m_TypeCodeKey[m_OldestItemInList] = TypeCodeKey;
                m_MethodInvoker[m_OldestItemInList] = MethodInvoker;

                m_OldestItemInList++;
                if (m_OldestItemInList > m_ListSize - 1)
                {
                    m_OldestItemInList = 0;
                }
            }

            public void GetFromList(RuntimeTypeHandle RuntimeTypeHandle, string Name, Int64 TypeCodeKey, out Fasterflect.MethodInvoker MethodInvoker)
            {
                //Search for the item starting at the youngest item in the list
                for (int i = m_OldestItemInList - 1; i >= 0; i--)
                {
                    if (m_RunTimeTypeHandle[i].Equals(RuntimeTypeHandle) && m_Name[i] == Name && m_TypeCodeKey[i] == TypeCodeKey)
                    {
                        MethodInvoker = m_MethodInvoker[i];
                        return;
                    }
                }

                for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
                {
                    if (m_RunTimeTypeHandle[i].Equals(RuntimeTypeHandle) && m_Name[i] == Name && m_TypeCodeKey[i] == TypeCodeKey)
                    {
                        MethodInvoker = m_MethodInvoker[i];
                        return;
                    }
                }

                MethodInvoker = null;
            }

            public MethodInvokerCircularList(int Length)
            {
                if (Length < 1)
                {
                    throw new Exception("Length must be positive");
                }

                m_ListSize = Length;
                m_RunTimeTypeHandle = new RuntimeTypeHandle[Length];
                m_Name = new string[Length];
                m_MethodInvoker = new Fasterflect.MethodInvoker[Length];
                m_TypeCodeKey = new Int64[Length];
            }
        }

        private class MemberGetterCircularList
        {
            RuntimeTypeHandle[] m_RunTimeTypeHandle;
            string[] m_Name;
            Fasterflect.MemberGetter[] m_MemberGetter;
            int m_OldestItemInList;
            int m_ListSize;

            public void AddToList(RuntimeTypeHandle RuntimeTypeHandle, string Name, Fasterflect.MemberGetter MemberGetter)
            {
                //Add an item to the list replacing the eldest item
                m_RunTimeTypeHandle[m_OldestItemInList] = RuntimeTypeHandle;
                m_Name[m_OldestItemInList] = Name;
                m_MemberGetter[m_OldestItemInList] = MemberGetter;

                m_OldestItemInList++;
                if (m_OldestItemInList > m_ListSize - 1)
                {
                    m_OldestItemInList = 0;
                }
            }

            public void GetFromList(RuntimeTypeHandle RuntimeTypeHandle, string Name, out Fasterflect.MemberGetter MemberGetter)
            {
                //Search for the item starting at the youngest item in the list
                for (int i = m_OldestItemInList - 1; i >= 0; i--)
                {
                    if (m_RunTimeTypeHandle[i].Equals(RuntimeTypeHandle) && m_Name[i] == Name)
                    {
                        MemberGetter = m_MemberGetter[i];
                        return;
                    }
                }

                for (int i = m_ListSize - 1; i >= m_OldestItemInList; i--)
                {
                    if (m_RunTimeTypeHandle[i].Equals(RuntimeTypeHandle) && m_Name[i] == Name)
                    {
                        MemberGetter = m_MemberGetter[i];
                        return;
                    }
                }

                MemberGetter = null;
            }

            public MemberGetterCircularList(int Length)
            {
                if (Length < 1)
                {
                    throw new Exception("Length must be positive");
                }

                m_ListSize = Length;
                m_RunTimeTypeHandle = new RuntimeTypeHandle[Length];
                m_Name = new string[Length];
                m_MemberGetter = new Fasterflect.MemberGetter[Length];
            }
        }
    }

    #region Native API Signatures and Types
    
    #endregion
}

