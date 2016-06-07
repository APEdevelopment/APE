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
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using APE.Capture;
using APE.Communication;
using System.Threading;
using System.Drawing.Imaging;
using Microsoft.Win32;
using NM = APE.Native.NativeMethods;
using System.Security.Principal;

namespace APE.Language
{
    public enum CellClickLocation
    {
        LeftSideOfCell = 0,
        CentreOfCell = 1,
        RightSideOfCell = 2,
        ExpandCollapseIconOfCell = 3,
        LeftSideOfTreeItem = 4,
    }

    public static class GUI
    {
        internal static APEIPC m_APE;
        internal static Process m_AttachedProcess;
        internal static ViewPort ViewPort;
        private static Thread tViewPort;
        private static bool m_IsElevatedAdmin = false;

        public delegate void LoggerDelegate(string textToLog, LogItemTypeEnum type);
        public static LoggerDelegate Logger;

        static GUI()
        {
            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            m_IsElevatedAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);

            if (tViewPort == null)
            {
                tViewPort = new Thread(new ThreadStart(DisplayViewPort));
                tViewPort.SetApartmentState(ApartmentState.STA);
                tViewPort.IsBackground = true;
                tViewPort.Start();
            }
        }

        // TODO is this needed?
        public static Graphics CreateGraphics()
        {
            return ViewPort.CreateGraphics();
        }

        public static void WaitForInputIdle(GUIObject Control)
        {
            Input.WaitForInputIdle(Control.Handle, m_APE.TimeOut);
        }

        private static void DisplayViewPort()
        {
            ViewPort = new ViewPort();
            ViewPort.StartPosition = FormStartPosition.Manual;
            ViewPort.Location = new Point(0, Screen.PrimaryScreen.WorkingArea.Height - ViewPort.Height);
            ViewPort.Width = Screen.PrimaryScreen.WorkingArea.Width;
            Application.Run(ViewPort);
        }

        public static void Log(string textToLog, LogItemTypeEnum type)
        {
            while (ViewPort == null || ViewPort.IsHandleCreated == false || ViewPort.Visible == false)
            {
                Thread.Sleep(15);
            }
            ViewPort.AppendToLog(textToLog, type);

            if (GUI.Logger != null)
            {
                Logger.Invoke(textToLog, type);
            }
        }

        public static void AttachToProcess(Process p)
        {
            Log("Attached to process [" + p.ProcessName + "]", LogItemTypeEnum.ApeContext);
            p.WaitForInputIdle();

            //Instead of m_APE.RemoveFileMapping we could do a
            //m_APE = null;
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            if (m_APE != null)
            {
                m_APE.RemoveFileMapping();
            }
            m_APE = new APEIPC(p);
            m_AttachedProcess = p;
        }

        public static Process GetAttachedProcess()
        {
            return m_AttachedProcess;
        }

        public static void SetTimeOuts(int msTimeOut)
        {
            m_APE.TimeOut = (uint)msTimeOut;
        }

        public static int GetTimeOuts()
        {
            return (int)m_APE.TimeOut;
        }

        public static void GarbageCollect()
        {
            GarbageCollect(GC.MaxGeneration);
        }

        public static void GarbageCollect(int generation)
        {
            m_APE.AddMessageGarbageCollect(generation);
            m_APE.SendMessages(APEIPC.EventSet.APE);
            m_APE.WaitForMessages(APEIPC.EventSet.APE);

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        internal static bool IsElevatedAdmin
        {
            get
            {
                return m_IsElevatedAdmin;
            }
        }

        internal static bool FormAnimationDisabled
        {
            get
            {
                string MinAnimate = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Control Panel\Desktop\WindowMetrics", "MinAnimate", "-1");
                if (MinAnimate == "0")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }
    }

    public enum Identifiers : int
    {
        Handle = 1, //do we want to allow this?
        Name = 2,
        TechnologyType = 3,
        TypeNameSpace = 4,
        TypeName = 5,
        ModuleName = 6,
        AssemblyName = 7,
        Index = 8,
        Text = 9,
        ChildOf = 10
    }

    [Flags]
    public enum MouseKeyModifier : int
    {
        None = 0,
        Control = 2,
        Shift = 4,
    }

    //TODO replace with MouseButtons.Left from winforms?
    public enum MouseButton : int
    {
        Left = 0,
        Right = 1,
        Middle = 2,
    }

    public class Identifier
    {
        public dynamic IdentifierValue;
        public Identifiers IdentifierType;

        public Identifier(Identifiers type, dynamic value)
        {
            IdentifierType = type;
            IdentifierValue = value;
        }
    }

    interface IGUITextBox
    {
        void SetText(string text);
    }

    interface IGUIForm
    {
        //GUIObject
        string ClassName { get; }
        void MouseSingleClick(int X, int Y, MouseButton Button);
        void MouseSingleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys);
        void MouseDoubleClick(int X, int Y, MouseButton Button);
        void MouseTripleClick(int X, int Y, MouseButton Button);
        void MouseDown(int X, int Y, MouseButton Button);
        void MouseUp(int X, int Y, MouseButton Button);
        bool IsEnabled { get; }
        bool Exists { get; }
        long ExtendedStyle { get; }
        int Height { get; }
        IntPtr Handle { get; }
        int Id { get; }
        int Left { get; }
        string Name { get; }
        long Style { get; }
        string Text { get; }
        int Top { get; }
        string TypeName { get; }
        bool IsVisible { get; }
        int Width { get; }
        string TechnologyType { get; }
        string TypeNameSpace { get; }
        string ModuleName { get; }
        string AssemblyName { get; }

        //GUIFocusableObject
        bool HasFocus { get; }
        void SetFocus();

        //GUIForm
        //void SendKeys(string TextToSend);
        //void Activate();
        void Close();
        void Maximise();
        void Minimise();
        void Move(int X, int Y);
        void Restore();
        void Scroll();
        void Resize();
    }

    internal class AnimationUtils
    {
        public enum WaitForAnimationSource
        {
            Form = 1,
        }

        private static bool IsTopLevelWindow(IntPtr Window)
        {
            NM.WindowStyles Style = (NM.WindowStyles)(long)NM.GetWindowLongPtr(Window, NM.GWL.GWL_STYLE);

            // WS_OVERLAPPED and WS_POPUP indicate a top level window.
            // WS_OVERLAPPED constant is 0, it does not make a good mask.  But all
            // WS_OVERLAPPED windows MUST have a caption so use WS_CAPTION instead.
            return Style.HasFlag(NM.WindowStyles.WS_CAPTION) || Style.HasFlag(NM.WindowStyles.WS_POPUP);
        }

        internal void WaitForAnimation(IntPtr Handle, bool ClearClientArea, WaitForAnimationSource Source)
        {
            if (Source == WaitForAnimationSource.Form)
            {
                if (GUI.FormAnimationDisabled)
                {
                    return;
                }
            }

            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Window is not visible");
                }

                Thread.Sleep(15);
            }
            while (NM.IsWindowVisible(Handle) == false);

            NM.tagRect theRect = Display.GetWindowRectangleDIP(Handle);

            int Width;
            //if (theRect.Right - theRect.Left > 50)
            //{
            //    Width = 50;
            //}
            //else
            //{
                Width = theRect.right - theRect.left;
            //}

            int Height;
            //if (theRect.Bottom - theRect.Top > 50)
            //{
            //    Height = 50;
            //}
            //else
            //{
                Height = theRect.bottom - theRect.top;
            //}

            Bitmap A = new Bitmap(Width, Height, PixelFormat.Format24bppRgb);
            Bitmap B = new Bitmap(Width, Height, PixelFormat.Format24bppRgb);

            int SameCount = 0;
            int i = 0;
            int Loops = 0;
            //Debug.WriteLine("till bitmap loop " + timer.ElapsedMilliseconds.ToString());
            do
            {
                Loops++;

                if (i == 0)
                {
                    Display.GetWindowBitmap(Handle, theRect.left, theRect.top, ref A, false, ClearClientArea);

                    //A.Save(@"c:\temp\" + y.ToString() + ".png");
                    //A.Save(@"c:\temp\a.png");
                    i++;
                }
                else
                {
                    Display.GetWindowBitmap(Handle, theRect.left, theRect.top, ref B, false, ClearClientArea);

                    //B.Save(@"c:\temp\" + y.ToString() + ".png");
                    //B.Save(@"c:\temp\b.png");
                    i--;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    timer.Stop();
                    throw new Exception("Failed to detect window finishing animation");
                }

                if (Display.CompareBitmap(A, B))
                {
                    SameCount++;
                    //Debug.WriteLine("same " + SameCount.ToString() + " " + timer.ElapsedMilliseconds.ToString());
                }
                else
                {
                    SameCount = 0;
                    //Debug.WriteLine("reset " + timer.ElapsedMilliseconds.ToString());
                }

                //Thread.Sleep(15);
            }
            while (SameCount < 2 || Loops < 9);

            timer.Stop();
            //Debug.Listeners[0].WriteLine("\t Loops: " + Loops.ToString() + " Time: " + timer.ElapsedMilliseconds.ToString());
        }
    }

    internal static class VersionUtils
    {
        internal static bool IsWindowsVistaOrHigher
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major >= 6);
            }
        }

        internal static bool IsWindows7OrHigher
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major >= 6) && (Environment.OSVersion.Version.Minor >= 1);
            }
        }

        internal static bool IsWindows8OrHigher
        {
            get
            {
                return (Environment.OSVersion.Platform == PlatformID.Win32NT) && (Environment.OSVersion.Version.Major >= 6) && (Environment.OSVersion.Version.Minor >= 2);
            }
        }
    }

    internal class MenuUtils
    {
        public void ClickMenuItem(IntPtr Parent, IntPtr Control, int MenuIndex, string MenuItem, ref ControlIdentifier ControlIdentity)
        {
            bool Visible;
            bool Enabled;
            int Width;
            int Height;
            int X;
            int Y;
            
            //Get visible, enabled and bounds
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, MenuIndex));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "Enabled", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store5, "Bounds", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store5, DataStores.Store6, "Width", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store5, DataStores.Store7, "Height", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store5, DataStores.Store8, "X", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store5, DataStores.Store9, "Y", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddMessageGetValue(DataStores.Store6);
            GUI.m_APE.AddMessageGetValue(DataStores.Store7);
            GUI.m_APE.AddMessageGetValue(DataStores.Store8);
            GUI.m_APE.AddMessageGetValue(DataStores.Store9);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            Visible = GUI.m_APE.GetValueFromMessage();
            Enabled = GUI.m_APE.GetValueFromMessage();
            Width = GUI.m_APE.GetValueFromMessage();
            Height = GUI.m_APE.GetValueFromMessage();
            X = GUI.m_APE.GetValueFromMessage();
            Y = GUI.m_APE.GetValueFromMessage();

            //Check its enabled, visible and click on it
            if (Visible)
            {
                if (Enabled)
                {
                    IntPtr Temp = ControlIdentity.Handle;
                    try
                    {
                        Input.MouseSingleClick(Parent, Control, X + (Width / 2), Y + (Height / 2), MouseButton.Left, MouseKeyModifier.None);
                    }
                    finally
                    {
                        ControlIdentity.Handle = Temp;
                    }
                }
                else
                {
                    throw new Exception("Menu item [" + MenuItem + "] not enabled");
                }
            }
            else
            {
                throw new Exception("Menu item [" + MenuItem + "] not visible");
            }
        }

        public IntPtr GetDropDown(IntPtr Parent, IntPtr Control, int MenuIndex)
        {
            bool HasDropDown;

            //check we have a drop down
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, MenuIndex));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "HasDropDown", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            HasDropDown = GUI.m_APE.GetValueFromMessage();

            if (HasDropDown == false)
            {
                throw new Exception("Menu does not have a drop down");
            }

            //get the dropdown and its handle
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, MenuIndex));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "DropDown", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            return (IntPtr)GUI.m_APE.GetValueFromMessage();
        }

        public int GetIndexOfMenuItem(IntPtr Parent, IntPtr Control, string MenuItem)
        {
            int Items;

            //Get the number of items on the menustrip
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            Items = GUI.m_APE.GetValueFromMessage();

            //Loop through looking for the item we want
            for (int Item = 0; Item < Items; Item++)
            {
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Item));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //get the values returned
                string ItemText = GUI.m_APE.GetValueFromMessage();

                if (ItemText == MenuItem)
                {
                    //found it
                    return Item;
                }
            }

            //Failed to find it
            throw new Exception("Failed to find menu item [" + MenuItem + "]");
        }
    }
}