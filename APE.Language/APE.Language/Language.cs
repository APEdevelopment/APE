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
using System.Drawing;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using APE.Capture;
using APE.Communication;
using NM = APE.Native.NativeMethods;
using System.Threading;
using System.Drawing.Imaging;
using Microsoft.Win32;
using System.Security.Principal;

namespace APE.Language
{
    /// <summary>
    /// The type of information that is being logged
    /// </summary>
    public enum LogItemType
    {
        /// <summary>
        /// Used for logging user interaction with the application such as clicking on a button
        /// </summary>
        Action,
        /// <summary>
        /// Used to denote the start of something
        /// </summary>
        Start,
        /// <summary>
        /// Used to denote the end of something
        /// </summary>
        Finish,
        /// <summary>
        /// Used to denote something has passed
        /// </summary>
        Pass,
        /// <summary>
        /// Used to denote something has failed
        /// </summary>
        Fail,
        /// <summary>
        /// Used to denote something which is disabled
        /// </summary>
        Disabled,
        /// <summary>
        /// Used to denote something which is N/A
        /// </summary>
        NA,
        /// <summary>
        /// Used for logging non-user interaction as debug severity
        /// </summary>
        Debug,
        /// <summary>
        /// Used for logging non-user interaction as information severity
        /// </summary>
        Information,
        /// <summary>
        /// Used for logging non-user interaction as warning severity
        /// </summary>
        Warning,
        /// <summary>
        /// Used for logging non-user interaction as error severity
        /// </summary>
        Error,
    }

    /// <summary>
    /// The property to use to use to identify a control
    /// </summary>
    public enum Identifiers : int
    {
        /// <summary>
        /// Windows handle
        /// </summary>
        Handle = 1,
        /// <summary>
        /// Name property
        /// </summary>
        Name = 2,
        /// <summary>
        /// The technology the control uses.  Valid options are:
        ///   Windows Forms (WinForms)
        ///   Windows Presentation Foundation (WPF)
        ///   Windows Native
        /// </summary>
        TechnologyType = 3,
        /// <summary>
        /// The name space of the type the control belongs to
        /// </summary>
        TypeNameSpace = 4,
        /// <summary>
        /// The name of the type the control belongs to
        /// </summary>
        TypeName = 5,
        /// <summary>
        /// The module name the control belongs to
        /// </summary>
        ModuleName = 6,
        /// <summary>
        /// The assembly name the control belongs to
        /// </summary>
        AssemblyName = 7,
        /// <summary>
        /// The index of the control relative to other control matching other identifier properties
        /// </summary>
        Index = 8,
        /// <summary>
        /// The text caption of the control.  Regular expression syntax is supported see msdn for details:
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx
        /// </summary>
        Text = 9,
        /// <summary>
        /// A control which the control we are identifying is a child of
        /// </summary>
        ChildOf = 10,
        /// <summary>
        /// A control which the control we are identifying is a sibling of
        /// </summary>
        SiblingOf = 11
    }

    /// <summary>
    /// Keyboard keys which modify the action of a mouse click
    /// </summary>
    [Flags]
    public enum MouseKeyModifier : int
    {
        /// <summary>
        /// No keyboard keys are pressed while performing a mouse action
        /// </summary>
        None = 0,
        /// <summary>
        /// The control key is pressed while performing a mouse action
        /// </summary>
        Control = 2,
        /// <summary>
        /// The shift key is pressed while performing a mouse action
        /// </summary>
        Shift = 4,
    }
    
    /// <summary>
    /// The mouse button to press or release while performing a mouse action
    /// </summary>
    public enum MouseButton : int
    {
        /// <summary>
        /// The left mouse button
        /// </summary>
        Left = APEIPC.MouseButton.Left,
        /// <summary>
        /// The right mouse button
        /// </summary>
        Right = APEIPC.MouseButton.Right,
        /// <summary>
        /// The middle mouse button
        /// </summary>
        Middle = APEIPC.MouseButton.Middle,
    }

    /// <summary>
    /// Provides setup and teardown of APE
    /// </summary>
    public static class GUI
    {
        internal static APEIPC m_APE;
        internal static Process m_AttachedProcess;
        internal static ViewPort m_ViewPort;
        private static Thread m_threadViewPort;
        private static bool m_IsElevatedAdmin = false;
        private static int m_MsTimeOut = 30000;

        /// <summary>
        /// Delegate method to provide custom logging
        /// </summary>
        /// <param name="textToLog">The text to be logged</param>
        /// <param name="type">The type of information to be logged</param>
        public delegate void LoggerDelegate(string textToLog, LogItemType type);
        /// <summary>
        /// Assign a method to this field which matches the LoggerDelegate to allow custom logging 
        /// </summary>
        public static LoggerDelegate Logger;

        static GUI()
        {
            if (NM.IsWow64Process(Process.GetCurrentProcess()))
            {
                // On a 64-bit OS APE should run as a 64-bit process
                throw new Exception("Detected running under Wow64, please untick 'prefer 32-bit' in the project build properties page");
            }

            WindowsIdentity identity = WindowsIdentity.GetCurrent();
            WindowsPrincipal principal = new WindowsPrincipal(identity);
            m_IsElevatedAdmin = principal.IsInRole(WindowsBuiltInRole.Administrator);

            if (m_threadViewPort == null)
            {
                m_threadViewPort = new Thread(new ThreadStart(DisplayViewPort));
                m_threadViewPort.SetApartmentState(ApartmentState.STA);
                m_threadViewPort.IsBackground = true;
                m_threadViewPort.Start();
            }
        }
        
        /// <summary>
        /// Waits for the specified controls gui thread to be idle waiting for user input
        /// </summary>
        /// <param name="Control"></param>
        public static void WaitForInputIdle(GUIObject Control)
        {
            Input.WaitForInputIdle(Control.Handle, m_APE.TimeOut);
        }

        private static void DisplayViewPort()
        {
            m_ViewPort = new ViewPort();
            m_ViewPort.StartPosition = FormStartPosition.Manual;
            m_ViewPort.Location = new Point(0, Screen.PrimaryScreen.WorkingArea.Height - m_ViewPort.Height);
            m_ViewPort.Width = Screen.PrimaryScreen.WorkingArea.Width;
            Application.Run(m_ViewPort);
        }

        /// <summary>
        /// Logs the specified text with as specified type to the viewport and any method assigned to the GUI.Logger
        /// </summary>
        /// <param name="textToLog"></param>
        /// <param name="type"></param>
        public static void Log(string textToLog, LogItemType type)
        {
            while (m_ViewPort == null || m_ViewPort.IsHandleCreated == false || m_ViewPort.Visible == false)
            {
                Thread.Sleep(15);
            }
            m_ViewPort.AppendToLog(textToLog, type);

            if (GUI.Logger != null)
            {
                Logger.Invoke(textToLog, type);
            }
        }
        
        /// <summary>
        /// Attaches APE to the specified process so it can automate it
        /// </summary>
        /// <param name="process">The process to attach to</param>
        public static void AttachToProcess(Process process)
        {
            AttachToProcess(process, "DefaultDomain");
        }

        /// <summary>
        /// Attaches APE to the specified process so it can automate it
        /// </summary>
        /// <param name="process">The process to attach to</param>
        /// <param name="domain">The domain in the process to attach to</param>
        public static void AttachToProcess(Process process, string domain)
        {
            Log("Attached to process [" + process.ProcessName + "]", LogItemType.Information);

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (process.MainWindowHandle != IntPtr.Zero)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GetTimeOut())
                {
                    throw new Exception("Failed to locate process main window within " + GetTimeOut() + "ms timeout");
                }
            }
            process.WaitForInputIdle();

            //Instead of GUI.m_APE.RemoveFileMapping we could do a
            //GUI.m_APE = null;
            //GC.Collect();
            //GC.WaitForPendingFinalizers();
            if (m_APE != null)
            {
                m_APE.RemoveFileMapping();
            }
            m_APE = new APEIPC(process, domain);
            m_AttachedProcess = process;

            //Set the default timeout
            SetTimeOut(GetTimeOut());
        }

        /// <summary>
        /// Returns the process object for the currently attached process
        /// </summary>
        /// <returns>The process object</returns>
        public static Process AttachedProcess
        {
            get
            {
                return m_AttachedProcess;
            }
        }

        /// <summary>
        /// If APE is attached to a process and the process exists this property will be true
        /// </summary>
        public static bool AttachedProcessExists
        {
            get
            {
                bool exists = false;
                if (m_AttachedProcess != null && !m_AttachedProcess.HasExited)
                {
                    exists = true;
                }
                return exists;
            }
        }

        /// <summary>
        /// Sets the timeout for an automation action to complete within, in milliseconds
        /// </summary>
        /// <param name="msTimeOut"></param>
        public static void SetTimeOut(int msTimeOut)
        {
            m_MsTimeOut = msTimeOut;
            if (m_APE != null)
            {
                if (!m_AttachedProcess.HasExited)
                {
                    m_APE.TimeOut = (uint)msTimeOut;
                }
            }
        }

        /// <summary>
        /// Gets the timeout value that an automation action must complete within, in milliseconds
        /// </summary>
        /// <returns></returns>
        public static int GetTimeOut()
        {
            return m_MsTimeOut;
        }

        /// <summary>
        /// Performs a max generation garbage collection in both this and the application being automated
        /// </summary>
        public static void GarbageCollect()
        {
            GarbageCollect(GC.MaxGeneration);
        }

        /// <summary>
        /// Gets the handle of the control which currently has focus
        /// </summary>
        public static IntPtr GetFocus()
        {
            return Input.GetFocus();
        }

        /// <summary>
        /// Checks if the specified form currently exists
        /// </summary>
        /// <param name="identParams">The identifier(s) of the form</param>
        /// <returns>True if the form exists otheriwse false</returns>
        public static bool Exists(params Identifier[] identParams)
        {
            return Exists(null, identParams);
        }

        /// <summary>
        /// Checks if the specified control currently exists
        /// </summary>
        /// <param name="parentForm">The identifier(s) of the control</param>
        /// <param name="identParams"></param>
        /// <returns></returns>
        public static bool Exists(GUIForm parentForm, params Identifier[] identParams)
        {
            ControlIdentifier Identity = GUI.BuildIdentity(parentForm, null, identParams);

            GUI.m_APE.AddFirstMessageControlExistsByProperty(Identity);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            GUI.m_APE.DecodeControl(1, out Identity);
            if (Identity.Handle == IntPtr.Zero)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal static ControlIdentifier BuildIdentity(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
        {
            ControlIdentifier Identity = new ControlIdentifier();

            Identity.Description = descriptionOfControl;
            if (parentForm != null)
            {
                Identity.ParentHandle = parentForm.Handle;
            }

            foreach (Identifier i in identParams)
            {
                switch (i.IdentifierType)
                {
                    case Identifiers.Handle:
                        Identity.Handle = i.IdentifierValue;
                        break;
                    case Identifiers.Name:
                        Identity.Name = i.IdentifierValue;
                        break;
                    case Identifiers.TechnologyType:
                        Identity.TechnologyType = i.IdentifierValue;
                        break;
                    case Identifiers.TypeNameSpace:
                        Identity.TypeNameSpace = i.IdentifierValue;
                        break;
                    case Identifiers.TypeName:
                        Identity.TypeName = i.IdentifierValue;
                        break;
                    case Identifiers.ModuleName:
                        Identity.ModuleName = i.IdentifierValue;
                        break;
                    case Identifiers.AssemblyName:
                        Identity.AssemblyName = i.IdentifierValue;
                        break;
                    case Identifiers.Index:
                        Identity.Index = i.IdentifierValue;
                        break;
                    case Identifiers.Text:
                        Identity.Text = i.IdentifierValue;
                        break;
                    case Identifiers.ChildOf:
                        Identity.ChildOf = i.IdentifierValue.Handle;
                        break;
                    case Identifiers.SiblingOf:
                        Identity.SiblingOf = i.IdentifierValue.Handle;
                        break;
                    default:
                        throw new Exception("Unsupported identifier: " + i.ToString());
                }
            }

            if (Identity.TechnologyType == null)
            {
                Identity.TechnologyType = "Windows Forms (WinForms)";
            }

            return Identity;
        }

        /// <summary>
        /// Performs a garbage collection using the specified generation in both this and the application being automated
        /// </summary>
        /// <param name="generation"></param>
        public static void GarbageCollect(int generation)
        {
            m_APE.AddFirstMessageGarbageCollect(generation);
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);

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

    /// <summary>
    /// This class should never be used as its only purpose is to force some references
    /// so that some dependant files get copied local when the language assembly is referenced
    /// </summary>
    class ForceReferences
    {
        /// <summary>
        /// Will raise an exception if this method is ever called, its not supposed to be called
        /// </summary>
        private ForceReferences()
        {
            // Force references for loader so they are copied local
            Type dummyLoaderType32 = typeof(APE.Loader.Dummy32);
            Type dummyLoaderType64 = typeof(APE.Loader.Dummy64);

            // Force references for injector so they are copied local
            Type dummyInjectorType32 = typeof(APE.Injector.Dummy32);
            Type dummyInjectorType64 = typeof(APE.Injector.Dummy64);

            // Force reference for watcher so it is copied local
            Type dummyWatcherType = typeof(APE.Watcher.Watcher);
        }
    }

    /// <summary>
    /// Class which holds an identifier and its value, used to find a control
    /// </summary>
    public class Identifier
    {
        /// <summary>
        /// The value of the identifier
        /// </summary>
        public dynamic IdentifierValue;
        /// <summary>
        /// The type of identifier
        /// </summary>
        public Identifiers IdentifierType;
        
        /// <summary>
        /// Initialises a new instance of the Identifier class using the type and value provided
        /// </summary>
        /// <param name="type">The type of identifier</param>
        /// <param name="value">The value of the identifier</param>
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
        void SingleClick(int X, int Y, MouseButton Button);
        void SingleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys);
        void DoubleClick(int X, int Y, MouseButton Button);
        void TripleClick(int X, int Y, MouseButton Button);
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
        public void ClickMenuItem(IntPtr parent, IntPtr control, int menuIndex, string menuItem, ref ControlIdentifier ControlIdentity)
        {
            bool Visible;
            bool Enabled;
            int Width;
            int Height;
            int X;
            int Y;

            //Check its enabled and visible then click on it
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                //Get visible, enabled and bounds
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, parent, control);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, menuIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "Enabled", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store5, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Width", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Height", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store8, "X", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store9, "Y", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //get the values returned
                Visible = GUI.m_APE.GetValueFromMessage();
                Enabled = GUI.m_APE.GetValueFromMessage();
                Width = GUI.m_APE.GetValueFromMessage();
                Height = GUI.m_APE.GetValueFromMessage();
                X = GUI.m_APE.GetValueFromMessage();
                Y = GUI.m_APE.GetValueFromMessage();

                if (Visible)
                {
                    if (Enabled)
                    {
                        IntPtr Temp = ControlIdentity.Handle;
                        try
                        {
                            Input.MouseSingleClick(parent, control, X + (Width / 2), Y + (Height / 2), MouseButton.Left, MouseKeyModifier.None);
                            break;
                        }
                        finally
                        {
                            ControlIdentity.Handle = Temp;
                        }
                    }
                }
                
                if (timer.ElapsedMilliseconds > 3000)
                {
                    if (!Visible)
                    {
                        throw new Exception("Menu item " + menuItem + " not visible");
                    }
                    else
                    {
                        throw new Exception("Menu item " + menuItem + " not enabled");
                    }
                }

                Thread.Sleep(15);
            }
        }

        public IntPtr GetDropDown(IntPtr Parent, IntPtr Control, int MenuIndex)
        {
            bool HasDropDown;

            //check we have a drop down
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, MenuIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HasDropDown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            HasDropDown = GUI.m_APE.GetValueFromMessage();

            if (HasDropDown == false)
            {
                throw new Exception("Menu does not have a drop down");
            }

            //get the dropdown and its handle
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, MenuIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "DropDown", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            return (IntPtr)GUI.m_APE.GetValueFromMessage();
        }

        public int GetIndexOfMenuItem(IntPtr Parent, IntPtr Control, string MenuItem)
        {
            int Items;

            //Get the number of items on the menustrip
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Parent, Control);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            Items = GUI.m_APE.GetValueFromMessage();

            //Loop through looking for the item we want
            for (int Item = 0; Item < Items; Item++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Parent, Control);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Item));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
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