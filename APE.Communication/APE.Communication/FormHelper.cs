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
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using NM = APE.Native.NativeMethods;
using NV = APE.Native.NativeVersion;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    public partial class APEIPC
    {
        private Type m_ScrollableControl = typeof(WF.ScrollableControl);

        /// <summary>
        /// Initialise the form helper delegates for use
        /// </summary>
        private void SetupFormHelperDelegates()
        {
            m_ScrollControlIntoViewDelegater = new ScrollControlIntoViewDelegate(ScrollControlIntoViewInternal);
            m_PeakMessagDelegater = new PeakMessageDelegate(PeakMessageInternal);
            m_SetFocusDelegater = new SetFocusDelegate(SetFocusInternal);
            m_SetFocusAsyncDelegater = new SetFocusDelegate(SetFocusAsyncInternal);
        }

        //
        //  GetTitleBarItemRectangle
        //

        unsafe public void AddFirstMessageGetTitleBarItemRectangle(IntPtr handle, NM.TitleBarStateElement item)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetTitleBarItemRectangle;

            float screenScalingFactor;
            if (NV.IsWindows10OrHigher)
            {
                screenScalingFactor = 1;
            }
            else
            {
                using (Graphics desktopGraphics = Graphics.FromHwnd(handle))
                {
                    IntPtr desktopDeviceContext = desktopGraphics.GetHdc();
                    int logicalScreenHeight = NM.GetDeviceCaps(desktopDeviceContext, NM.DeviceCap.VERTRES);
                    int physicalScreenHeight = NM.GetDeviceCaps(desktopDeviceContext, NM.DeviceCap.DESKTOPVERTRES);
                    desktopGraphics.ReleaseHdc();
                    screenScalingFactor = (float)physicalScreenHeight / (float)logicalScreenHeight;
                }
            }

            Parameter handleParam = new Parameter(this, handle);
            Parameter itemParam = new Parameter(this, (int)item);
            Parameter screenScalingFactorParam = new Parameter(this, screenScalingFactor);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetTitleBarItemRectangle(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GetTitleBarItemRectangle must be first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);

            // p2  = TitleBarStateElement
            NM.TitleBarStateElement item = (NM.TitleBarStateElement)GetParameterInt32(ptrMessage, 1);

            // p3 = scalingfactor
            float screenScalingFactor = GetParameterSingle(ptrMessage, 2);

            CleanUpMessage(ptrMessage);

            NM.TITLEBARINFOEX tbi = new NM.TITLEBARINFOEX();
            tbi.cbSize = Marshal.SizeOf(typeof(NM.TITLEBARINFOEX));

            // Send the WM_GETTITLEBARINFOEX message
            IntPtr result;
            IntPtr smtReturn;
            smtReturn = NM.SendMessageTimeout(handle, NM.WM_GETTITLEBARINFOEX, IntPtr.Zero, ref tbi, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out result);

            uint state;
            int left;
            int top;
            int right;
            int bottom;

            switch (item)
            {
                case NM.TitleBarStateElement.TitleBar:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.TitleBar];
                    top = tbi.rcTitleBar.top;
                    left = tbi.rcTitleBar.left;
                    bottom = tbi.rcTitleBar.bottom;
                    right = tbi.rcTitleBar.right;
                    break;
                case NM.TitleBarStateElement.Minimize:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Minimize];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].right;
                    break;
                case NM.TitleBarStateElement.Maximize:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Maximize];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].right;
                    break;
                case NM.TitleBarStateElement.Help:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Help];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Help].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Help].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Help].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Help].right;
                    break;
                case NM.TitleBarStateElement.Close:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Close];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Close].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Close].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Close].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Close].right;
                    break;
                default:
                    throw new Exception("Unknown titlebar element: " + item.ToString());
            }

            //HACK - Windows 10 1809 has a bug where by when the window is maximised and you have a docked app bar above it then the top value is wrong
            if (top < tbi.rcTitleBar.top)
            {
                top = tbi.rcTitleBar.top;
            }

            left = (int)(Math.Round((float)(left) / screenScalingFactor));
            top = (int)(Math.Round((float)(top) / screenScalingFactor));
            right = (int)(Math.Round((float)(right) / screenScalingFactor));
            bottom = (int)(Math.Round((float)(bottom) / screenScalingFactor));

            AddReturnValue(new Parameter(this, state));
            AddReturnValue(new Parameter(this, top));
            AddReturnValue(new Parameter(this, left));
            AddReturnValue(new Parameter(this, bottom));
            AddReturnValue(new Parameter(this, right));
        }

        //
        //  ScrollControlIntoView
        //

        private delegate void ScrollControlIntoViewDelegate(IntPtr handle);
        private ScrollControlIntoViewDelegate m_ScrollControlIntoViewDelegater;

        unsafe public void AddFirstMessageScrollControlIntoView(IntPtr handle)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.ScrollControlIntoView;

            Parameter HandleParam = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void ScrollControlIntoView(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("ScrollControlIntoView must be the first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);
            object[] theParameters = { handle };

            //TODO WPF etc
            WF.Control childControl = WF.Control.FromHandle(handle);

            // Check we found the control, if not sliently ignore it as it shouldnt be fatal at this point
            if (childControl != null)
            {
                childControl.Invoke(m_ScrollControlIntoViewDelegater, theParameters);
            }
            CleanUpMessage(ptrMessage);
        }

        private unsafe void ScrollControlIntoViewInternal(IntPtr handle)
        {
            WF.Control childControl = WF.Control.FromHandle(handle);
            if (childControl != null)
            {
                WF.Control control = childControl.Parent;
                while (control != null)
                {
                    if (control.GetType().IsSubclassOf(m_ScrollableControl))
                    {
                        ((WF.ScrollableControl)control).ScrollControlIntoView(childControl);
                    }

                    control = control.Parent;
                }
            }
        }

        //
        //  setfocus
        //

        private delegate void SetFocusDelegate(IntPtr control);
        private SetFocusDelegate m_SetFocusDelegater;

        unsafe public void AddFirstMessageSetFocus(IntPtr control)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.SetFocus;

            Parameter HandleParam = new Parameter(this, control);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void SetFocus(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("SetFocus must be the first message");
            }

            // p1  = handle
            IntPtr control = GetParameterIntPtr(ptrMessage, 0);

            //TODO WPF etc
            WF.Control childControl = WF.Control.FromHandle(control);
            if (childControl != null)
            {
                object[] theParameters = { control };
                childControl.Invoke(m_SetFocusDelegater, theParameters);
            }
            CleanUpMessage(ptrMessage);
        }

        private unsafe void SetFocusInternal(IntPtr control)
        {
            WF.Control childControl = WF.Control.FromHandle(control);

            if (childControl != null)
            {
                if (childControl.CanFocus)
                {
                    if (!childControl.ContainsFocus)
                    {
                        childControl.Focus();
                    }
                }
            }
        }

        //
        //  setfocusasync
        //

        private delegate void SetFocusAsyncDelegate(IntPtr control);
        private SetFocusDelegate m_SetFocusAsyncDelegater;

        unsafe public void AddFirstMessageSetFocusAsync(IntPtr control)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.SetFocusAsync;

            Parameter HandleParam = new Parameter(this, control);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void SetFocusAsync(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("SetFocusAsync must be the first message");
            }

            // p1  = handle
            IntPtr control = GetParameterIntPtr(ptrMessage, 0);

            //TODO WPF etc
            WF.Control childControl = WF.Control.FromHandle(control);
            if (childControl != null)
            {
                object[] theParameters = { control };
                childControl.BeginInvoke(m_SetFocusAsyncDelegater, theParameters);
            }
            CleanUpMessage(ptrMessage);
        }

        private unsafe void SetFocusAsyncInternal(IntPtr control)
        {
            WF.Control childControl = WF.Control.FromHandle(control);

            if (childControl != null)
            {
                if (childControl.CanFocus)
                {
                    // We dont want to check if childControl.ContainsFocus for this call unlike SetFocusInternal
                    childControl.Focus();
                }
            }
        }

        //
        //  PeakMessage
        //

        private delegate bool PeakMessageDelegate();
        private PeakMessageDelegate m_PeakMessagDelegater;

        unsafe public void AddFirstMessagePeakMessage(IntPtr handle, int timeoutMs)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.PeakMessage;

            Parameter HandleParam = new Parameter(this, handle);
            Parameter TimeoutParam = new Parameter(this, timeoutMs);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void PeakMessage(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("PeakMessage must be the first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);
            int timeoutMS = GetParameterInt32(ptrMessage, 1);

            int threadId = NM.GetWindowThreadProcessId(handle, out int pid);
            m_AllControls = new List<IntPtr>();
            NM.EnumThreadWindows((uint)threadId, EnumThreadProcedue, IntPtr.Zero);

            Stopwatch timer = Stopwatch.StartNew();
            for (int loop = 0; loop < 2; loop++)
            {
                foreach (IntPtr hWnd in m_AllControls)
                {
                    WF.Control control = WF.Control.FromHandle(hWnd);

                    if (control == null)
                    {
                        //Todo
                    }
                    else
                    {
                        while (true)
                        {
                            bool messageAvailble = false;
                            if (control.IsDisposed)
                            {
                                // Nothing
                            }
                            else if (control.Disposing)
                            {
                                messageAvailble = true; // Don't invoke anything just continue to loop till the control is fully disposed
                            }
                            else if (!control.IsHandleCreated)
                            {
                                // Nothing as to get to here the handle must have existed at some point so it must have been destroyed
                            }
                            else if (control.RecreatingHandle)
                            {
                                messageAvailble = true; // Don't invoke anything just continue to loop till the control has recreated the handle
                            }
                            else if (!control.Enabled)
                            {
                                // Nothing move on to the next window
                            }
                            else
                            {
                                try
                                {
                                    //control.BeginInvoke(m_RefreshControlDelegater, new object[] { control });
                                    IAsyncResult result = control.BeginInvoke(m_PeakMessagDelegater, null);
                                    while (true)
                                    {
                                        int innerLoop = 0;
                                        if (result.IsCompleted)
                                        {
                                            messageAvailble = (bool)control.EndInvoke(result);
                                            break;
                                        }

                                        if (control.IsDisposed)
                                        {
                                            // Nothing
                                            break;
                                        }
                                        else if (control.Disposing)
                                        {
                                            // Don't do anything just continue to loop till the control is fully disposed
                                        }
                                        else if (!control.IsHandleCreated)
                                        {
                                            // Nothing as to get to here the handle must have existed at some point so it must have been destroyed
                                            break;
                                        }
                                        else if (control.RecreatingHandle)
                                        {
                                            messageAvailble = true; // Don't invoke anything just continue to loop till the control has recreated the handle
                                            break;
                                        }
                                        else if (!control.Enabled)
                                        {
                                            // Nothing move on to the next window
                                            break;
                                        }

                                        if (timer.ElapsedMilliseconds > timeoutMS)
                                        {
                                            throw new Exception("Thread failed to have zero messages within timeout");
                                        }

                                        innerLoop++;

                                        if (innerLoop == 100)
                                        {
                                            innerLoop = 0;
                                            Thread.Sleep(15);
                                        }
                                        else
                                        {
                                            Thread.Yield();
                                        }
                                    }
                                }
                                catch (ObjectDisposedException) { }
                                catch (InvalidAsynchronousStateException) { }
                                catch (NullReferenceException) { }
                                catch (InvalidOperationException) { }
                            }

                            if (!messageAvailble)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > timeoutMS)
                            {
                                throw new Exception("Thread failed to have zero messages within timeout");
                            }

                            Thread.Sleep(15);
                        }
                        break;
                    }
                }
            }

            CleanUpMessage(ptrMessage);
        }

        private bool PeakMessageInternal()
        {
            // Might want to include timer here as well at some point ( PM_QS_POSTMESSAGE ) but it might effect things to much
            NativeMessage msg;
            bool messageAvailble = PeekMessage(out msg, IntPtr.Zero, 0, 0, PeekMessageFlags.NoRemove | PeekMessageFlags.NoYield | PeekMessageFlags.PM_QS_INPUT);// | PeekMessageFlags.PM_QS_PAINT);
            return messageAvailble;
        }

        [SuppressUnmanagedCodeSecurity]
        [return: MarshalAs(UnmanagedType.Bool)]
        [DllImport("User32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool PeekMessage(out NativeMessage message, IntPtr handle, uint filterMin, uint filterMax, PeekMessageFlags flags);

        [StructLayout(LayoutKind.Sequential)]
        public struct NativeMessage
        {
            public IntPtr handle;
            public uint msg;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public System.Drawing.Point p;
        }

        private const int QS_KEY = 0x0001,
                            QS_MOUSEMOVE = 0x0002,
                            QS_MOUSEBUTTON = 0x0004,
                            QS_POSTMESSAGE = 0x0008,
                            QS_TIMER = 0x0010,
                            QS_PAINT = 0x0020,
                            QS_SENDMESSAGE = 0x0040,
                            QS_HOTKEY = 0x0080,
                            QS_ALLPOSTMESSAGE = 0x0100,
                            QS_MOUSE = QS_MOUSEMOVE | QS_MOUSEBUTTON,
                            QS_INPUT = QS_MOUSE | QS_KEY,
                            QS_ALLEVENTS = QS_INPUT | QS_POSTMESSAGE | QS_TIMER | QS_PAINT | QS_HOTKEY,
                            QS_ALLINPUT = QS_INPUT | QS_POSTMESSAGE | QS_TIMER | QS_PAINT | QS_HOTKEY | QS_SENDMESSAGE;

        [Flags]
        public enum PeekMessageFlags : uint
        {
            NoRemove = 0,
            Remove = 1,
            NoYield = 2,
            PM_QS_INPUT = unchecked(QS_INPUT << 16),
            PM_QS_PAINT = unchecked(QS_PAINT << 16),
            PM_QS_POSTMESSAGE = unchecked((QS_POSTMESSAGE | QS_HOTKEY | QS_TIMER) << 16),
            PM_QS_SENDMESSAGE = unchecked(QS_SENDMESSAGE << 16),
        }

        //
        //  DumpControl
        //

        private static FieldInfo HashBucketsArrayFieldInfo = typeof(NativeWindow).GetField("hashBuckets", BindingFlags.Static | BindingFlags.NonPublic);
        private static FieldInfo HashBucketHandleFieldInfo = ((Array)HashBucketsArrayFieldInfo.GetValue(null)).GetValue(0).GetType().GetField("handle", BindingFlags.Instance | BindingFlags.Public);
        private static FieldInfo HashBucketGCHandleFieldInfo = ((Array)HashBucketsArrayFieldInfo.GetValue(null)).GetValue(0).GetType().GetField("window", BindingFlags.Instance | BindingFlags.Public);
        private static IntPtr MinusOneIntPtr = new IntPtr(-1);

        unsafe public void AddFirstMessageDumpControl()
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.DumpControl;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void DumpControl(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("DumpControl must be first message");
            }

            CleanUpMessage(ptrMessage);

            StringBuilder dump = new StringBuilder();

            //Get the buckets
            Array hashBuckets = (Array)HashBucketsArrayFieldInfo.GetValue(null);

            for (int item = 0; item < hashBuckets.Length; item++)
            {
                IntPtr handle = (IntPtr)HashBucketHandleFieldInfo.GetValue(hashBuckets.GetValue(item));
                if (handle == IntPtr.Zero || handle == MinusOneIntPtr)
                {
                    //Nothing as not a real window handle
                }
                else
                {
                    Control control = Control.FromHandle(handle);
                    if (control == null)
                    {
                        GCHandle gcHandle = (GCHandle)HashBucketGCHandleFieldInfo.GetValue(hashBuckets.GetValue(item));
                        Type controlType = gcHandle.Target.GetType();
                        dump.Append("TypeName: " + controlType.Name);
                        dump.Append(" TypeNameSpace: " + controlType.Namespace);
                        dump.Append(" ModuleName: " + controlType.Module.Name);
                    }
                    else
                    {
                        dump.Append("Name: " + control.Name);
                        Type controlType = control.GetType();
                        dump.Append(" TypeName: " + controlType.Name);
                        dump.Append(" TypeNameSpace: " + controlType.Namespace);
                        dump.Append(" ModuleName: " + controlType.Module.Name);
                    }

                    //IntPtr parent = NM.GetAncestor(handle, NM.GetAncestorFlags.GetRoot);
                    IntPtr parent = NM.GetAncestor(handle, NM.GetAncestorFlags.GetParent);
                    dump.Append(" Handle: " + handle.ToString());
                    dump.Append(" Parent: " + parent.ToString());
                    dump.AppendLine();
                }
            }

            AddReturnValue(new Parameter(this, dump.ToString()));
        }

        //
        //  GetToolTip
        //

        unsafe public void AddFirstMessageGetToolTip(IntPtr handle)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetToolTip;

            Parameter HandleParam = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetToolTip(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GetToolTip must be first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);

            CleanUpMessage(ptrMessage);

            NM.EnumWindowsProc windowsToGetToolTipsCallback = new NM.EnumWindowsProc(EnumWindowsToGetToolTips);

            m_ToolTipWindows = new List<IntPtr>();
            NM.EnumWindows(windowsToGetToolTipsCallback, IntPtr.Zero);

            IntPtr toolTipHandle = IntPtr.Zero;
            string toolTipTitle = null;
            Rectangle toolTipRectangle = new Rectangle(0, 0, 0, 0);

            foreach (IntPtr hWnd in m_ToolTipWindows)
            {
                NM.ToolInfo info = NM.GetToolInfo(hWnd, TimeOut);

                if (info.hWnd == handle)
                {
                    //we have the tooltip so return infomation about it
                    toolTipHandle = hWnd;
                    toolTipTitle = GetWindowTextViaWindowMessage(toolTipHandle);

                    NM.tagRect windowPosition;
                    NM.tagRect windowSize;
                    NM.GetWindowRect(toolTipHandle, out windowPosition);
                    windowSize = NM.GetClipBox(toolTipHandle);

                    toolTipRectangle = new Rectangle(windowPosition.left, windowPosition.top, windowSize.right, windowSize.bottom);
                    break;
                }
            }

            AddReturnValue(new Parameter(this, toolTipHandle));
            AddReturnValue(new Parameter(this, toolTipTitle));
            AddReturnValue(new Parameter(this, toolTipRectangle.X));
            AddReturnValue(new Parameter(this, toolTipRectangle.Y));
            AddReturnValue(new Parameter(this, toolTipRectangle.Width));
            AddReturnValue(new Parameter(this, toolTipRectangle.Height));
        }

        private List<IntPtr> m_ToolTipWindows;

        private bool EnumWindowsToGetToolTips(IntPtr hWnd, IntPtr lParam)
        {
            uint pid;
            NM.GetWindowThreadProcessId(hWnd, out pid);

            uint currentProcessId = (uint)AUTProcess.Id;
            if (pid == currentProcessId)
            {
                if (NM.IsWindowVisible(hWnd))
                {
                    NM.tagRect WindowSize;
                    NM.GetClientRect(hWnd, out WindowSize);
                    if (WindowSize.right > 0)   //If the window has 0 width then ignore it
                    {
                        string className = NM.GetClassName(hWnd);
                        if (className.Contains("tooltips_class"))
                        {
                            m_ToolTipWindows.Add(hWnd);
                        }
                    }
                }
            }
            return true;
        }
    }
}
