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
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;
using NM = APE.Native.NativeMethods;
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
            m_PeakMessagePaintDelegater = new PeakMessageDelegate(PeakMessagePaintInternal);
            m_PeakMessageKeyDelegater = new PeakMessageDelegate(PeakMessageKeyInternal);
            m_PeakMessageMouseDelegater = new PeakMessageDelegate(PeakMessageMouseInternal);
            m_PeakMessageFocusDelegater = new PeakMessageDelegate(PeakMessageFocusInternal);
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

            Parameter handleParam = new Parameter(this, handle);
            Parameter itemParam = new Parameter(this, (int)item);

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
            childControl.Invoke(m_ScrollControlIntoViewDelegater, theParameters);
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
        private PeakMessageDelegate m_PeakMessagePaintDelegater;
        private PeakMessageDelegate m_PeakMessageKeyDelegater;
        private PeakMessageDelegate m_PeakMessageMouseDelegater;
        private PeakMessageDelegate m_PeakMessageFocusDelegater;

        private const uint WM_PAINT = 0x000F;
        private const uint WM_SETFOCUS = 0x0007;
        private const uint WM_KILLFOCUS = 0x0008;
        private const uint WM_KEYFIRST = 0x0100;
        private const uint WM_KEYLAST = 0x0108;
        private const uint WM_MOUSEFIRST = 0x0200;
        private const uint WM_MOUSELAST = 0x020D;

        unsafe public void AddFirstMessagePeakMessage(IntPtr handle)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.PeakMessage;

            Parameter HandleParam = new Parameter(this, handle);

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
            object[] theParameters = { handle };

            //TODO WPF etc
            WF.Control control = WF.Control.FromHandle(handle);

            if (control != null)
            {
                Stopwatch timer = Stopwatch.StartNew();
                while (true)
                {
                    bool messageAvailble = false;
                    if (!control.IsDisposed)
                    {
                        try
                        {
                            // Might want to include time here as well at some point
                            messageAvailble = (bool)control.Invoke(m_PeakMessageFocusDelegater, null);
                            if (!messageAvailble)
                            {
                                messageAvailble = (bool)control.Invoke(m_PeakMessageKeyDelegater, null);
                                if (!messageAvailble)
                                {
                                    messageAvailble = (bool)control.Invoke(m_PeakMessageMouseDelegater, null);
                                    if (!messageAvailble)
                                    {
                                        messageAvailble = (bool)control.Invoke(m_PeakMessagePaintDelegater, null);
                                    }
                                }
                            }
                        }
                        catch (ObjectDisposedException)
                        {
                        }
                    }
                    if (!messageAvailble)
                    {
                        break;
                    }

                    if (!NM.IsWindowEnabled(handle))
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > m_TimeOut)
                    {
                        throw new Exception("Thread failed to have zero messages within timeout");
                    }

                    Thread.Sleep(50);
                }
            }

            CleanUpMessage(ptrMessage);
        }

        private unsafe bool PeakMessagePaintInternal()
        {
            NativeMessage msg;
            bool messageAvailble = PeekMessage(out msg, IntPtr.Zero, WM_PAINT, WM_PAINT, PeekMessageFlags.NoRemove | PeekMessageFlags.NoYield);
            return messageAvailble;
        }

        private unsafe bool PeakMessageKeyInternal()
        {
            NativeMessage msg;
            bool messageAvailble = PeekMessage(out msg, IntPtr.Zero, WM_KEYFIRST, WM_KEYLAST, PeekMessageFlags.NoRemove | PeekMessageFlags.NoYield);
            return messageAvailble;
        }

        private unsafe bool PeakMessageMouseInternal()
        {
            NativeMessage msg;
            bool messageAvailble = PeekMessage(out msg, IntPtr.Zero, WM_MOUSEFIRST, WM_MOUSELAST, PeekMessageFlags.NoRemove | PeekMessageFlags.NoYield);
            return messageAvailble;
        }

        private unsafe bool PeakMessageFocusInternal()
        {
            NativeMessage msg;
            bool messageAvailble = PeekMessage(out msg, IntPtr.Zero, WM_SETFOCUS, WM_KILLFOCUS, PeekMessageFlags.NoRemove | PeekMessageFlags.NoYield);
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

        [Flags]
        public enum PeekMessageFlags : uint
        {
            NoRemove = 0,
            Remove = 1,
            NoYield = 2,
        }
    }
}
