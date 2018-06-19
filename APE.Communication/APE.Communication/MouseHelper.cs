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
using NM = APE.Native.NativeMethods;
using System.Diagnostics;
using System.Threading;
using System.Runtime.InteropServices;
using System.ComponentModel;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to improve reliability of mouse input to the AUT
    /// </summary>
    public partial class APEIPC
    {
        private bool m_WM_LBUTTONDOWN = false;
        private bool m_WM_LBUTTONUP = false;
        private bool m_WM_LBUTTONDBLCLK = false;
        private bool m_WM_RBUTTONDOWN = false;
        private bool m_WM_RBUTTONUP = false;
        private bool m_WM_RBUTTONDBLCLK = false;
        private bool m_WM_MBUTTONDOWN = false;
        private bool m_WM_MBUTTONUP = false;
        private bool m_WM_MBUTTONDBLCLK = false;

        private NM.HookProc MouseHookProcedure;
        private int m_hMouseHook = 0;
        private IntPtr m_HookWindow;

        private bool m_MouseClicked = false;
        private WF.MouseEventHandler m_MouseClick = null;
        private WF.Control m_MouseClickControl = null;

        /// <summary>
        /// The mouse button to press or release while performing a mouse action
        /// </summary>
        public enum MouseButton : int
        {
            /// <summary>
            /// The left mouse button
            /// </summary>
            Left = 0,
            /// <summary>
            /// The right mouse button
            /// </summary>
            Right = 1,
            /// <summary>
            /// The middle mouse button
            /// </summary>
            Middle = 2,
        }

        /// <summary>
        /// Initialise the mouse helper hooks for use
        /// </summary>
        private void SetupmouseHelperHooks()
        {
            MouseHookProcedure = new NM.HookProc(MouseHookProc);
            m_MouseClick = new WF.MouseEventHandler(MouseClickHandler);
        }
        
        /// <summary>
        /// Resets all the mouse window message variables back to false
        /// </summary>
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

        /// <summary>
        /// Calls into the AUT to add a mouse hook to capture mouse events on the thread belonging to the specified window
        /// </summary>
        /// <param name="handle">A handle to the window of the thread you want to hook</param>
        unsafe public void AddFirstMessageAddMouseHook(IntPtr handle)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->Action = MessageAction.AddMouseHook;

            Parameter handleParameter = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Gets the parameters from the message then adds a mouse hook to the specified thread
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        /// <param name="messageNumber">The message number</param>
        private unsafe void AddMouseHook(Message* ptrMessage, int messageNumber)
        {
            if (messageNumber != 1)
            {
                throw new Exception("AddMouseHook must be the first message");
            }

            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);
            CleanUpMessage(ptrMessage);

            int threadId = NM.GetWindowThreadProcessId(handle, IntPtr.Zero);
            m_HookWindow = handle;

            // Add the mouse hook
            DebugLogging.WriteLog("Adding Mouse hook");
            m_hMouseHook = NM.SetWindowsHookEx(NM.WH_MOUSE, MouseHookProcedure, IntPtr.Zero, threadId);
            if (m_hMouseHook == 0)
            {
                throw new Exception("SetWindowsHookEx failed to add mouse hook");
            }
            DebugLogging.WriteLog("Added Mouse hook");

            ClearMouseState();
        }

        /// <summary>
        /// Calls into the AUT to remove a previously added mouse hook on the thread belonging to the specified window
        /// </summary>
        /// <param name="handle">A handle to the window of the thread you want to remove the hook from</param>
        unsafe public void AddFirstMessageRemoveMouseHook(IntPtr handle)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->Action = MessageAction.RemoveMouseHook;

            Parameter handleParameter = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Gets the parameters from the message then removes the mouse hook on the specified thread
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        /// <param name="messageNumber">The message number</param>
        private unsafe void RemoveMouseHook(Message* ptrMessage, int messageNumber)
        {
            try
            {
                if (messageNumber != 1)
                {
                    throw new Exception("RemoveMouseHook must be the first message");
                }

                IntPtr handle = GetParameterIntPtr(ptrMessage, 0); ;
                CleanUpMessage(ptrMessage);

                // Remove the hook
                DebugLogging.WriteLog("Removing Mouse hook");
                bool returnValue = NM.UnhookWindowsHookEx(m_hMouseHook);
                if (!returnValue)
                {
                    string errorMessage = null;
                    bool ignoreFailedToUnhook;

                    if (Marshal.GetLastWin32Error() == 0)
                    {
                        ignoreFailedToUnhook = true;
                    }
                    else
                    {
                        ignoreFailedToUnhook = false;
                        errorMessage = new Win32Exception(Marshal.GetLastWin32Error()).Message;
                    }

                    if (!ignoreFailedToUnhook)
                    {
                        throw new Exception("UnhookWindowsHookEx Failed for Mouse hook: " + errorMessage);
                    }
                }
            }
            finally
            {
                m_hMouseHook = 0;
                ClearMouseState();
                DebugLogging.WriteLog("Removed Mouse hook");
            }
           
        }

        /// <summary>
        /// Calls into the AUT to wait for a previously hooked thread to recieve the desired mouse messages
        /// </summary>
        /// <param name="button">The mouse button to wait for</param>
        /// <param name="mouseDown">Whether the button should be down or up</param>
        /// <param name="firstClick">Whether the click to wait for is the first click or not</param>
        unsafe public void AddFirstMessageWaitForMouseState(MouseButton button, bool mouseDown, bool firstClick)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->Action = MessageAction.WaitForMouseState;

            Parameter buttonParameter = new Parameter(this, (int)button);
            Parameter mouseDownParameter = new Parameter(this, mouseDown);
            Parameter firstClickParameter = new Parameter(this, firstClick);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }
        
        /// <summary>
        /// Gets the parameters from the message then waits for the desired mouse messages to arrive
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        /// <param name="messageNumber">The message number</param>
        private unsafe void WaitForMouseState(Message* ptrMessage, int messageNumber)
        {
            if (messageNumber != 1)
            {
                throw new Exception("WaitForMouseState must be the first message");
            }

            MouseButton button = (MouseButton)GetParameterInt32(ptrMessage, 0);
            bool mouseDown = GetParameterBoolean(ptrMessage, 1);
            bool firstClick = GetParameterBoolean(ptrMessage, 2);
            CleanUpMessage(ptrMessage);
            Stopwatch timer = Stopwatch.StartNew();

            if (mouseDown)
            {
                DebugLogging.WriteLog("Waiting on " + button.ToString() + " mouse down");
            }
            else
            {
                DebugLogging.WriteLog("Waiting on " + button.ToString() + " mouse up");
            }

            try
            {
                bool done = false;
                while (!done)
                {
                    switch (button)
                    {
                        case MouseButton.Left:
                            if (mouseDown)
                            {
                                if (firstClick)
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
                            if (mouseDown)
                            {
                                if (firstClick)
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
                            if (mouseDown)
                            {
                                if (firstClick)
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
                        Thread.Sleep(15);
                    }

                    if (timer.ElapsedMilliseconds > m_TimeOut)
                    {
                        timer.Stop();
                        if (mouseDown)
                        {
                            throw new Exception("Failed to find " + button.ToString() + " mouse down");
                        }
                        else
                        {
                            throw new Exception("Failed to find " + button.ToString() + " mouse up");
                        }

                    }
                }
            }
            finally
            {
                ClearMouseState();
            }

            DebugLogging.WriteLog("Mouse State done");
        }

        /// <summary>
        /// Callback that is called after the mouse hook has been installed when the thread recieves a mouse message
        /// </summary>
        /// <param name="nCode">A code that the hook procedure uses to determine how to process the message</param>
        /// <param name="wParam">The identifier of the mouse message</param>
        /// <param name="lParam">A pointer to a MOUSEHOOKSTRUCT structure</param>
        /// <returns>Value returned by CallNextHookEx</returns>
        public int MouseHookProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode < 0)
            {
                return NM.CallNextHookEx(m_hMouseHook, nCode, wParam, lParam);
            }
            else
            {
                NM.MouseHookStruct MyMouseHookStruct = (NM.MouseHookStruct)Marshal.PtrToStructure(lParam, typeof(NM.MouseHookStruct));

                //Check if the window or its parent recieved a mouse message (we check for a parent as they have a bad habit of eating some clicks)
                if (MyMouseHookStruct.hwnd == m_HookWindow || NM.IsChild(MyMouseHookStruct.hwnd, m_HookWindow))
                {
                    if (nCode == NM.HC_ACTION)
                    {
                        switch (wParam.ToInt32())
                        {
                            case NM.WM_LBUTTONDOWN:
                                DebugLogging.WriteLog("Left Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDOWN = true;
                                break;
                            case NM.WM_NCLBUTTONDOWN:
                                DebugLogging.WriteLog("NCLeft Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDOWN = true;
                                break;
                            case NM.WM_LBUTTONUP:
                                DebugLogging.WriteLog("Left Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONUP = true;
                                break;
                            case NM.WM_NCLBUTTONUP:
                                DebugLogging.WriteLog("NCLeft Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONUP = true;
                                break;
                            case NM.WM_LBUTTONDBLCLK:
                                DebugLogging.WriteLog("Left Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDBLCLK = true;
                                break;
                            case NM.WM_NCLBUTTONDBLCLK:
                                DebugLogging.WriteLog("NCLeft Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_LBUTTONDBLCLK = true;
                                break;
                            case NM.WM_RBUTTONDOWN:
                                DebugLogging.WriteLog("Right Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDOWN = true;
                                break;
                            case NM.WM_NCRBUTTONDOWN:
                                DebugLogging.WriteLog("NCRight Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDOWN = true;
                                break;
                            case NM.WM_RBUTTONUP:
                                DebugLogging.WriteLog("Right Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONUP = true;
                                break;
                            case NM.WM_NCRBUTTONUP:
                                DebugLogging.WriteLog("NCRight Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONUP = true;
                                break;
                            case NM.WM_RBUTTONDBLCLK:
                                DebugLogging.WriteLog("Right Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDBLCLK = true;
                                break;
                            case NM.WM_NCRBUTTONDBLCLK:
                                DebugLogging.WriteLog("NCRight Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_RBUTTONDBLCLK = true;
                                break;
                            case NM.WM_MBUTTONDOWN:
                                DebugLogging.WriteLog("Middle Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDOWN = true;
                                break;
                            case NM.WM_NCMBUTTONDOWN:
                                DebugLogging.WriteLog("NCMiddle Down " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDOWN = true;
                                break;
                            case NM.WM_MBUTTONUP:
                                DebugLogging.WriteLog("Middle Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                break;
                            case NM.WM_NCMBUTTONUP:
                                DebugLogging.WriteLog("NCMiddle Up " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONUP = true;
                                break;
                            case NM.WM_MBUTTONDBLCLK:
                                DebugLogging.WriteLog("Middle Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDBLCLK = true;
                                break;
                            case NM.WM_NCMBUTTONDBLCLK:
                                DebugLogging.WriteLog("NCMiddle Double " + nCode.ToString() + " " + MyMouseHookStruct.hwnd.ToString() + " " + MyMouseHookStruct.pt.x.ToString() + " x " + MyMouseHookStruct.pt.y.ToString() + " " + MyMouseHookStruct.wHitTestCode.ToString());
                                m_WM_MBUTTONDBLCLK = true;
                                break;
                        }
                    }
                }
                return NM.CallNextHookEx(m_hMouseHook, nCode, wParam, lParam);
            }
        }

        //
        //  AddMouseClickHandler
        //

        /// <summary>
        /// Calls into the AUT to add a mouse click handler to the specific control
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the control object</param>
        unsafe public void AddQueryMessageAddMouseClickHandler(DataStores sourceStore)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the dictionary before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->Action = MessageAction.AddMouseClickHandler;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Adds a mouse click handler to the control
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void AddMouseClickHandler(Message* ptrMessage)
        {
            m_MouseClickControl = (WF.Control)GetObjectFromDatastore(ptrMessage->SourceStore);

            m_MouseClicked = false;

            if (m_MouseClickControl != null)
            {
                m_MouseClickControl.MouseClick += m_MouseClick;
                m_MouseClickControl.MouseUp += m_MouseClick;
                m_MouseClickControl.MouseDown += m_MouseClick;
            }

            CleanUpMessage(ptrMessage);
        }

        private void MouseClickHandler(object sender, WF.MouseEventArgs e)
        {
            m_MouseClicked = true;
        }

        //
        //  WaitForAndRemoveMouseClickHandler
        //

        /// <summary>
        /// Calls into the AUT to wait for then remove the mouse click handler
        /// </summary>
        unsafe public void AddFirstMessageWaitForAndRemoveMouseClickHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.WaitForAndRemoveMouseClickHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Waits for the handler to set the mouse click variable to true then removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void WaitForAndRemoveMouseClickHandler(Message* ptrMessage)
        {
            if (m_MouseClickControl != null)
            {
                try
                {
                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (m_MouseClicked || m_MouseClickControl.Disposing || m_MouseClickControl.IsDisposed)
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > m_TimeOut)
                        {
                            throw new Exception("Failed to detect mouse event");
                        }

                        Thread.Sleep(15);
                    }
                }
                finally
                {
                    m_MouseClickControl.MouseClick -= m_MouseClick;
                    m_MouseClickControl.MouseUp -= m_MouseClick;
                    m_MouseClickControl.MouseDown -= m_MouseClick;
                    m_MouseClickControl = null;
                }
            }

            CleanUpMessage(ptrMessage);
        }

        //
        //  RemoveMouseClickHandler
        //

        /// <summary>
        /// Calls into the AUT to remove the mouse click handler
        /// </summary>
        unsafe public void AddFirstMessageRemoveMouseClickHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.RemoveMouseClickHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void RemoveMouseClickHandler(Message* ptrMessage)
        {
            if (m_MouseClickControl != null)
            {
                m_MouseClickControl.MouseClick -= m_MouseClick;
                m_MouseClickControl.MouseUp -= m_MouseClick;
                m_MouseClickControl.MouseDown -= m_MouseClick;
                m_MouseClickControl = null;
            }

            CleanUpMessage(ptrMessage);
        }
    }
}
