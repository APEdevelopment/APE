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
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using APE.Communication;
using System.Windows.Forms;
using APE.Native;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    internal static class Input
    {
        public static int m_Count = 0;

        public static void SendKeys(IntPtr Handle, string text)
        {
            if (!WaitForInputIdle(Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception("Window did not go idle within timeout");
            }
            TimerResolution.SetMaxTimerResolution();
            System.Windows.Forms.SendKeys.SendWait(text);
            TimerResolution.UnsetMaxTimerResolution();
        }

        public static void MouseSingleClick(IntPtr ParentHandle, IntPtr Handle, int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            if (!WaitForInputIdle(Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception("Window did not go idle within timeout");
            }

            uint DoubleClickTimer = (uint)SystemInformation.DoubleClickTime;
            Block(ParentHandle, Handle);
            try
            {
                TimerResolution.SetMaxTimerResolution();
                NM.SetDoubleClickTime(1);
    
                ClickCommon(ParentHandle, Handle, X, Y);
                
                GUI.m_APE.AddMessageAddMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));
                
                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                // Some controls don't like it if the mouse is released too quick (For instance Listview group selecting)
                Thread.Sleep(32);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));
                
                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, true);
                GUI.m_APE.AddMessageRemoveMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            }
            catch
            {
                NM.SetDoubleClickTime(DoubleClickTimer);        //Reset double click timer
                Reset();    //Reset the mouse blocking
                throw;
            }
            finally
            {
                TimerResolution.UnsetMaxTimerResolution();
                NM.SetDoubleClickTime(DoubleClickTimer);
                Unblock();
            }
        }

        public static void MouseDoubleClick(IntPtr ParentHandle, IntPtr Handle, int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            if (!WaitForInputIdle(Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception("Window did not go idle within timeout");
            }
            
            Block(ParentHandle, Handle);
            try
            {
                TimerResolution.SetMaxTimerResolution();

                ClickCommon(ParentHandle, Handle, X, Y);

                GUI.m_APE.AddMessageAddMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));
                
                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));
                
                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));
                
                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, false);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));
                
                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, false);
                GUI.m_APE.AddMessageRemoveMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            }
            catch
            {
                Reset();    //Reset the mouse blocking
                throw;
            }
            finally
            {
                TimerResolution.UnsetMaxTimerResolution();
                Unblock();
            }
        }

        public static void MouseTripleClick(IntPtr ParentHandle, IntPtr Handle, int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            if (!WaitForInputIdle(Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception("Window did not go idle within timeout");
            }
            
            Block(ParentHandle, Handle);
            try
            {
                TimerResolution.SetMaxTimerResolution();

                ClickCommon(ParentHandle, Handle, X, Y);

                GUI.m_APE.AddMessageAddMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, false);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, false);
                GUI.m_APE.AddMessageRemoveMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            }
            catch
            {
                Reset();    //Reset the mouse blocking
                throw;
            }
            finally
            {
                TimerResolution.UnsetMaxTimerResolution();
                Unblock();
            }
        }

        public static void MouseDown(IntPtr ParentHandle, IntPtr Handle, int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            if (!WaitForInputIdle(Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception("Window did not go idle within timeout");
            }

            Block(ParentHandle, Handle);
            try
            {
                TimerResolution.SetMaxTimerResolution();

                ClickCommon(ParentHandle, Handle, X, Y);

                GUI.m_APE.AddMessageAddMouseHook(Handle);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, true, false, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, true, true);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            }
            catch
            {
                Reset();    //Reset the mouse blocking
                throw;
            }
            finally
            {
                TimerResolution.UnsetMaxTimerResolution();
                Unblock();
            }
        }

        public static void MouseUp(IntPtr ParentHandle, IntPtr Handle, int x, int y, MouseButton Button, MouseKeyModifier Keys)
        {
            Block(ParentHandle, Handle);
            try
            {
                NM.tagPoint thePoint = new NM.tagPoint();
                thePoint.x = x;
                thePoint.y = y;
                IntPtr TopLevelHandle = NM.ChildWindowFromPoint(NM.GetDesktopWindow(), thePoint);

                NM.tagRect WindowSize;
                NM.GetWindowRect(Handle, out WindowSize);

                thePoint.x = x + WindowSize.left;
                thePoint.y = y + WindowSize.top;
                IntPtr ChildHandle = NM.WindowFromPoint(thePoint);

                if (!WaitForInputIdle(ChildHandle, GUI.m_APE.TimeOut))
                {
                    throw new Exception("Window did not go idle within timeout");
                }

                IntPtr ActualParent;

                if (ParentHandle == IntPtr.Zero)
                {
                    ActualParent = Handle;
                }
                else
                {
                    ActualParent = ParentHandle;
                }

                TimerResolution.SetMaxTimerResolution();

                //TODO this looks wrong should use clickcommon only for this
                if (ChildHandle == ActualParent)
                {
                    ClickCommon(ParentHandle, Handle, x, y);
                }
                else
                {
                    GUI.m_APE.MouseMove(Handle, x, y, false);
                }

                if (Handle == ActualParent)
                {
                    GUI.m_APE.AddMessageAddMouseHook(Handle);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                }

                GUI.m_APE.MouseClick((APEIPC.MouseButton)Button, false, true, 1, Keys.HasFlag(MouseKeyModifier.Control), Keys.HasFlag(MouseKeyModifier.Shift));

                if (Handle == ActualParent)
                {
                    GUI.m_APE.AddMessageWaitForMouseState((APEIPC.MouseButton)Button, false, true);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                }
            }
            catch
            {
                Reset();    //Reset the mouse blocking
                throw;
            }
            finally
            {
                TimerResolution.UnsetMaxTimerResolution();
                Unblock();
            }
        }

        public static bool ActiveWindow(IntPtr Parent)
        {
            NM.GUITHREADINFO CurrentGuiInfo = new NM.GUITHREADINFO();
            CurrentGuiInfo.cbSize = Marshal.SizeOf(CurrentGuiInfo);
            NM.GetGUIThreadInfo(0, ref CurrentGuiInfo);

            if (CurrentGuiInfo.hwndActive == Parent)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static bool HasFocus(IntPtr Parent, IntPtr Control)
        {
            NM.GUITHREADINFO CurrentGuiInfo = new NM.GUITHREADINFO();
            CurrentGuiInfo.cbSize = Marshal.SizeOf(CurrentGuiInfo);
            NM.GetGUIThreadInfo(0, ref CurrentGuiInfo);

            //Debug.Listeners[0].WriteLine("\t CurrentGuiInfo.hwndActive: " + CurrentGuiInfo.hwndActive.ToString());
            //Debug.Listeners[0].WriteLine("\t CurrentGuiInfo.hwndFocus: " + CurrentGuiInfo.hwndFocus.ToString());

            if (Parent == IntPtr.Zero)
            {
                if (CurrentGuiInfo.hwndActive == Control)
                {
                    return true;
                }
            }
            else
            {
                if (CurrentGuiInfo.hwndActive == Parent)
                {

                    if (CurrentGuiInfo.hwndFocus == Control)
                    {
                        return true;
                    }
                    else if (NM.IsChild(Control, CurrentGuiInfo.hwndFocus))
                    {
                        return true;
                    }
                }
                else
                {
                    if (CurrentGuiInfo.hwndActive == Control)
                    {

                        if (CurrentGuiInfo.hwndFocus == Control)
                        {
                            return true;
                        }
                        else if (NM.IsChild(Control, CurrentGuiInfo.hwndFocus))
                        {
                            return true;
                        }
                    }
                }
            }

            return false;
        }

        public static IntPtr GetFocus()
        {
            NM.GUITHREADINFO CurrentGuiInfo = new NM.GUITHREADINFO();
            CurrentGuiInfo.cbSize = Marshal.SizeOf(CurrentGuiInfo);
            NM.GetGUIThreadInfo(0, ref CurrentGuiInfo);

            return CurrentGuiInfo.hwndFocus;
        }

        public static void SetFocus(IntPtr Parent, IntPtr Control)
        {
            Stopwatch timer = Stopwatch.StartNew();
            IntPtr ActualParent;
            
            if (Parent == IntPtr.Zero)
            {
                ActualParent = Control;
            }
            else
            {
                ActualParent = Parent;
            }

            if (!HasFocus(IntPtr.Zero, ActualParent))
            {
                NM.BringWindowToTop(ActualParent);

                if (!NM.SetForegroundWindow(ActualParent))
                {
                    Debug.WriteLine("Falling back to Hotkey method for parent");
                    // Fall back to the Hotkey (which will have SetForegroundWindow permission)
                    GUI.m_ViewPort.Foreground = ActualParent;

                    while (GUI.m_ViewPort.Foreground != IntPtr.Zero)
                    {
                        // Sendkeys won't work so use keybd_event (TODO could also use SendInput)
                        NM.keybd_event(NM.VK_PAUSE, 0, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        NM.keybd_event(NM.VK_PAUSE, 0, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            timer.Stop();
                            throw new Exception("Viewport SetForegroundWindow appeared to not trigger");
                        }

                        Thread.Sleep(15);
                    }
                }

                while (!HasFocus(IntPtr.Zero, ActualParent))
                {
                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        timer.Stop();
                        throw new Exception("Failed to set focus to the toplevel window");
                    }

                    Thread.Sleep(15);
                }
            }

            if (Parent != IntPtr.Zero)
            {
                if (!HasFocus(Parent, Control))
                {
                    if (!NM.SetForegroundWindow(Control))
                    {
                        Debug.WriteLine("Falling back to Hotkey method for child");
                        // Fall back to the Hotkey (which will have SetForegroundWindow permission)
                        GUI.m_ViewPort.Foreground = Control;

                        // Sendkeys won't work so use keybd_event (TODO could also use SendInput)
                        NM.keybd_event(NM.VK_PAUSE, 0x8f, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        NM.keybd_event(NM.VK_PAUSE, 0x8f, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);

                        while (GUI.m_ViewPort.Foreground != IntPtr.Zero)
                        {
                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                timer.Stop();
                                throw new Exception("SetForegroundWindow failed to set focus to the window");
                            }

                            Thread.Sleep(15);
                        }
                    }

                    while (!HasFocus(Parent, Control))
                    {
                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            timer.Stop();
                            throw new Exception("Failed to set focus to the window");
                        }
                        
                        Thread.Sleep(15);
                    }
                    timer.Stop();
                }
            }

            return;
        }

        public static void ClickCommon(IntPtr Parent, IntPtr Handle, int X, int Y)
        {
            if (NM.IsWindowVisible(Handle))
            {
                if (NM.IsWindowEnabled(Handle))
                {
                    IntPtr ActualParent;

                    if (Parent == IntPtr.Zero)
                    {
                        ActualParent = Handle;
                    }
                    else
                    {
                        ActualParent = Parent;
                    }

                    if (!ActiveWindow(ActualParent))
                    {
                        SetFocus(Parent, Handle);
                    }
                    else
                    {
                        NM.BringWindowToTop(ActualParent);
                    }

                    NM.tagPoint thePoint = GUI.m_APE.MouseMove(Handle, X, Y);

                    IntPtr WindowAtPoint = NM.WindowFromPoint(thePoint);

                    if (WindowAtPoint != Handle)
                    {
                        throw new Exception("Window is obscured");
                    }
                }
                else
                {
                    throw new Exception("Window is not enabled");
                }
            }
            else
            {
                throw new Exception("Window is not visible");
            }
        }

        public static void Block(IntPtr ParentHandle, IntPtr ControlHandle)
        {
            if (m_Count == 0)
            {
                if (GUI.IsElevatedAdmin)
                {
                    bool Return = NM.BlockInput(true);
                    if (!Return)
                    {
                        throw new Exception("Failed to block input");
                    }
                    Thread.Sleep(15);

                    NM.GetKeyState(0);

                    Byte[] state = new byte[256];
                    bool Reset = false;

                    NM.GetKeyboardState(state);

                    //reset the global mouse states
                    if ((state[NM.VK_LBUTTON] & 0x80) != 0)
                    {
                        //TODO use sendinput
                        NM.mouse_event(NM.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
                        Reset = true;
                    }

                    if ((state[NM.VK_MBUTTON] & 0x80) != 0)
                    {
                        //TODO use sendinput
                        NM.mouse_event(NM.MOUSEEVENTF_MIDDLEUP, 0, 0, 0, UIntPtr.Zero);
                        Reset = true;
                    }

                    if ((state[NM.VK_RBUTTON] & 0x80) != 0)
                    {
                        //TODO use sendinput
                        NM.mouse_event(NM.MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
                        Reset = true;
                    }

                    if ((state[NM.VK_XBUTTON1] & 0x80) != 0)
                    {
                        //TODO use sendinput
                        NM.mouse_event(NM.MOUSEEVENTF_XUP, 0, 0, 0, new UIntPtr(1));
                        Reset = true;
                    }

                    if ((state[NM.VK_XBUTTON2] & 0x80) != 0)
                    {
                        //TODO use sendinput
                        NM.mouse_event(NM.MOUSEEVENTF_XUP, 0, 0, 0, new UIntPtr(2));
                        Reset = true;
                    }

                    // 3 is VK_CANCEL
                    if ((state[NM.VK_CANCEL] & 0x80) != 0)
                    {
                        NM.keybd_event(NM.VK_CANCEL, 0, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        NM.keybd_event(NM.VK_CANCEL, 0, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Reset = true;
                    }

                    // 7+ is non-mouse (apart from 3) and 255 seems to break things so we do till 254
                    for (byte vk = 7; vk < unchecked((byte)255); vk++)
                    {
                        if ((state[vk] & 0x80) != 0)
                        {
                            NM.keybd_event(vk, 0, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                            NM.keybd_event(vk, 0, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);
                            Reset = true;
                        }
                    }

                    //check for toggle keys
                    if ((state[NM.VK_CAPITAL] & 0x1) != 0)
                    {
                        //TODO use sendinput
                        NM.keybd_event(NM.VK_CAPITAL, 0, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        NM.keybd_event(NM.VK_CAPITAL, 0, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Reset = true;
                    }

                    if ((state[NM.VK_NUMLOCK] & 0x1) != 0)
                    {
                        //TODO use sendinput
                        NM.keybd_event(NM.VK_NUMLOCK, 0, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        NM.keybd_event(NM.VK_NUMLOCK, 0, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Reset = true;
                    }

                    if ((state[NM.VK_SCROLL] & 0x1) != 0)
                    {
                        //TODO use sendinput
                        NM.keybd_event(NM.VK_SCROLL, 0, NM.KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                        NM.keybd_event(NM.VK_SCROLL, 0, NM.KEYEVENTF_KEYUP, UIntPtr.Zero);
                        Reset = true;
                    }

                    NM.GetKeyState(0);

                    if (Reset)
                    {
                        Thread.Sleep(50);
                    }
                }
            }

            m_Count++;
        }

        public static void Unblock()
        {
            if (m_Count > 0)
            {
                m_Count--;
            }

            if (m_Count == 0)
            {
                Reset();
            }
        }

        public static void Reset()
        {
            if (GUI.IsElevatedAdmin)
            {
                bool Return = NM.BlockInput(false);
            }

            m_Count = 0;
        }

        public static bool WaitForInputIdle(IntPtr handle, uint timeoutMs)
        {
            int processId;
            int threadId = NM.GetWindowThreadProcessId(handle, out processId);

            if (threadId == 0)
            {
                throw new Exception("Failed to get thread for window");
            }

            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (IsThreadIdle(processId, threadId))
                {
                    return true;
                }

                if (timer.ElapsedMilliseconds > timeoutMs)
                {
                    return false;
                }

                Thread.Sleep(15);
            }
            while (true);
        }

        private static bool IsThreadIdle(int processId, int threadId)
        {
            ProcessThreadCollection threadCollection = Process.GetProcessById(processId).Threads;

            for (int i = 0; i < threadCollection.Count; i++)
            {
                if (threadCollection[i].Id == threadId)
                {
                    if (threadCollection[i].ThreadState == System.Diagnostics.ThreadState.Wait)
                    {
                        if (threadCollection[i].WaitReason == ThreadWaitReason.UserRequest)
                        {
                            return true;
                        }
                    }
                    break;
                }
            }
            return false;
        }
    }

    static internal class TimerResolution
    {
        private static bool m_ResolutionSet = false;
        private static uint m_MaximumResolution = 0;

        static TimerResolution()
        {
            uint minimumResolution;
            uint maximumResolution;
            uint currentResolution;

            NM.NtQueryTimerResolution(out minimumResolution, out maximumResolution, out currentResolution);

            m_MaximumResolution = maximumResolution;
        }

        public static void SetMaxTimerResolution()
        {
            if (!m_ResolutionSet)
            {
                uint currentResolution;
                uint result = NM.NtSetTimerResolution(m_MaximumResolution, true, out currentResolution);

                switch (result)
                {
                    case NM.STATUS_SUCCESS:
                        m_ResolutionSet = true;
                        break;
                    default:
                        break;
                }
            }
        }

        public static void UnsetMaxTimerResolution()
        {
            if (m_ResolutionSet)
            {
                uint currentResolution;
                uint result = NM.NtSetTimerResolution(m_MaximumResolution, false, out currentResolution);

                switch (result)
                {
                    case NM.STATUS_SUCCESS:
                        m_ResolutionSet = false;
                        break;
                    default:
                        break;
                }
            }
        }
    }
}
