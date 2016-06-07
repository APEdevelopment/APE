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
using System.Security.Principal;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms
    /// </summary>
    /// 
    public sealed class GUIForm : GUIFocusableObject, IGUIForm
    {
        /// <summary>
        /// Constructor used for form controls
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIForm(string descriptionOfControl, params Identifier[] identParams)
            : base(descriptionOfControl, identParams)
        {
            //When a form is opened it is animated, so wait for it to finish
            WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

            Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);
        }

        public void Close()
        {
            if (NM.IsWindow(Identity.Handle))
            {
                if (NM.IsWindowVisible(Identity.Handle))
                {
                    if (NM.IsIconic(Identity.Handle))
                    {
                        throw new Exception("Can not close the window as it is minimised");
                    }

                    GUI.m_APE.AddMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Close);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
                    int Top = GUI.m_APE.GetValueFromMessage();
                    int Left = GUI.m_APE.GetValueFromMessage();
                    int Bottom = GUI.m_APE.GetValueFromMessage();
                    int Right = GUI.m_APE.GetValueFromMessage();

                    // Check the close button is actually displayed
                    if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
                    {
                        throw new Exception("Can not close the window as the close button is in state '" + State.ToString() + "'");
                    }

                    NM.RECT WindowRect;
                    NM.GetWindowRect(Identity.Handle, out WindowRect);

                    int X = Left + ((Right - Left) / 2) - WindowRect.Left;
                    int Y = Top + ((Bottom - Top) / 2) - WindowRect.Top;

                    Input.Block(Identity.ParentHandle, Identity.Handle);
                    try
                    {
                        GUI.Log("Close the " + m_DescriptionOfControl, LogItemTypeEnum.Action);
                        base.MouseSingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                        //Wait for the window to disappear
                        Stopwatch timer = Stopwatch.StartNew();
                        do
                        {
                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw new Exception("Failed to close the window");
                            }

                            if (!NM.IsWindowVisible(Identity.Handle))
                            {
                                break;
                            }

                            Thread.Sleep(15);
                        }
                        while (true);
                        timer.Stop();
                    }
                    finally
                    {
                        Input.Unblock();
                    }
                }
            }
        }

        public void Maximise()
        {
            if (NM.IsWindow(Identity.Handle))
            {
                if (NM.IsWindowVisible(Identity.Handle))
                {
                    GUI.Log("Maximise the " + m_DescriptionOfControl, LogItemTypeEnum.Action);

                    // Get the windows current state
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string WindowState = GUI.m_APE.GetValueFromMessage();

                    if (WindowState == "Minimized")
                    {
                        throw new Exception("Can not maximise the window as it is minimised");
                    }

                    if (WindowState == "Maximized")
                    {
                        throw new Exception("Can not maximise the window as it already is");
                    }

                    GUI.m_APE.AddMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Maximize);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
                    int Top = GUI.m_APE.GetValueFromMessage();
                    int Left = GUI.m_APE.GetValueFromMessage();
                    int Bottom = GUI.m_APE.GetValueFromMessage();
                    int Right = GUI.m_APE.GetValueFromMessage();

                    // Check the maximise button is actually displayed
                    if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
                    {
                        throw new Exception("Can not maximise the window as the maximise button is in state '" + State.ToString() + "'");
                    }

                    NM.RECT WindowRect;
                    NM.GetWindowRect(Identity.Handle, out WindowRect);

                    int X = Left + ((Right - Left) / 2) - WindowRect.Left;
                    int Y = Top + ((Bottom - Top) / 2) - WindowRect.Top;

                    Input.Block(Identity.ParentHandle, Identity.Handle);
                    try
                    {
                        base.MouseSingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                        // Wait for the animation to finish
                        base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                        // Check the window is now maximised
                        Stopwatch timer = Stopwatch.StartNew();
                        do
                        {
                            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            WindowState = GUI.m_APE.GetValueFromMessage();

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                timer.Stop();
                                throw new Exception("Failed to maximize the window");
                            }

                            Thread.Sleep(15);
                        }
                        while (WindowState != "Maximized");
                        timer.Stop();
                    }
                    finally
                    {
                        Input.Unblock();
                    }
                }
            }
        }

        public void Minimise()
        {
            if (NM.IsWindow(Identity.Handle))
            {
                if (NM.IsWindowVisible(Identity.Handle))
                {
                    GUI.Log("Minimise the " + m_DescriptionOfControl, LogItemTypeEnum.Action);

                    // Get the windows current state
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string WindowState = GUI.m_APE.GetValueFromMessage();

                    if (WindowState == "Minimized")
                    {
                        throw new Exception("Can not minimise the window as it already is");
                    }

                    GUI.m_APE.AddMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Minimize);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
                    int Top = GUI.m_APE.GetValueFromMessage();
                    int Left = GUI.m_APE.GetValueFromMessage();
                    int Bottom = GUI.m_APE.GetValueFromMessage();
                    int Right = GUI.m_APE.GetValueFromMessage();

                    // Check the maximise button is actually displayed
                    if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
                    {
                        throw new Exception("Can not minimise the window as the minimised button is in state '" + State.ToString() + "'");
                    }

                    NM.RECT WindowRect;
                    NM.GetWindowRect(Identity.Handle, out WindowRect);

                    int X = Left + ((Right - Left) / 2) - WindowRect.Left;
                    int Y = Top + ((Bottom - Top) / 2) - WindowRect.Top;

                    Input.Block(Identity.ParentHandle, Identity.Handle);
                    try
                    {
                        base.MouseSingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                        // Wait for the animation to finish
                        base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                        // Check the window is now minimised
                        Stopwatch timer = Stopwatch.StartNew();
                        do
                        {
                            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            WindowState = GUI.m_APE.GetValueFromMessage();

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                timer.Stop();
                                throw new Exception("Failed to minimise the window");
                            }

                            Thread.Sleep(15);
                        }
                        while (WindowState != "Minimized");
                        timer.Stop();
                    }
                    finally
                    {
                        Input.Unblock();
                    }
                }
            }
        }

        public void Restore()
        {
            if (NM.IsWindow(Identity.Handle))
            {
                if (NM.IsWindowVisible(Identity.Handle))
                {
                    GUI.Log("Restore the " + m_DescriptionOfControl, LogItemTypeEnum.Action);

                    // Get the windows current state
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string WindowState = GUI.m_APE.GetValueFromMessage();

                    if (WindowState == "Normal")
                    {
                        throw new Exception("Can not restore the window as it already is");
                    }

                    if (WindowState == "Minimized")
                    {
                        Input.Block(Identity.ParentHandle, Identity.Handle);
                        try
                        {
                            NM.ShowWindow(Identity.Handle, NM.SW_RESTORE);

                            //Bring it to the front (can sometimes get restored to the background using the api)
                            SetFocus();

                            // Wait for the animation to finish
                            base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                            // Check the window is now normal
                            Stopwatch timer = Stopwatch.StartNew();
                            do
                            {
                                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                                GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                                WindowState = GUI.m_APE.GetValueFromMessage();

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    timer.Stop();
                                    throw new Exception("Failed to restore the window");
                                }

                                Thread.Sleep(15);
                            }
                            while (WindowState != "Normal");
                            timer.Stop();
                        }
                        finally
                        {
                            Input.Unblock();
                        }
                    }
                    else
                    {
                        GUI.m_APE.AddMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Maximize);
                        GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                        GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
                        int Top = GUI.m_APE.GetValueFromMessage();
                        int Left = GUI.m_APE.GetValueFromMessage();
                        int Bottom = GUI.m_APE.GetValueFromMessage();
                        int Right = GUI.m_APE.GetValueFromMessage();

                        // Check the maximise button is actually displayed
                        if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
                        {
                            throw new Exception("Can not restore the window as the restore button is in state '" + State.ToString() + "'");
                        }

                        NM.RECT WindowRect;
                        NM.GetWindowRect(Identity.Handle, out WindowRect);

                        int X = Left + ((Right - Left) / 2) - WindowRect.Left;
                        int Y = Top + ((Bottom - Top) / 2) - WindowRect.Top;

                        Input.Block(Identity.ParentHandle, Identity.Handle);
                        try
                        {
                            base.MouseSingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                            // Wait for the animation to finish
                            base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                            // Check the window is now minimised
                            Stopwatch timer = Stopwatch.StartNew();
                            do
                            {
                                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                                GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                                WindowState = GUI.m_APE.GetValueFromMessage();

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    timer.Stop();
                                    throw new Exception("Failed to restore the window");
                                }

                                Thread.Sleep(15);
                            }
                            while (WindowState != "Normal");
                            timer.Stop();
                        }
                        finally
                        {
                            Input.Unblock();
                        }
                    }
                }
            }
        }

        public void Move(int DestinationUpperLeftX, int DestinationUpperLeftY)
        {
            //TODO get the titlebar left and click there
            Move(33, 10, DestinationUpperLeftX, DestinationUpperLeftY);
        }

        public void Move(int MouseDownX, int MouseDownY, int DestinationUpperLeftX, int DestinationUpperLeftY)
        {
            GUI.Log("Move the " + m_DescriptionOfControl + " window to " + DestinationUpperLeftX.ToString() + ", " + DestinationUpperLeftY.ToString(), LogItemTypeEnum.Action);

            NM.RECT WindowRect;
            NM.GetWindowRect(Identity.Handle, out WindowRect);

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                base.MouseDownInternal(MouseDownX, MouseDownY, MouseButton.Left, MouseKeyModifier.None);
                base.MouseUpInternal(DestinationUpperLeftX + MouseDownX - WindowRect.Left, DestinationUpperLeftY + MouseDownY - WindowRect.Top, MouseButton.Left, MouseKeyModifier.None);
            }
            finally
            {
                Input.Unblock();
            }
        }

        public void Scroll()
        {
            //TODO
        }

        public void Resize()
        {
            //TODO
        }
    }
}
