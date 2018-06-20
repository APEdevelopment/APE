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
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using APE.Capture;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;
using System.Drawing;

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
        }

        /// <summary>
        /// Initialises a new instance of the GUIForm class for a child form
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIForm(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
            //When a form is opened it is animated, so wait for it to finish
            WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);
        }

        /// <summary>
        /// Closes the form by click the 'x' in the top right hand corner of the form
        /// </summary>
        public void Close()
        {
            if (NM.IsWindow(Identity.Handle))
            {
                if (NM.IsWindowVisible(Identity.Handle))
                {
                    if (NM.IsIconic(Identity.Handle))
                    {
                        throw GUI.ApeException("Can not close the window as it is minimised");
                    }

                    GUI.m_APE.AddFirstMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Close);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
                    int Top = GUI.m_APE.GetValueFromMessage();
                    int Left = GUI.m_APE.GetValueFromMessage();
                    int Bottom = GUI.m_APE.GetValueFromMessage();
                    int Right = GUI.m_APE.GetValueFromMessage();

                    // Check the close button is actually displayed
                    if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
                    {
                        throw GUI.ApeException("Can not close the window as the close button is in state '" + State.ToString() + "'");
                    }

                    NM.tagRect WindowRect;
                    NM.GetWindowRect(Identity.Handle, out WindowRect);

                    int X = Left + ((Right - Left) / 2) - WindowRect.left;
                    int Y = Top + ((Bottom - Top) / 2) - WindowRect.top;

                    Input.Block();
                    try
                    {
                        GUI.Log("Close the " + Identity.Description, LogItemType.Action);
                        
                        // OOM
                        //Input.ClickCommon(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y);
                        //Thread.Sleep(100);  //Move to SetFocus?

                        base.SingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                        //Wait for the window to disappear
                        base.WaitForControlToNotBeVisible();
                    }
                    catch when (Input.ResetInputFilter())
                    {
                        // Will never be reached as ResetInputFilter always returns false
                    }
                    finally
                    {
                        Input.Unblock();
                    }
                }
            }
        }

        /// <summary>
        /// Gets if the form is maximised
        /// </summary>
        public bool IsMaximised
        {
            get
            {
                string windowState = FormWindowState();

                if (windowState == "Maximized")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Gets if the form is minimised
        /// </summary>
        public bool IsMinimised
        {
            get
            {
                string windowState = FormWindowState();

                if (windowState == "Minimized")
                {
                    return true;
                }
                else
                {
                    return false;
                }
            }
        }

        /// <summary>
        /// Get the forms window state
        /// </summary>
        /// <returns>The window state</returns>
        private string FormWindowState()
        {
            if (NM.IsWindow(Identity.Handle))
            {
                if (NM.IsWindowVisible(Identity.Handle))
                {
                    if (Identity.TechnologyType == "Windows Forms (WinForms)")
                    {
                        // Get the windows current state
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "WindowState", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string windowState = GUI.m_APE.GetValueFromMessage();
                        return windowState;
                    }
                    else
                    {
                        if (NM.IsIconic(Handle))
                        {
                            return "Minimized";
                        }
                        else if (NM.IsZoomed(Handle))
                        {
                            return "Maximized";
                        }
                        else
                        {
                            return "Normal";
                        }
                    }
                }
                else
                {
                    throw GUI.ApeException("Form is not visible");
                }
            }
            else
            {
                throw GUI.ApeException("Form does not exist");
            }
        }

        /// <summary>
        /// Maximises the form by clicking the maximise button in the top right hand corner of the form
        /// </summary>
        public void Maximise()
        {
            GUI.Log("Maximise the " + Identity.Description, LogItemType.Action);

            string windowState = FormWindowState();
            if (windowState == "Minimized")
            {
                throw GUI.ApeException("Can not maximise the window as it is minimised");
            }

            if (windowState == "Maximized")
            {
                throw GUI.ApeException("Can not maximise the window as it already is");
            }

            GUI.m_APE.AddFirstMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Maximize);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
            int Top = GUI.m_APE.GetValueFromMessage();
            int Left = GUI.m_APE.GetValueFromMessage();
            int Bottom = GUI.m_APE.GetValueFromMessage();
            int Right = GUI.m_APE.GetValueFromMessage();

            // Check the maximise button is actually displayed
            if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
            {
                throw GUI.ApeException("Can not maximise the window as the maximise button is in state '" + State.ToString() + "'");
            }

            NM.tagRect WindowRect;
            NM.GetWindowRect(Identity.Handle, out WindowRect);

            int X = Left + ((Right - Left) / 2) - WindowRect.left;
            int Y = Top + ((Bottom - Top) / 2) - WindowRect.top;

            Input.Block();
            try
            {
                base.SingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                // Wait for the animation to finish
                base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                // Check the window is now maximised
                Stopwatch timer = Stopwatch.StartNew();
                do
                {
                    windowState = FormWindowState();

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        timer.Stop();
                        throw GUI.ApeException("Failed to maximize the window");
                    }

                    Thread.Sleep(15);
                }
                while (windowState != "Maximized");
                timer.Stop();
            }
            catch when (Input.ResetInputFilter())
            {
                // Will never be reached as ResetInputFilter always returns false
            }
            finally
            {
                Input.Unblock();
            }
        }

        /// <summary>
        /// Minimises the form by clicking the minimise button in the top right hand corner of the form
        /// </summary>
        public void Minimise()
        {
            GUI.Log("Minimise the " + Identity.Description, LogItemType.Action);

            string windowState = FormWindowState();
            if (windowState == "Minimized")
            {
                throw GUI.ApeException("Can not minimise the window as it already is");
            }

            GUI.m_APE.AddFirstMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Minimize);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
            int Top = GUI.m_APE.GetValueFromMessage();
            int Left = GUI.m_APE.GetValueFromMessage();
            int Bottom = GUI.m_APE.GetValueFromMessage();
            int Right = GUI.m_APE.GetValueFromMessage();

            // Check the maximise button is actually displayed
            if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
            {
                throw GUI.ApeException("Can not minimise the window as the minimised button is in state '" + State.ToString() + "'");
            }

            NM.tagRect WindowRect;
            NM.GetWindowRect(Identity.Handle, out WindowRect);

            int X = Left + ((Right - Left) / 2) - WindowRect.left;
            int Y = Top + ((Bottom - Top) / 2) - WindowRect.top;

            Input.Block();
            try
            {
                base.SingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                // Wait for the animation to finish
                base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                // Check the window is now minimised
                Stopwatch timer = Stopwatch.StartNew();
                do
                {
                    windowState = FormWindowState();

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        timer.Stop();
                        throw GUI.ApeException("Failed to minimise the window");
                    }

                    Thread.Sleep(15);
                }
                while (windowState != "Minimized");
                timer.Stop();
            }
            catch when (Input.ResetInputFilter())
            {
                // Will never be reached as ResetInputFilter always returns false
            }
            finally
            {
                Input.Unblock();
            }
        }

        /// <summary>
        /// Restores the window back clicking the restore button in the top right corner of the form
        /// Note if the window is minimised then a non-gui method is used to restore the form
        /// </summary>
        public void Restore()
        {
            GUI.Log("Restore the " + Identity.Description, LogItemType.Action);

            string windowState = FormWindowState();
            if (windowState == "Normal")
            {
                throw GUI.ApeException("Can not restore the window as it already is");
            }

            if (windowState == "Minimized")
            {
                Input.Block();
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
                        windowState = FormWindowState();

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            timer.Stop();
                            throw GUI.ApeException("Failed to restore the window");
                        }

                        Thread.Sleep(15);
                    }
                    while (windowState != "Normal");
                    timer.Stop();
                }
                catch when (Input.ResetInputFilter())
                {
                    // Will never be reached as ResetInputFilter always returns false
                }
                finally
                {
                    Input.Unblock();
                }
            }
            else
            {
                GUI.m_APE.AddFirstMessageGetTitleBarItemRectangle(Identity.Handle, NM.TitleBarStateElement.Maximize);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                NM.StateSystem State = (NM.StateSystem)GUI.m_APE.GetValueFromMessage();
                int Top = GUI.m_APE.GetValueFromMessage();
                int Left = GUI.m_APE.GetValueFromMessage();
                int Bottom = GUI.m_APE.GetValueFromMessage();
                int Right = GUI.m_APE.GetValueFromMessage();

                // Check the maximise button is actually displayed
                if (State != NM.StateSystem.STATE_SYSTEM_NORMAL && State != NM.StateSystem.STATE_SYSTEM_PRESSED)
                {
                    throw GUI.ApeException("Can not restore the window as the restore button is in state '" + State.ToString() + "'");
                }

                NM.tagRect WindowRect;
                NM.GetWindowRect(Identity.Handle, out WindowRect);

                int X = Left + ((Right - Left) / 2) - WindowRect.left;
                int Y = Top + ((Bottom - Top) / 2) - WindowRect.top;

                Input.Block();
                try
                {
                    base.SingleClickInternal(X, Y, MouseButton.Left, MouseKeyModifier.None);

                    // Wait for the animation to finish
                    base.WaitForAnimation(Identity.Handle, true, AnimationUtils.WaitForAnimationSource.Form);

                    // Check the window is now minimised
                    Stopwatch timer = Stopwatch.StartNew();
                    do
                    {
                        windowState = FormWindowState();

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            timer.Stop();
                            throw GUI.ApeException("Failed to restore the window");
                        }

                        Thread.Sleep(15);
                    }
                    while (windowState != "Normal");
                    timer.Stop();
                }
                catch when (Input.ResetInputFilter())
                {
                    // Will never be reached as ResetInputFilter always returns false
                }
                finally
                {
                    Input.Unblock();
                }
            }
        }

        /// <summary>
        /// Moves the specified window to the specified location
        /// </summary>
        /// <param name="destinationUpperLeftX">The new location for the left side of the form</param>
        /// <param name="destinationUpperLeftY">The new location for the top of the form</param>
        public void Move(int destinationUpperLeftX, int destinationUpperLeftY)
        {
            //TODO get the titlebar left and click there
            Move(60, 10, destinationUpperLeftX, destinationUpperLeftY, true);
        }

        /// <summary>
        /// Moves the specified window to the specified location
        /// </summary>
        /// <param name="destinationUpperLeftX">The new location for the left side of the form</param>
        /// <param name="destinationUpperLeftY">The new location for the top of the form</param>
        /// <param name="checkMoveCompleted">Whether to check if the form is now in the location specified</param>
        public void Move(int destinationUpperLeftX, int destinationUpperLeftY, bool checkMoveCompleted)
        {
            //TODO get the titlebar left and click there
            Move(60, 10, destinationUpperLeftX, destinationUpperLeftY, checkMoveCompleted);
        }

        /// <summary>
        /// Moves the specified window to the specified location by clicking at the specified point to drag the form
        /// </summary>
        /// <param name="mouseDownX">The x coordinate inside the form to perform a mouse down at</param>
        /// <param name="mouseDownY">The y coordinate inside the form to perform a mouse down at</param>
        /// <param name="destinationUpperLeftX">The new location for the left side of the form</param>
        /// <param name="destinationUpperLeftY">The new location for the top of the form</param>
        public void Move(int mouseDownX, int mouseDownY, int destinationUpperLeftX, int destinationUpperLeftY)
        {
            Move(mouseDownX, mouseDownY, destinationUpperLeftX, destinationUpperLeftY, true);
        }

        /// <summary>
        /// Moves the specified window to the specified location by clicking at the specified point to drag the form
        /// </summary>
        /// <param name="mouseDownX">The x coordinate inside the form to perform a mouse down at</param>
        /// <param name="mouseDownY">The y coordinate inside the form to perform a mouse down at</param>
        /// <param name="destinationUpperLeftX">The new location for the left side of the form</param>
        /// <param name="destinationUpperLeftY">The new location for the top of the form</param>
        /// <param name="checkMoveCompleted">Whether to check if the form is now in the location specified</param>
        public void Move(int mouseDownX, int mouseDownY, int destinationUpperLeftX, int destinationUpperLeftY, bool checkMoveCompleted)
        {
            if (mouseDownX < 3)
            {
                mouseDownX = 3;
            }
            if (mouseDownY < 3)
            {
                mouseDownY = 3;
            }

            Rectangle workingArea = Screen.GetWorkingArea(new Point(destinationUpperLeftX, destinationUpperLeftY));
            if (destinationUpperLeftX < workingArea.Left )
            {
                destinationUpperLeftX = workingArea.Left;
            }
            if (destinationUpperLeftY < workingArea.Top)
            {
                destinationUpperLeftY = workingArea.Top;
            }
            if (destinationUpperLeftX > workingArea.Right - mouseDownX - 3)
            {
                destinationUpperLeftX = workingArea.Right- mouseDownX - 3;
            }
            if (destinationUpperLeftY > workingArea.Bottom - mouseDownY - 3)
            {
                destinationUpperLeftY = workingArea.Bottom - mouseDownY - 3;
            }

            GUI.Log("Move the " + Identity.Description + " window to " + destinationUpperLeftX.ToString() + ", " + destinationUpperLeftY.ToString(), LogItemType.Action);

            NM.tagRect windowRect;
            NM.GetWindowRect(Identity.Handle, out windowRect);

            Input.Block();
            try
            {
                base.MouseDownInternal(mouseDownX, mouseDownY, MouseButton.Left, MouseKeyModifier.None);
                base.MouseUpInternal(destinationUpperLeftX + mouseDownX - windowRect.left, destinationUpperLeftY + mouseDownY - windowRect.top, MouseButton.Left, MouseKeyModifier.None);
                
                if (checkMoveCompleted)
                {
                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        NM.GetWindowRect(Identity.Handle, out windowRect);

                        //if we are in the right place +/- 2 pixel then break the loop
                        if (windowRect.left > destinationUpperLeftX - 2 && windowRect.left < destinationUpperLeftX + 2)
                        {
                            if (windowRect.top > destinationUpperLeftY - 2 && windowRect.top < destinationUpperLeftY + 2)
                            {
                                break;
                            }
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw GUI.ApeException("Failed to move the form");
                        }

                        Thread.Sleep(15);
                    }
                }
            }
            catch when (Input.ResetInputFilter())
            {
                // Will never be reached as ResetInputFilter always returns false
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
