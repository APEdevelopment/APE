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
using System.Reflection;
using System.Diagnostics;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// TD.SandDock.DocumentContainer
    /// </summary>
    public sealed class GUIDocumentContainer : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIDocumentContainer(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        private int DockControls()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetLength", MemberTypes.Method, new Parameter(GUI.m_APE, 0));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int NumberOfItems = GUI.m_APE.GetValueFromMessage();
            return NumberOfItems;
        }

        /// <summary>
        /// Returns the number of controls contained in the document container
        /// </summary>
        /// <returns>The number of controls in the document container</returns>
        public int ItemsCount()
        {
            int Count = 0;
            int NumberOfDockControls = DockControls();

            for (int Tab = 0; Tab < NumberOfDockControls; Tab++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                IntPtr ParentHandle = GUI.m_APE.GetValueFromMessage();

                if (ParentHandle == this.Handle)
                {
                    Count++;
                }
            }

            return Count;
        }

        /// <summary>
        /// Checks if the specified item exists in the document container
        /// </summary>
        /// <param name="item"></param>
        /// <returns>True if it exists otherwise false</returns>
        public bool ItemExists(string item)
        {
            int NumberOfDockControls = DockControls();

            for (int Tab = 0; Tab < NumberOfDockControls; Tab++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string TabText = GUI.m_APE.GetValueFromMessage();
                IntPtr ParentHandle = GUI.m_APE.GetValueFromMessage();

                if (ParentHandle == this.Handle)
                {
                    if (TabText == item)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Selects the specified item in the document container
        /// </summary>
        /// <param name="item">The item in the document container</param>
        public void SingleClickItem(string item)
        {
            SingleClickItem(item, MouseButton.Left);
        }

        /// <summary>
        /// Clicks on the specified item in the document container using the specified mouse button
        /// </summary>
        /// <param name="item">The item in the document container</param>
        /// <param name="button">The mouse button to click with</param>
        public void SingleClickItem(string item, MouseButton button)
        {
            GUI.Log("Select [" + item + "] from " + Identity.Description, LogItemType.Action);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetLength", MemberTypes.Method, new Parameter(GUI.m_APE, 0));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int NumberOfItems = GUI.m_APE.GetValueFromMessage();

            for (int Tab = 0; Tab < NumberOfItems; Tab++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string TabText = GUI.m_APE.GetValueFromMessage();

                if (TabText == item)
                {
                    int rightPadding = RightPadding();
                    Rectangle leftScrollButton = ScrollButtonBounds(ButtonBounds.Left);
                    int leftScrollButtonX = leftScrollButton.X + (leftScrollButton.Width / 2);
                    int leftScrollButtonY = leftScrollButton.Y + (leftScrollButton.Height / 2);
                    Rectangle rightScrollButton = ScrollButtonBounds(ButtonBounds.Right);
                    int rightScrollButtonX = rightScrollButton.X + (rightScrollButton.Width / 2);
                    int rightScrollButtonY = rightScrollButton.Y + (rightScrollButton.Height / 2);

                    int tabX = -1;
                    int tabY = -1;
                    int width = this.Width;

                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        Rectangle tabBounds = GetTabBounds(Tab);
                        tabX = tabBounds.X + (tabBounds.Width / 2);
                        tabY = tabBounds.Y + (tabBounds.Height / 2);

                        if (tabX < 5)
                        {
                            base.SingleClickInternal(leftScrollButtonX, leftScrollButtonY, MouseButton.Left, MouseKeyModifier.None);
                        }
                        else if (tabX > width - rightPadding)
                        {
                            base.SingleClickInternal(rightScrollButtonX, rightScrollButtonY, MouseButton.Left, MouseKeyModifier.None);                            
                        }
                        else
                        {
                            // The centre of the tab can now be clicked
                            break;
                        }

                        //Wait for the scroll to finish happening
                        if (!Input.WaitForInputIdle(Handle, GUI.m_APE.TimeOut))
                        {
                            throw new Exception("Window did not go idle within timeout after scrolling");
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to scroll the document container tab into view");
                        }
                    }

                    base.SingleClickInternal(tabX, tabY, button, MouseKeyModifier.None);

                    //Wait for the active document to be the tab we selected
                    string ActiveTab;
                    timer = Stopwatch.StartNew();
                    do
                    {
                        ActiveTab = ActiveItem();

                        if (ActiveTab == TabText)
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to select document container tab");
                        }

                        Thread.Sleep(50);
                    }
                    while (true);
                    timer.Stop();

                    return;
                }
            }

            throw new Exception("Failed to find DocumentContainer tab");
        }

        private Rectangle GetTabBounds(int tab)
        {
            Rectangle tabBounds = new Rectangle();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, tab));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabBounds", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "X", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store6, "Y", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store7, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store8, "Height", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            tabBounds.X = GUI.m_APE.GetValueFromMessage();
            tabBounds.Y = GUI.m_APE.GetValueFromMessage();
            tabBounds.Width = GUI.m_APE.GetValueFromMessage();
            tabBounds.Height = GUI.m_APE.GetValueFromMessage();

            return tabBounds;
        }

        /// <summary>
        /// The currently selected item in the document container
        /// </summary>
        /// <returns>The text of the currently selected item</returns>
        public string ActiveItem()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ActiveTabbedDocument", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "TabText", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string TabText = GUI.m_APE.GetValueFromMessage();

            return TabText;
        }

        /// <summary>
        /// Removes the specified item from the document container
        /// First selects the item if need be then clicks on the 'x' at the top right of the document container
        /// </summary>
        /// <param name="item"></param>
        public void RemoveItem(string item)
        {
            if (ActiveItem() != item)
            {
                SingleClickItem(item);
            }

            //Get the size of the window
            NM.tagRect ClientRect;
            NM.GetClientRect(Identity.Handle, out ClientRect);

            int InitialItems = ItemsCount();
            int CurrentItems = InitialItems;

            GUI.Log("Remove [" + item + "] from " + Identity.Description, LogItemType.Action);

            //Click 10 pixels in from the right hand side of the window and 10 pixels down
            base.MoveTo(ClientRect.right - 10, 10);
            Thread.Sleep(20);
            base.SingleClickInternal(ClientRect.right - 10, 10, MouseButton.Left, MouseKeyModifier.None);

            //Wait for the number of items to decrease by one
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (InitialItems > 1)
                {
                    CurrentItems = ItemsCount();
                }
                else
                {
                    if (!this.Exists)
                    {
                        CurrentItems = 0;
                    }
                }

                if (CurrentItems == InitialItems - 1)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to remove document container tab");
                }

                Thread.Sleep(50);
            }
            while (true);
            timer.Stop();
        }

        private int RightPadding()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "System.Drawing.Point", MemberTypes.Constructor, new Parameter(GUI.m_APE, 10), new Parameter(GUI.m_APE, 10));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "GetLayoutSystemAt", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store1));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "RightPadding", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int padding = GUI.m_APE.GetValueFromMessage();

            return padding;
        }

        private enum ButtonBounds
        {
            Left = 0,
            Right = 1,
        }
        private Rectangle ScrollButtonBounds(ButtonBounds buttonBounds)
        {
            Rectangle leftScrollButtonBounds = new Rectangle();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "System.Drawing.Point", MemberTypes.Constructor, new Parameter(GUI.m_APE, 10), new Parameter(GUI.m_APE, 10));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "GetLayoutSystemAt", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store1));
            switch (buttonBounds)
            {
                case ButtonBounds.Left:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "LeftScrollButtonBounds", MemberTypes.Property);
                    break;
                case ButtonBounds.Right:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "RightScrollButtonBounds", MemberTypes.Property);
                    break;
                default:
                    throw new Exception("Unsupported button bounds: " + buttonBounds.ToString());
            }
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "X", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Y", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store7, "Height", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            leftScrollButtonBounds.X = GUI.m_APE.GetValueFromMessage();
            leftScrollButtonBounds.Y = GUI.m_APE.GetValueFromMessage();
            leftScrollButtonBounds.Width = GUI.m_APE.GetValueFromMessage();
            leftScrollButtonBounds.Height = GUI.m_APE.GetValueFromMessage();

            return leftScrollButtonBounds;
        }
    }
}
