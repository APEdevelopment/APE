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
using System.Drawing;
using System.Reflection;
using System.Threading;
using APE.Communication;
using NM = APE.Native.NativeMethods;
using System.Text.RegularExpressions;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// TD.SandDock.DockContainer
    /// </summary>
    public class GUIDockContainer : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIDockContainer(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        internal int DockControls()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetLength", MemberTypes.Method, new Parameter(GUI.m_APE, 0));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int numberOfItems = GUI.m_APE.GetValueFromMessage();
            return numberOfItems;
        }

        /// <summary>
        /// Returns the number of controls contained in the container
        /// </summary>
        /// <returns>The number of controls in the container</returns>
        public int ItemsCount()
        {
            int count = 0;

            int numberOfDockControls = DockControls();
            for (int dockControlIndex = 0; dockControlIndex < numberOfDockControls; dockControlIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                dynamic parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle != null && parentHandle == this.Handle)
                {
                    count++;
                }
            }

            return count;
        }

        /// <summary>
        /// Checks if the specified item exists in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <returns>True if it exists otherwise false</returns>
        public bool ItemExists(string itemPattern)
        {
            if (ItemIndex(itemPattern) != -1)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Gets the index of the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <returns>True if it exists otherwise false</returns>
        public int ItemIndex(string itemPattern)
        {
            int numberOfDockControls = DockControls();
            int index = -1;
            for (int dockControlIndex = 0; dockControlIndex < numberOfDockControls; dockControlIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string tabText = GUI.m_APE.GetValueFromMessage();
                dynamic parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle != null && parentHandle == this.Handle)
                {
                    index++;
                    if (Regex.IsMatch(tabText, itemPattern))
                    {
                        return index;
                    }
                }
            }

            return -1;
        }

        /// <summary>
        /// Gets the text of the specified item index in the container
        /// </summary>
        /// <param name="itemIndex">The text of the item.</param>
        /// <returns>True if it exists otherwise false</returns>
        public string ItemText(int itemIndex)
        {
            int numberOfDockControls = DockControls();
            int index = -1;
            for (int dockControlIndex = 0; dockControlIndex < numberOfDockControls; dockControlIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemText = GUI.m_APE.GetValueFromMessage();
                dynamic parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle != null && parentHandle == this.Handle)
                {
                    index++;
                    if (index == itemIndex)
                    {
                        return itemText;
                    }
                }
            }

            throw new Exception("Failed to find the " + Description + " item");
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button
        /// </summary>
        /// <param name="itemIndex">The index of the item.</param>
        public virtual void SingleClickItem(int itemIndex)
        {
            SingleClickItem(itemIndex, MouseButton.Left);
        }

        /// <summary>
        /// Clicks on the specified item in the container using the specified mouse button
        /// </summary>
        /// <param name="itemIndex">The index of the item.</param>
        /// <param name="button">The mouse button to click with</param>
        public virtual void SingleClickItem(int itemIndex, MouseButton button)
        {
            GUI.Log("Single " + button.ToString() + " click on item index " + itemIndex.ToString() + " in the " + Identity.Description, LogItemType.Action);

            string tabText;
            int dockControlIndex = GetDockControlIndex(itemIndex, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }
           
            if (HasTabInternal(dockControlIndex))
            {
                ClickTabInternal(dockControlIndex, tabText, button);
            }
            else if(HasTitlebarInternal(dockControlIndex))
            {
                ClickTitlebar(dockControlIndex, tabText, button);
            }
            else
            {
                ParentForm.SingleClickInternal(ParentForm.Width / 2, 10, button, MouseKeyModifier.None);
            }
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public virtual void SingleClickItem(string itemPattern)
        {
            SingleClickItem(itemPattern, MouseButton.Left);
        }

        /// <summary>
        /// Clicks on the specified item in the container using the specified mouse button
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <param name="button">The mouse button to click with</param>
        public virtual void SingleClickItem(string itemPattern, MouseButton button)
        {
            GUI.Log("Single " + button.ToString() + " click on the " + itemPattern + " item in the " + Identity.Description, LogItemType.Action);

            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }

            if (HasTabInternal(dockControlIndex))
            {
                ClickTabInternal(dockControlIndex, tabText, button);
            }
            else if (HasTitlebarInternal(dockControlIndex))
            {
                ClickTitlebar(dockControlIndex, tabText, button);
            }
            else
            {
                ParentForm.SingleClickInternal(ParentForm.Width / 2, 10, button, MouseKeyModifier.None);
            }
        }

        internal void ClickTitlebar(int dockControlIndex, string itemText, MouseButton button)
        {
            int tabX = -1;
            int tabY = -1;

            Rectangle tabBounds = GetTabBoundsInternal(dockControlIndex);
            tabX = tabBounds.X + (tabBounds.Width / 2);
            tabY = tabBounds.Y + 10;

            base.SingleClickInternal(tabX, tabY, button, MouseKeyModifier.None);

            //Wait for the active item to be the tab we selected
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                string activeItem = ActiveItem();

                if (activeItem == itemText)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    GUI.Log("Failed to select item [" + itemText + "] the currently active item is [" + activeItem + "] in the " + Description, LogItemType.Information);
                    throw new Exception("Failed to select the " + Description + " item");
                }

                Thread.Sleep(50);
            }
        }

        internal void ClickTabInternal(int dockControlIndex, string itemText, MouseButton button)
        {
            int tabX = -1;
            int tabY = -1;

            Rectangle tabBounds = GetTabBoundsInternal(dockControlIndex);
            tabX = tabBounds.X + (tabBounds.Width / 2);
            tabY = tabBounds.Y + (tabBounds.Height / 2);

            base.SingleClickInternal(tabX, tabY, button, MouseKeyModifier.None);

            //Wait for the active item to be the tab we selected
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                string activeItem = ActiveItem();

                if (activeItem == itemText)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    GUI.Log("Failed to select item [" + itemText + "] the currently active item is [" + activeItem + "] in the " + Description, LogItemType.Information);
                    throw new Exception("Failed to select the " + Description + " item");
                }

                Thread.Sleep(50);
            }
        }

        internal void ClickButtonInternal(int dockControlIndex, ButtonToGet button)
        {
            Rectangle buttonBounds = GetButtonBounds(dockControlIndex, button);
            if (buttonBounds.Width == 0 || buttonBounds.Height == 0)
            {
                throw new Exception("Failed to find the " + Description + " " + button.ToString().ToLower() + " button");
            }

            int x = buttonBounds.Left + (buttonBounds.Width / 2);
            int y = buttonBounds.Top + (buttonBounds.Height / 2);

            // Move the mouse then sleep for a bit before clicking
            base.MoveTo(x, y);
            Thread.Sleep(50);
            base.SingleClickInternal(x, y, MouseButton.Left, MouseKeyModifier.None);
        }

        internal bool HasTabInternal(int dockControlIndex)
        {
            Rectangle tabBounds = GetTabBoundsInternal(dockControlIndex);
            
            if (tabBounds.Width == 0 || tabBounds.Height == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal bool HasTitlebarInternal(int dockControlIndex)
        {
            Rectangle titlebarBounds = GetTitlebarBoundsInternal(dockControlIndex);

            if (titlebarBounds.Width == 0 || titlebarBounds.Height == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal Rectangle GetTabBoundsInternal(int dockControlIndex)
        {
            Rectangle tabBounds = new Rectangle();
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
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

        internal Rectangle GetTitlebarBoundsInternal(int dockControlIndex)
        {
            Rectangle titlebarBounds = new Rectangle();
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TitlebarBounds", MemberTypes.Property);
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
            titlebarBounds.X = GUI.m_APE.GetValueFromMessage();
            titlebarBounds.Y = GUI.m_APE.GetValueFromMessage();
            titlebarBounds.Width = GUI.m_APE.GetValueFromMessage();
            titlebarBounds.Height = GUI.m_APE.GetValueFromMessage();
            return titlebarBounds;
        }

        /// <summary>
        /// Checks if the pin button is currently visible for the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <returns>True if visible otherwise false</returns>
        public bool PinButtonVisible(string itemPattern)
        {
            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }

            bool visible = GetButtonVisible(dockControlIndex, ButtonToGet.Pin);
            return visible;
        }

        /// <summary>
        /// Checks if the pin button is currently enabled for the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <returns>True if enabled otherwise false</returns>
        public bool PinButtonEnabled(string itemPattern)
        {
            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }

            bool enabled = GetButtonEnabled(dockControlIndex, ButtonToGet.Pin);
            return enabled;
        }

        /// <summary>
        /// Clicks on the pin button for the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public void PinButtonClick(string itemPattern)
        {
            GUI.Log("Single Left click on the pin button of the " + itemPattern + " item in the " + Identity.Description, LogItemType.Action);

            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }
            if (!PinButtonVisible(itemPattern))
            {
                throw new Exception("Pin button is not visible for the item in the " + Description);
            }
            if (!PinButtonEnabled(itemPattern))
            {
                throw new Exception("Pin button is not enabled for the item in the " + Description);
            }
            ClickButtonInternal(dockControlIndex, ButtonToGet.Pin);
        }

        /// <summary>
        /// Checks if the window button is currently visible for the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <returns>True if visible otherwise false</returns>
        public bool WindowButtonVisible(string itemPattern)
        {
            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }

            bool visible = GetButtonVisible(dockControlIndex, ButtonToGet.Window);
            return visible;
        }

        /// <summary>
        /// Checks if the window button is currently enabled for the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <returns>True if enabled otherwise false</returns>
        public bool WindowButtonEnabled(string itemPattern)
        {
            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }

            bool enabled = GetButtonEnabled(dockControlIndex, ButtonToGet.Window);
            return enabled;
        }

        /// <summary>
        /// Clicks on the window button for the specified item in the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public void WindowButtonClick(string itemPattern)
        {
            GUI.Log("Single Left click on the window button of the " + itemPattern + " item in the " + Identity.Description, LogItemType.Action);

            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }
            if (!WindowButtonVisible(itemPattern))
            {
                throw new Exception("Window button is not visible for the item in the " + Description);
            }
            if (!WindowButtonEnabled(itemPattern))
            {
                throw new Exception("Window button is not enabled for the item in the " + Description);
            }
            ClickButtonInternal(dockControlIndex, ButtonToGet.Window);
        }

        internal int GetDockControlIndex(string itemPattern, out string tabText)
        {
            int numberOfDockControls = DockControls();
            for (int dockControlIndex = 0; dockControlIndex < numberOfDockControls; dockControlIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                tabText = GUI.m_APE.GetValueFromMessage();
                IntPtr parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle == this.Handle)
                {
                    if (Regex.IsMatch(tabText, itemPattern))
                    {
                        return dockControlIndex;
                    }
                }
            }
            tabText = null;
            return -1;
        }

        internal int GetDockControlIndex(int itemIndex, out string tabText)
        {
            int numberOfDockControls = DockControls();
            int index = -1;
            for (int dockControlIndex = 0; dockControlIndex < numberOfDockControls; dockControlIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Parent", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                tabText = GUI.m_APE.GetValueFromMessage();
                dynamic parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle != null && parentHandle == this.Handle)
                {
                    index++;
                    if (index == itemIndex)
                    {
                        return dockControlIndex;
                    }
                }
            }
            tabText = null;
            return -1;
        }

        internal enum ButtonToGet
        {
            ScrollLeft,
            ScrollRight,
            Menu,
            Close,
            Window,
            Pin,
        }

        internal bool GetButtonEnabled(int dockControlIndex, ButtonToGet button)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "LayoutSystem", MemberTypes.Property);
            switch (button)
            {
                case ButtonToGet.ScrollLeft:
                case ButtonToGet.ScrollRight:
                    if (this is GUIDocumentContainer)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                case ButtonToGet.Menu:
                    if (!(this is GUIDocumentContainer))
                    {
                        return false;
                    }
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "MenuButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "enabled", MemberTypes.Field);
                    break;
                case ButtonToGet.Close:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "CloseButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "enabled", MemberTypes.Field);
                    break;
                case ButtonToGet.Window:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "WindowButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "enabled", MemberTypes.Field);
                    break;
                case ButtonToGet.Pin:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "PinButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "enabled", MemberTypes.Field);
                    break;
                default:
                    throw new Exception("Unsupported button: " + button.ToString());
            }
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool enabled = GUI.m_APE.GetValueFromMessage();
            return enabled;
        }

        internal bool GetButtonVisible(int dockControlIndex, ButtonToGet button)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "LayoutSystem", MemberTypes.Property);
            switch (button)
            {
                case ButtonToGet.ScrollLeft:
                case ButtonToGet.ScrollRight:
                    if (this is GUIDocumentContainer)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                case ButtonToGet.Menu:
                    if (!(this is GUIDocumentContainer))
                    {
                        return false;
                    }
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "MenuButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "visible", MemberTypes.Field);
                    break;
                case ButtonToGet.Close:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "CloseButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "visible", MemberTypes.Field);
                    break;
                case ButtonToGet.Window:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "WindowButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "visible", MemberTypes.Field);
                    break;
                case ButtonToGet.Pin:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "PinButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "visible", MemberTypes.Field);
                    break;
                default:
                    throw new Exception("Unsupported button: " + button.ToString());
            }
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool visible = GUI.m_APE.GetValueFromMessage();
            return visible;
        }

        internal Rectangle GetButtonBounds(int dockControlIndex, ButtonToGet button)
        {
            Rectangle buttonBounds = new Rectangle();
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "LayoutSystem", MemberTypes.Property);
            switch (button)
            {
                case ButtonToGet.ScrollLeft:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store4, "LeftScrollButtonBounds", MemberTypes.Property);
                    break;
                case ButtonToGet.ScrollRight:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store4, "RightScrollButtonBounds", MemberTypes.Property);
                    break;
                case ButtonToGet.Menu:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "MenuButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                case ButtonToGet.Close:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "CloseButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                case ButtonToGet.Window:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "WindowButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                case ButtonToGet.Pin:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store3, "PinButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                default:
                    throw new Exception("Unsupported button: " + button.ToString());
            }
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
            buttonBounds.X = GUI.m_APE.GetValueFromMessage();
            buttonBounds.Y = GUI.m_APE.GetValueFromMessage();
            buttonBounds.Width = GUI.m_APE.GetValueFromMessage();
            buttonBounds.Height = GUI.m_APE.GetValueFromMessage();
            return buttonBounds;
        }

        /// <summary>
        /// The last active item in all dock containers
        /// </summary>
        /// <returns>The text of the currently selected item</returns>
        public string ActiveItem()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "FindMostRecentlyUsedWindow", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "TabText", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string tabText = GUI.m_APE.GetValueFromMessage();
            return tabText;
        }

        /// <summary>
        /// Removes the specified item from the container
        /// First selects the item if need be then clicks on the 'x' at the top right of the container
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public void RemoveItem(string itemPattern)
        {
            string activeItem = ActiveItem();
            if (!Regex.IsMatch(activeItem, itemPattern))
            {
                SingleClickItem(itemPattern);
            }

            int initialItems = ItemsCount();
            int currentItems = initialItems;

            GUI.Log("Remove the [" + itemPattern + "] item from the " + Description, LogItemType.Action);

            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw new Exception("Failed to find the " + Description + " item");
            }

            ClickButtonInternal(dockControlIndex, ButtonToGet.Close);

            //Wait for the number of items to decrease by one
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (initialItems > 1)
                {
                    currentItems = ItemsCount();
                }
                else
                {
                    if (!this.Exists)
                    {
                        currentItems = 0;
                    }
                }

                if (currentItems == initialItems - 1)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to remove the " + Description + " item");
                }

                Thread.Sleep(50);
            }
        }
    }
}
