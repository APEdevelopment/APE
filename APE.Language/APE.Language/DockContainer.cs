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

            throw new Exception("Failed to find " + Description + " tab");
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
            GUI.Log("Single " + button.ToString() + " click on item index " + itemIndex.ToString() + " in the" + Identity.Description, LogItemType.Action);

            int numberOfDockControls = DockControls();
            int count = -1;
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
                    count++;

                    if (count == itemIndex)
                    {
                        if (HasTab(dockControlIndex))
                        {
                            ClickTab(dockControlIndex, tabText, button);
                        }
                        else if(HasTitlebar(dockControlIndex))
                        {
                            ClickTitlebar(dockControlIndex, tabText, button);
                        }
                        else
                        {
                            ParentForm.SingleClickInternal(ParentForm.Width / 2, 10, button, MouseKeyModifier.None);
                        }
                        return;
                    }
                }
            }

            throw new Exception("Failed to find " + Description + " tab");
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
            GUI.Log("Single " + button.ToString() + " click on the " + itemPattern + " item in the" + Identity.Description, LogItemType.Action);
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
                string tabText = GUI.m_APE.GetValueFromMessage();
                dynamic parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle != null && parentHandle == this.Handle)
                {
                    if (Regex.IsMatch(tabText, itemPattern))
                    {
                        if (HasTab(dockControlIndex))
                        {
                            ClickTab(dockControlIndex, tabText, button);
                        }
                        else if (HasTitlebar(dockControlIndex))
                        {
                            ClickTitlebar(dockControlIndex, tabText, button);
                        }
                        else
                        {
                            ParentForm.SingleClickInternal(ParentForm.Width / 2, 10, button, MouseKeyModifier.None);
                        }
                        return;
                    }
                }
            }

            throw new Exception("Failed to find " + Description + " tab");
        }

        internal void ClickTitlebar(int dockControlIndex, string itemText, MouseButton button)
        {
            int tabX = -1;
            int tabY = -1;

            Rectangle tabBounds = GetTabBounds(dockControlIndex);
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
                    GUI.Log("Failed to select item [" + itemText + "] the currently active item is [" + activeItem + "] in the " + Description + " tab", LogItemType.Information);
                    throw new Exception("Failed to select " + Description + " item");
                }

                Thread.Sleep(50);
            }
        }

        internal void ClickTab(int dockControlIndex, string itemText, MouseButton button)
        {
            int tabX = -1;
            int tabY = -1;

            Rectangle tabBounds = GetTabBounds(dockControlIndex);
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
                    GUI.Log("Failed to select item [" + itemText + "] the currently active item is [" + activeItem + "] in the " + Description + " tab", LogItemType.Information);
                    throw new Exception("Failed to select " + Description + " item");
                }

                Thread.Sleep(50);
            }
        }

        internal bool HasTab(int dockControlsIndex)
        {
            Rectangle tabBounds = GetTabBounds(dockControlsIndex);
            
            if (tabBounds.Width == 0 || tabBounds.Height == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal bool HasTitlebar(int dockControlsIndex)
        {
            Rectangle titlebarBounds = GetTitlebarBounds(dockControlsIndex);

            if (titlebarBounds.Width == 0 || titlebarBounds.Height == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal Rectangle GetTabBounds(int dockControlsIndex)
        {
            Rectangle tabBounds = new Rectangle();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlsIndex));
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

        internal Rectangle GetTitlebarBounds(int dockControlsIndex)
        {
            Rectangle titlebarBounds = new Rectangle();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, dockControlsIndex));
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

            GUI.Log("Remove [" + itemPattern + "] from the " + Description, LogItemType.Action);

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
                string tabText = GUI.m_APE.GetValueFromMessage();
                dynamic parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle != null && parentHandle == this.Handle)
                {
                    if (Regex.IsMatch(tabText, itemPattern))
                    {
                        Rectangle titlebarBounds = GetTitlebarBounds(dockControlIndex);

                        //Click 10 pixels in from the right hand side of the window and 10 pixels down
                        base.MoveTo(titlebarBounds.Right - 10, titlebarBounds.Top + 10);
                        Thread.Sleep(20);
                        base.SingleClickInternal(titlebarBounds.Right - 10, titlebarBounds.Top + 10, MouseButton.Left, MouseKeyModifier.None);

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
                                throw new Exception("Failed to remove the " + Description + " tab");
                            }

                            Thread.Sleep(50);
                        }
                    }
                }
            }

            throw new Exception("Failed to find " + Description + " tab");
        }
    }
}
