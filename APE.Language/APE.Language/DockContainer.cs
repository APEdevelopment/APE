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
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// Selects the specified item in the container
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
            GUI.Log("Select [" + itemPattern + "] from " + Identity.Description, LogItemType.Action);

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
                            string activeTab = ActiveItem();

                            if (activeTab == tabText)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                GUI.Log("Failed to select tab [" + tabText + "] the currently active tab is [" + activeTab + "] in the " + Description + " tab", LogItemType.Information);
                                throw new Exception("Failed to select " + Description + " tab");
                            }

                            Thread.Sleep(50);
                        }

                        return;
                    }
                }
            }

            throw new Exception("Failed to find " + Description + " tab");
        }

        internal Rectangle GetTabBounds(int tab)
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
        /// The currently selected item in the container
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
            string tabText = ActiveItem();
            if (!Regex.IsMatch(tabText, itemPattern))
            {
                SingleClickItem(itemPattern);
            }

            //Get the size of the window
            NM.tagRect ClientRect;
            NM.GetClientRect(Identity.Handle, out ClientRect);

            int InitialItems = ItemsCount();
            int CurrentItems = InitialItems;

            GUI.Log("Remove [" + itemPattern + "] from the " + Description, LogItemType.Action);

            //Click 10 pixels in from the right hand side of the window and 10 pixels down
            base.MoveTo(ClientRect.right - 10, 10);
            Thread.Sleep(20);
            base.SingleClickInternal(ClientRect.right - 10, 10, MouseButton.Left, MouseKeyModifier.None);

            //Wait for the number of items to decrease by one
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
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
                    throw new Exception("Failed to remove " + Description + " tab");
                }

                Thread.Sleep(50);
            }
        }
    }
}
