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
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using APE.Communication;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// TD.SandDock.AutoHideBar
    /// </summary>
    public sealed class GUIAutoHideBar : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIAutoHideBar(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        internal int LayoutSystems()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int count = GUI.m_APE.GetValueFromMessage();
            return count;
        }

        internal int DockControlsInLayoutSystem(int layoutSystemIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, layoutSystemIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Controls", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int count = GUI.m_APE.GetValueFromMessage();
            return count;
        }

        internal bool IsShowingPopup()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ShowingLayoutSystem", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "IsPoppedUp", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic nullIfNotShowingPopup = GUI.m_APE.GetValueFromMessage();
            if (nullIfNotShowingPopup == null)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Returns the number of controls contained in the auto hide bar
        /// </summary>
        /// <returns>The number of controls in the auto hide bar</returns>
        public int ItemsCount()
        {
            int count = 0;
            int numberOfLayoutSystems = LayoutSystems();
            for (int layoutSystemIndex = 0; layoutSystemIndex < numberOfLayoutSystems; layoutSystemIndex++)
            {
                int numberOfDockControlsInLayoutSystem = DockControlsInLayoutSystem(layoutSystemIndex);
                for (int dockControlIndex = 0; dockControlIndex < numberOfDockControlsInLayoutSystem; dockControlIndex++)
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
            int index = -1;
            int numberOfLayoutSystems = LayoutSystems();
            for (int layoutSystemIndex = 0; layoutSystemIndex < numberOfLayoutSystems; layoutSystemIndex++)
            {
                int numberOfDockControlsInLayoutSystem = DockControlsInLayoutSystem(layoutSystemIndex);
                for (int dockControlIndex = 0; dockControlIndex < numberOfDockControlsInLayoutSystem; dockControlIndex++)
                {
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, layoutSystemIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Controls", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, dockControlIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "TabText", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string tabText = GUI.m_APE.GetValueFromMessage();

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
            int index = -1;
            int numberOfLayoutSystems = LayoutSystems();
            for (int layoutSystemIndex = 0; layoutSystemIndex < numberOfLayoutSystems; layoutSystemIndex++)
            {
                int numberOfDockControlsInLayoutSystem = DockControlsInLayoutSystem(layoutSystemIndex);
                for (int dockControlIndex = 0; dockControlIndex < numberOfDockControlsInLayoutSystem; dockControlIndex++)
                {
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, layoutSystemIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Controls", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, dockControlIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "TabText", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string tabText = GUI.m_APE.GetValueFromMessage();

                    index++;
                    if (index == itemIndex)
                    {
                        return tabText;
                    }
                }
            }
            throw new Exception("Failed to find the " + Description + " item");
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button.  If you need to show or hide an item then use ShowItem / HideItem as appropriate.
        /// </summary>
        /// <param name="itemIndex">The index of the item.</param>
        public void SingleClickItem(int itemIndex)
        {
            SingleClickItem(itemIndex, MouseButton.Left);
        }

        /// <summary>
        /// Clicks on the specified item in the container using the specified mouse button.  If you need to show or hide an item then use ShowItem / HideItem as appropriate.
        /// </summary>
        /// <param name="itemIndex">The index of the item.</param>
        /// <param name="button">The mouse button to click with</param>
        public void SingleClickItem(int itemIndex, MouseButton button)
        {
            GUI.Log("Single " + button.ToString() + " click on item index " + itemIndex.ToString() + " in the " + Identity.Description, LogItemType.Action);

            int index = -1;
            int numberOfLayoutSystems = LayoutSystems();
            for (int layoutSystemIndex = 0; layoutSystemIndex < numberOfLayoutSystems; layoutSystemIndex++)
            {
                int numberOfDockControlsInLayoutSystem = DockControlsInLayoutSystem(layoutSystemIndex);
                for (int dockControlIndex = 0; dockControlIndex < numberOfDockControlsInLayoutSystem; dockControlIndex++)
                {
                    index++;
                    if (index == itemIndex)
                    {
                        Rectangle tabBounds = new Rectangle();
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, layoutSystemIndex));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Controls", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, dockControlIndex));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "AutoHideTabBounds", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "X", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Y", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store8, "Width", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store9, "Height", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        tabBounds.X = GUI.m_APE.GetValueFromMessage();
                        tabBounds.Y = GUI.m_APE.GetValueFromMessage();
                        tabBounds.Width = GUI.m_APE.GetValueFromMessage();
                        tabBounds.Height = GUI.m_APE.GetValueFromMessage();

                        int tabX = tabBounds.X + (tabBounds.Width / 2);
                        int tabY = tabBounds.Y + (tabBounds.Height / 2);

                        base.SingleClickInternal(tabX, tabY, button, MouseKeyModifier.None);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button.  If you need to show or hide an item then use ShowItem / HideItem as appropriate.
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public void SingleClickItem(string itemPattern)
        {
            SingleClickItem(itemPattern, MouseButton.Left);
        }

        /// <summary>
        /// Clicks on the specified item in the container using the specified mouse button.  If you need to show or hide an item then use ShowItem / HideItem as appropriate.
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <param name="button">The mouse button to click with</param>
        public void SingleClickItem(string itemPattern, MouseButton button)
        {
            GUI.Log("Single " + button.ToString() + " click on the " + itemPattern + " item in the " + Identity.Description, LogItemType.Action);

            int numberOfLayoutSystems = LayoutSystems();
            for (int layoutSystemIndex = 0; layoutSystemIndex < numberOfLayoutSystems; layoutSystemIndex++)
            {
                int numberOfDockControlsInLayoutSystem = DockControlsInLayoutSystem(layoutSystemIndex);
                for (int dockControlIndex = 0; dockControlIndex < numberOfDockControlsInLayoutSystem; dockControlIndex++)
                {
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, layoutSystemIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Controls", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, dockControlIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "TabText", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string tabText = GUI.m_APE.GetValueFromMessage();

                    if (Regex.IsMatch(tabText, itemPattern))
                    {
                        Rectangle tabBounds = new Rectangle();
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystems", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, layoutSystemIndex));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Controls", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, dockControlIndex));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "AutoHideTabBounds", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "X", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Y", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store8, "Width", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store9, "Height", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        tabBounds.X = GUI.m_APE.GetValueFromMessage();
                        tabBounds.Y = GUI.m_APE.GetValueFromMessage();
                        tabBounds.Width = GUI.m_APE.GetValueFromMessage();
                        tabBounds.Height = GUI.m_APE.GetValueFromMessage();

                        int tabX = tabBounds.X + (tabBounds.Width / 2);
                        int tabY = tabBounds.Y + (tabBounds.Height / 2);

                        base.SingleClickInternal(tabX, tabY, button, MouseKeyModifier.None);
                        return;
                    }
                }
            }
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button and waits for it to be fully displayed
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public void ShowItem(string itemPattern)
        {
            SingleClickItem(itemPattern);

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (IsShowingPopup())
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw new Exception("Failed to show the " + Description + " item");
                }

                Thread.Sleep(15);
            }
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button and waits for it to be fully displayed
        /// </summary>
        /// <param name="itemIndex">The index of the item.</param>
        public void ShowItem(int itemIndex)
        {
            SingleClickItem(itemIndex);

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (IsShowingPopup())
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw new Exception("Failed to show the " + Description + " item");
                }

                Thread.Sleep(15);
            }
        }

        /// <summary>
        /// Removes focus from the attached process and waits for any displayed items to be hidden
        /// </summary>
        public void HideItem()
        {
            GUI.Log("Hide popup items in the " + Identity.Description, LogItemType.Action);
            // Move the mouse over and set focus to the viewport so the item will hide
            Input.MoveMouse(GUI.m_ViewPort.Left + 1, GUI.m_ViewPort.Top + 1);
            Input.SetFocus(GUI.m_ViewPort.m_ViewPortHandle, GUI.m_ViewPort.m_ViewPortHandle, "APE Viewport");
            
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (!IsShowingPopup())
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw new Exception("Failed to hide the " + Description + " item");
                }

                Thread.Sleep(15);
            }
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button and pins it
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        public void PinItem(string itemPattern)
        {
            ShowItem(itemPattern);
            IntPtr popupContainerHandle = GetPopupContainerHandle();
            GUIPopupContainer popupContainer = new GUIPopupContainer(ParentForm, Description + " popup container", new Identifier(Identifiers.Handle, popupContainerHandle));
            popupContainer.PinButtonClick();
        }

        /// <summary>
        /// Clicks on the specified item in the container using the left mouse button and pins it
        /// </summary>
        /// <param name="itemIndex">The index of the item.</param>
        public void PinItem(int itemIndex)
        {
            ShowItem(itemIndex);
            IntPtr popupContainerHandle = GetPopupContainerHandle();
            GUIPopupContainer popupContainer = new GUIPopupContainer(ParentForm, Description + " popup container", new Identifier(Identifiers.Handle, popupContainerHandle));
            popupContainer.PinButtonClick();
        }

        internal IntPtr GetPopupContainerHandle()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "PopupContainer", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            IntPtr popupContainerHandle = GUI.m_APE.GetValueFromMessage();
            return popupContainerHandle;
        }
    }
}
