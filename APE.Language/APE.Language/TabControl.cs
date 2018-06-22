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
using APE.Communication;
using System.Reflection;
using System.Drawing;
using NM = APE.Native.NativeMethods;
using System.Diagnostics;
using System.Threading;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.TabControl
    /// SftTabsLib.SftTabs
    /// MSComctlLib.TabStrip
    /// </summary>
    public sealed class GUITabControl : GUIFocusableObject
    {
        private int TabOffset = 0;

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITabControl(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
            if (Identity.TechnologyType == "Windows ActiveX" && Identity.TypeName == "TabStrip")
            {
                TabOffset = 1; //Convert from 1 based to 0 based
            }
        }

        /// <summary>
        /// Returns the number of tabs in the tabcontrol
        /// </summary>
        /// <returns>The number of tabs</returns>
        public int TabCount()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            switch (Identity.TechnologyType)
            {
                case "Windows Forms (WinForms)":
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "TabCount", MemberTypes.Property);
                    break;
                case "Windows ActiveX":
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Tabs", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
                    break;
                default:
                    throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
            }
            
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int tabCount = GUI.m_APE.GetValueFromMessage();
            return tabCount;
        }

        /// <summary>
        /// Returns the index of the tab with the specified text
        /// </summary>
        /// <param name="tabText">The text of the tab</param>
        /// <returns>The index of the tab</returns>
        public int TabIndex(string tabText)
        {
            int tabCount = TabCount();

            for (int tabIndex = 0; tabIndex < tabCount; tabIndex++)
            {
                if (TabText(tabIndex) == tabText)
                {
                    return tabIndex;
                }
            }

            throw GUI.ApeException("Failed to find tab " + tabText);
        }

        /// <summary>
        /// Returns the text of the tab at the specified index
        /// </summary>
        /// <param name="tabIndex">The tab index to select (0 based)</param>
        /// <returns>The text of the tab</returns>
        public string TabText(int tabIndex)
        {
            int actualTabIndex = tabIndex + TabOffset;
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            switch (Identity.TechnologyType)
            {
                case "Windows Forms (WinForms)":
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TabPages", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, actualTabIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                    break;
                case "Windows ActiveX":
                    switch (Identity.TypeName)
                    {
                        case "SftTabs":
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "Tab", MemberTypes.Property, new Parameter(GUI.m_APE, actualTabIndex));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                            break;
                        case "TabStrip":
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "Tabs", MemberTypes.Property, new Parameter(GUI.m_APE, actualTabIndex));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Caption", MemberTypes.Property);
                            break;
                        default:
                            throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
                    }
                    break;
                default:
                    throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
            }
            
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string tabText = GUI.m_APE.GetValueFromMessage();
            return tabText;
        }

        /// <summary>
        /// Select a tab
        /// </summary>
        /// <param name="tabText">The tab to select</param>
        public void SingleClickTab(string tabText)
        {
            int tabIndex = TabIndex(tabText);
            SingleClickTabInternal(tabIndex, tabText, MouseButton.Left);
        }

        /// <summary>
        /// Click on a tab using the specified mouse button
        /// </summary>
        /// <param name="tabText">The tab to select</param>
        /// <param name="button">The mouse button to click with</param>
        public void SingleClickTab(string tabText, MouseButton button)
        {
            int tabIndex = TabIndex(tabText);
            SingleClickTabInternal(tabIndex, tabText, button);
        }

        /// <summary>
        /// Select a tab
        /// </summary>
        /// <param name="tabIndex">The tab index to select (0 based)</param>
        public void SingleClickTab(int tabIndex)
        {
            string tabText = TabText(tabIndex);
            SingleClickTabInternal(tabIndex, tabText, MouseButton.Left);
        }

        /// <summary>
        /// Click on a tab using the specified mouse button
        /// </summary>
        /// <param name="tabIndex">The tab index to select (0 based)</param>
        /// <param name="button">The mouse button to click with</param>
        public void SingleClickTab(int tabIndex, MouseButton button)
        {
            string tabText = TabText(tabIndex);
            SingleClickTabInternal(tabIndex, tabText, button);
        }

        private void SingleClickTabInternal(int tabIndex, string tabName, MouseButton button)
        {
            if (SelectedTabText() == tabName)
            {
                GUI.Log("Ensure tab " + tabName + " is selected in the " + Identity.Description, LogItemType.Action);
                return;
            }

            int actualTabIndex = tabIndex + TabOffset;

            bool multiLine = MultiLine();
            Rectangle tabRectangle = TabRectangle(actualTabIndex);

            NM.tagRect tabRect;
            NM.GetClientRect(Identity.Handle, out tabRect);
            int width = tabRect.right;

            // multiline means no scroll button (only top and bottom placement support a scroll button)
            if (!multiLine)
            {
                //We may need to scroll
                if (Identity.TechnologyType == "Windows Forms (WinForms)" || (Identity.TechnologyType == "Windows ActiveX" && Identity.TypeName == "TabStrip"))
                {
                    IntPtr upDownControl = NM.GetWindow(Identity.Handle, NM.GetWindowCmd.GW_CHILD);

                    while (true)
                    {
                        if (upDownControl == IntPtr.Zero)
                        {
                            break;
                        }

                        string className = NM.GetClassName(upDownControl);

                        if (className == "msctls_updown32" || className == "msvb_lib_updown")
                        {
                            break;
                        }

                        upDownControl = NM.GetWindow(upDownControl, NM.GetWindowCmd.GW_HWNDNEXT);
                    }

                    if (upDownControl != IntPtr.Zero)
                    {
                        // work out if we need to scroll the tab into view
                        NM.tagRect scrollButtonsRect;
                        NM.GetClientRect(upDownControl, out scrollButtonsRect);

                        width = tabRect.right - scrollButtonsRect.right;

                        GUIButton scrollTabButton = new GUIButton(ParentForm, "Scroll button", new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.Handle, upDownControl));

                        while (true)
                        {
                            if (tabRectangle.X < 0)
                            {
                                GUI.Log("Click " + Identity.Description + " left button", LogItemType.Action);
                                scrollTabButton.SingleClickInternal(5, 5, MouseButton.Left, MouseKeyModifier.None);
                            }
                            else if (tabRectangle.X > width)
                            {
                                GUI.Log("Click " + Identity.Description + " right button", LogItemType.Action);
                                scrollTabButton.SingleClickInternal(scrollButtonsRect.right - 5, 5, MouseButton.Left, MouseKeyModifier.None);
                            }
                            else
                            {
                                break;
                            }

                            tabRectangle = TabRectangle(actualTabIndex);
                        }
                    }
                }
                else if (Identity.TechnologyType == "Windows ActiveX" && Identity.TypeName == "SftTabs")
                {
                    tabRectangle = TabRectangle(actualTabIndex);

                    while (tabRectangle.Width == 0 || tabRectangle.Height == 0)
                    {
                        int visibleTab = SftTabVisibleTab();
                        if (actualTabIndex < visibleTab)
                        {
                            SftTabClickLeftUp();
                        }
                        else
                        {
                            SftTabClickRightDown();
                        }
                        tabRectangle = TabRectangle(actualTabIndex);
                    }
                }
                else
                {
                    throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
                }
            }

            GUI.Log("Select tab " + tabName + " in the " + Identity.Description, LogItemType.Action);
            int x = tabRectangle.X + (tabRectangle.Width / 2);
            int y = tabRectangle.Y + (tabRectangle.Height / 2);
            if (!multiLine)
            {
                if (x > width - 1)
                {
                    x = width - 2;
                }
            }
            base.SingleClickInternal(x, y, button, MouseKeyModifier.None);

            // Confirm the expected tab has been selected
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (SelectedTabText() == tabName)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw GUI.ApeException("Failed to select tab " + tabName);
                }

                Thread.Sleep(15);
            }
        }

        private void SftTabClickLeftUp()
        {
            int originalVisibleTab = SftTabVisibleTab();
            GUIButton button = new GUIButton(ParentForm, Description + " left / up button", new Identifier(Identifiers.ChildOf, this), new Identifier(Identifiers.TypeName, "SftBtnControl"), new Identifier(Identifiers.Index, 1));
            button.SingleClick();

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                int currentVisibleTab = SftTabVisibleTab();
                if (originalVisibleTab != currentVisibleTab)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw GUI.ApeException("Failed to scroll the " + Description + " tab");
                }

                Thread.Sleep(15);
            }
        }

        private void SftTabClickRightDown()
        {
            int originalVisibleTab = SftTabVisibleTab();
            GUIButton button = new GUIButton(ParentForm, Description + " right / down button", new Identifier(Identifiers.ChildOf, this), new Identifier(Identifiers.TypeName, "SftBtnControl"), new Identifier(Identifiers.Index, 2));
            button.SingleClick();

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                int currentVisibleTab = SftTabVisibleTab();
                if (originalVisibleTab != currentVisibleTab)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw GUI.ApeException("Failed to scroll the " + Description + " tab");
                }

                Thread.Sleep(15);
            }
        }

        private int SftTabVisibleTab()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Scrolling", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "VisibleTab", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int visibleTab = GUI.m_APE.GetValueFromMessage();
            return visibleTab;
        }

        private string SelectedTabText()
        {
            string selectedTabText;
            switch (Identity.TechnologyType)
            {
                case "Windows Forms (WinForms)":
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedTab", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Text", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    selectedTabText = GUI.m_APE.GetValueFromMessage();
                    break;
                case "Windows ActiveX":
                    int currentTabIndex;
                    switch (Identity.TypeName)
                    {
                        case "SftTabs":
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Tabs", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Current", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            currentTabIndex = GUI.m_APE.GetValueFromMessage();
                            break;
                        case "TabStrip":
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedItem", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Index", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            currentTabIndex = GUI.m_APE.GetValueFromMessage();
                            break;
                        default:
                            throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
                    }
                    selectedTabText = TabText(currentTabIndex - TabOffset);
                    break;
                default:
                    throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
            }
            return selectedTabText;
        }

        private bool MultiLine()
        {
            bool multiLine;
            switch (Identity.TechnologyType)
            {
                case "Windows Forms (WinForms)":
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Multiline", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    multiLine = GUI.m_APE.GetValueFromMessage();
                    break;
                case "Windows ActiveX":
                    switch (Identity.TypeName)
                    {
                        case "SftTabs":
                            multiLine = false;
                            break;
                        case "TabStrip":
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "MultiRow", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            multiLine = GUI.m_APE.GetValueFromMessage();
                            break;
                        default:
                            throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
                    }
                    break;
                default:
                    throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
            }
            return multiLine;
        }

        private Rectangle TabRectangle(int actualTabIndex)
        {
            int x;
            int y;
            int width;
            int height;
            switch (Identity.TechnologyType)
            {
                case "Windows Forms (WinForms)":
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetTabRect", MemberTypes.Method, new Parameter(GUI.m_APE, actualTabIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "X", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Y", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store4, "Width", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Height", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    x = GUI.m_APE.GetValueFromMessage();
                    y = GUI.m_APE.GetValueFromMessage();
                    width = GUI.m_APE.GetValueFromMessage();
                    height = GUI.m_APE.GetValueFromMessage();
                    break;
                case "Windows ActiveX":
                    switch (Identity.TypeName)
                    {
                        case "SftTabs":
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Tab", MemberTypes.Method, new Parameter(GUI.m_APE, actualTabIndex));
                            GUI.m_APE.AddQueryMessageGetTabRect(DataStores.Store1);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            x = GUI.m_APE.GetValueFromMessage();
                            y = GUI.m_APE.GetValueFromMessage();
                            width = GUI.m_APE.GetValueFromMessage();
                            height = GUI.m_APE.GetValueFromMessage();
                            break;
                        case "TabStrip":
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Tabs", MemberTypes.Method, new Parameter(GUI.m_APE, actualTabIndex));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Left", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Top", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store4, "Width", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Height", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            x = (int)(Math.Round(GUI.m_APE.GetValueFromMessage()));
                            y = (int)(Math.Round(GUI.m_APE.GetValueFromMessage()));
                            width = (int)(Math.Round(GUI.m_APE.GetValueFromMessage()));
                            height = (int)(Math.Round(GUI.m_APE.GetValueFromMessage()));

                           IntPtr containerHandle = NM.GetAncestor(Handle, NM.GetAncestorFlags.GetParent);

                            NM.tagPoint containerClientPoint = new NM.tagPoint();
                            NM.ClientToScreen(containerHandle, ref containerClientPoint);

                            NM.tagPoint clientPoint = new NM.tagPoint();
                            NM.ClientToScreen(Handle, ref clientPoint);

                            int containerOffsetX = clientPoint.x - containerClientPoint.x;
                            int containerOffsetY = clientPoint.y - containerClientPoint.y;

                            ScalingMode scaleMode = TabStripScalingMode();
                            switch (scaleMode)
                            {
                                case ScalingMode.Twip: //Twip
                                    x = TwipsToPixels(x, Direction.Horizontal);
                                    y = TwipsToPixels(y, Direction.Vertical);
                                    width = TwipsToPixels(width, Direction.Horizontal);
                                    height = TwipsToPixels(height, Direction.Vertical);
                                    break;
                                case ScalingMode.Pixel: //Pixel
                                    //do nothing
                                    break;
                                case ScalingMode.User:
                                case ScalingMode.Point:
                                case ScalingMode.Character:
                                case ScalingMode.Inch:
                                case ScalingMode.Millimeter:
                                case ScalingMode.Centimeter:
                                    throw GUI.ApeException("The " + Description + " scaling mode is of unsupported type " + scaleMode.ToString());
                                default:
                                    throw GUI.ApeException("The " + Description + " scaling mode is of unsupported type " + scaleMode.ToString());
                            }

                            x = x - containerOffsetX;
                            y = y - containerOffsetY;

                            int placement = TabStripPlacement();
                            switch (placement)
                            {
                                case 0: //Top
                                case 1: //Bottom
                                    //do nothing
                                    break;
                                case 2: //Left
                                case 3: //Right
                                    //swap width and height
                                    int temp = width;
                                    width = height;
                                    height = temp;
                                    break;
                                default:
                                    throw GUI.ApeException("The " + Description + " placement is of unsupported type " + placement.ToString());
                            }

                            break;
                        default:
                            throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
                    }
                    break;
                default:
                    throw GUI.ApeException("The " + Description + " is of an unsupported type " + Identity.TypeNameSpace + "." + Identity.TypeName);
            }
            Rectangle tabRectangle = new Rectangle(x, y, width, height);
            return tabRectangle;
        }

        private int TabStripPlacement()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Placement", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int placement = GUI.m_APE.GetValueFromMessage();
            return placement;
        }

        private ScalingMode TabStripScalingMode()
        {
            IntPtr containerHandle = NM.GetAncestor(Handle, NM.GetAncestorFlags.GetParent);

            NM.tagPoint containerClientPoint = new NM.tagPoint();
            NM.ClientToScreen(containerHandle, ref containerClientPoint);

            NM.tagPoint clientPoint = new NM.tagPoint();
            NM.ClientToScreen(Handle, ref clientPoint);

            int containerOffsetX = clientPoint.x - containerClientPoint.x;
            int containerOffsetY = clientPoint.y - containerClientPoint.y;

            int placement = TabStripPlacement();
            switch (placement)
            {
                case 0: //Top
                case 1: //Bottom
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ClientLeft", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int clientLeft = (int)(Math.Round(GUI.m_APE.GetValueFromMessage()));
                    if (clientLeft - containerOffsetX < 10)
                    {
                        return ScalingMode.Pixel;
                    }
                    if (TwipsToPixels(clientLeft, Direction.Horizontal) - containerOffsetX < 10)
                    {
                        return ScalingMode.Twip;
                    }
                    throw GUI.ApeException("Unsupported scaling mode");
                case 2: //Left
                case 3: //Right
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ClientTop", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int clientTop = (int)(Math.Round(GUI.m_APE.GetValueFromMessage()));
                    if (clientTop - containerOffsetY < 10)
                    {
                        return ScalingMode.Pixel;
                    }
                    if (TwipsToPixels(clientTop, Direction.Vertical) - containerOffsetY < 10)
                    {
                        return ScalingMode.Twip;
                    }
                    throw GUI.ApeException("Unsupported scaling mode");
                default:
                    throw GUI.ApeException("The " + Description + " placement is of unsupported type " + placement.ToString());
            }
        }

        private enum ScalingMode : int
        {
            User = 0,
            Twip = 1,
            Point = 2,
            Pixel = 3,
            Character = 4,
            Inch = 5,
            Millimeter = 6,
            Centimeter = 7,
        }
    }
}