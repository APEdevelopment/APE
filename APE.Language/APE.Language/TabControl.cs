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
    /// TD.SandDock.DockableWindow
    /// </summary>
    public sealed class GUITabControl : GUIFocusableObject
    {
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
        }

        /// <summary>
        /// Returns the number of tabs in the tabcontrol
        /// </summary>
        /// <returns>The number of tabs</returns>
        public int Count()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TabCount", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
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
        public int Index(string tabText)
        {
            int tabCount = Count();

            for (int x = 0; x < tabCount; x++)
            {
                if (TabText(x) == tabText)
                {
                    return x;
                }
            }

            throw new Exception("Failed to find tab " + tabText);
        }

        /// <summary>
        /// Returns the text of the tab at the specified index
        /// </summary>
        /// <param name="tabIndex">The index of the tab</param>
        /// <returns>The text of the tab</returns>
        public string TabText(int tabIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TabPages", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, tabIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
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
        public void Select(string tabText)
        {
            int tabIndex = Index(tabText);
            SelectInternal(tabIndex, tabText, MouseButton.Left);
        }

        /// <summary>
        /// Click on a tab using the specified mouse button
        /// </summary>
        /// <param name="tabText">The tab to select</param>
        /// <param name="button">The mouse button to click with</param>
        public void Select(string tabText, MouseButton button)
        {
            int tabIndex = Index(tabText);
            SelectInternal(tabIndex, tabText, button);
        }

        /// <summary>
        /// Select a tab
        /// </summary>
        /// <param name="tabIndex">The tab index to select</param>
        public void Select(int tabIndex)
        {
            string tabText = TabText(tabIndex);
            SelectInternal(tabIndex, tabText, MouseButton.Left);
        }

        /// <summary>
        /// Click on a tab using the specified mouse button
        /// </summary>
        /// <param name="tabIndex">The tab index to select</param>
        /// <param name="button">The mouse button to click with</param>
        public void Select(int tabIndex, MouseButton button)
        {
            string tabText = TabText(tabIndex);
            SelectInternal(tabIndex, tabText, button);
        }

        private void SelectInternal(int tabIndex, string tabName, MouseButton button)
        {
            if (SelectedTabText() == tabName)
            {
                GUI.Log("Ensure tab " + tabName + " is selected in the " + m_DescriptionOfControl, LogItemType.Action);
                return;
            }

            bool multiLine = MultiLine();
            Rectangle tabRectangle = TabRectangle(tabIndex);

            NM.tagRect tabRect;
            NM.GetClientRect(Identity.Handle, out tabRect);
            int width = tabRect.right;

            // multiline means no scroll button
            if (!multiLine)
            {
                IntPtr upDownControl = NM.GetWindow(Identity.Handle, NM.GetWindowCmd.GW_CHILD);

                while (true)
                {
                    if (upDownControl == IntPtr.Zero)
                    {
                        break;
                    }

                    string className = NM.GetClassName(upDownControl);

                    if (className == "msctls_updown32")
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

                    GUIButton scrollTabButton = new GUIButton(m_ParentForm, "Scroll button", new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.Handle, upDownControl));

                    while (true)
                    {
                        if (tabRectangle.X < 0)
                        {
                            GUI.Log("Click " + m_DescriptionOfControl + " left button", LogItemType.Action);
                            scrollTabButton.MouseSingleClickInternal(5, 5, MouseButton.Left, MouseKeyModifier.None);
                        }
                        else if (tabRectangle.X > width)
                        {
                            GUI.Log("Click " + m_DescriptionOfControl + " right button", LogItemType.Action);
                            scrollTabButton.MouseSingleClickInternal(scrollButtonsRect.right - 5, 5, MouseButton.Left, MouseKeyModifier.None);
                        }
                        else
                        {
                            break;
                        }

                        tabRectangle = TabRectangle(tabIndex);
                    }
                }
            }

            GUI.Log("Select tab " + tabName + " in " + m_DescriptionOfControl, LogItemType.Action);
            int x = tabRectangle.X + (tabRectangle.Width / 2);
            int y = tabRectangle.Y + (tabRectangle.Height / 2);
            if (!multiLine)
            {
                if (x > width - 1)
                {
                    x = width - 2;
                }
            }
            base.MouseSingleClickInternal(x, y, button, MouseKeyModifier.None);

            // Confirm the expected tab has been selected
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw new Exception("Failed to select tab " + tabName);
                }

                if (SelectedTabText() == tabName)
                {
                    break;
                }

                Thread.Sleep(15);
            }
        }

        private string SelectedTabText()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedTab", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Text", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string selectedTabText = GUI.m_APE.GetValueFromMessage();

            return selectedTabText;
        }

        private bool MultiLine()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Multiline", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool multiLine = GUI.m_APE.GetValueFromMessage();

            return multiLine;
        }

        private Rectangle TabRectangle(int tabIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetTabRect", MemberTypes.Method, new Parameter(GUI.m_APE, tabIndex));
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
            int x = GUI.m_APE.GetValueFromMessage();
            int y = GUI.m_APE.GetValueFromMessage();
            int width = GUI.m_APE.GetValueFromMessage();
            int height = GUI.m_APE.GetValueFromMessage();

            Rectangle tabRectangle = new Rectangle(x, y, width, height);
            
            return tabRectangle;
        }
    }
}