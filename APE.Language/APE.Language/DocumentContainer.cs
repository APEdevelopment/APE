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
using System.Text.RegularExpressions;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// TD.SandDock.DocumentContainer
    /// </summary>
    public sealed class GUIDocumentContainer : GUIDockContainer
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

        /// <summary>
        /// Clicks on the specified item in the document container using the specified mouse button
        /// </summary>
        /// <param name="itemPattern">The text of the item.  Regular expression syntax is supported see msdn for details: 
        /// https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx </param>
        /// <param name="button">The mouse button to click with</param>
        public override void SingleClickItem(string itemPattern, MouseButton button)
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
                IntPtr parentHandle = GUI.m_APE.GetValueFromMessage();

                if (parentHandle == this.Handle)
                {
                    if (Regex.IsMatch(tabText, itemPattern))
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
                            Rectangle tabBounds = GetTabBounds(dockControlIndex);
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
                                throw new Exception(Description + " did not go idle within timeout after scrolling");
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw new Exception("Failed to scroll the " + Description + " tab into view");
                            }
                        }

                        base.SingleClickInternal(tabX, tabY, button, MouseKeyModifier.None);

                        //Wait for the active document to be the tab we selected
                        timer = Stopwatch.StartNew();
                        do
                        {
                            string activeTab = ActiveItem();

                            if (activeTab == tabText)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw new Exception("Failed to select the " + Description + " tab");
                            }

                            Thread.Sleep(50);
                        }
                        while (true);
                        timer.Stop();

                        return;
                    }
                }
            }

            throw new Exception("Failed to find the " + Description + " tab");
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
