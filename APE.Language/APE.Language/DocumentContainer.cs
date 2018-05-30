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

            string tabText;
            int dockControlIndex = GetDockControlIndex(itemPattern, out tabText);
            if (dockControlIndex == -1)
            {
                throw GUI.ApeException("Failed to find the " + Description + " item");
            }

            int rightPadding = RightPadding();
            Rectangle leftScrollButton = GetButtonBounds(dockControlIndex, ButtonToGet.ScrollLeft);
            int leftScrollButtonX = leftScrollButton.X + (leftScrollButton.Width / 2);
            int leftScrollButtonY = leftScrollButton.Y + (leftScrollButton.Height / 2);
            Rectangle rightScrollButton = GetButtonBounds(dockControlIndex, ButtonToGet.ScrollRight);
            int rightScrollButtonX = rightScrollButton.X + (rightScrollButton.Width / 2);
            int rightScrollButtonY = rightScrollButton.Y + (rightScrollButton.Height / 2);

            int tabX = -1;
            int tabY = -1;
            int width = this.Width;

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                Rectangle tabBounds = GetTabBoundsInternal(dockControlIndex);
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
                    throw GUI.ApeException(Description + " did not go idle within timeout after scrolling");
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException("Failed to scroll the " + Description + " tab into view");
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
                    throw GUI.ApeException("Failed to select the " + Description + " tab");
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
    }
}
