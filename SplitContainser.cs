//
//Copyright 2016-2017 David Beales
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
using APE.Communication;
using System;
using System.Drawing;
using System.Reflection;
using System.Windows.Forms;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.SplitContainer
    /// </summary>
    public sealed class GUISplitContainer : GUIObject
    {
        /// <summary>
        /// The top panel in the split container
        /// </summary>
        public TopPanelInternal TopPanel;
        /// <summary>
        /// The bottom panel in the split container
        /// </summary>
        public BottomPanelInternal BottomPanel;
        /// <summary>
        /// The left panel in the split container
        /// </summary>
        public LeftPanelInternal LeftPanel;
        /// <summary>
        /// The right panel in the split container
        /// </summary>
        public RightPanelInternal RightPanel;

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUISplitContainer(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
            TopPanel = new TopPanelInternal(this);
            BottomPanel = new BottomPanelInternal(this);
            LeftPanel = new LeftPanelInternal(this);
            RightPanel = new RightPanelInternal(this);
        }

        private enum SplitterPanel
        {
            Panel1,
            Panel2,
        }

        private enum SplitterDimension
        {
            Height,
            Width,
        }

        private enum Axis
        {
            X,
            Y,
        }

        private void Resize(int adjustment, Axis axis)
        {
            Rectangle splitter = GetSplitterRectangle();
            int x = splitter.X + (splitter.Width / 2);
            int y = splitter.Y + (splitter.Height / 2);

            uint DoubleClickTimer = (uint)SystemInformation.DoubleClickTime;
            Input.Block();
            try
            {
                TimerResolution.SetMaxTimerResolution();
                NM.SetDoubleClickTime(1);

                // Mouse down over the splitter
                base.MouseDownInternal(x, y, MouseButton.Left, MouseKeyModifier.None);
                switch (axis)
                {
                    case Axis.X:
                        // Move the mouse to half way to make sure the control enters dragmode
                        Input.MouseMove(Handle, Description, x - (adjustment / 2), y, false);
                        // Mouse up where we want the splitter
                        base.MouseUpInternal(x - adjustment, y, MouseButton.Left, MouseKeyModifier.None);
                        break;
                    case Axis.Y:
                        // Move the mouse to half way to make sure the control enters dragmode
                        Input.MouseMove(Handle, Description, x, y - (adjustment / 2), false);
                        // Mouse up where we want the splitter
                        base.MouseUpInternal(x, y - adjustment, MouseButton.Left, MouseKeyModifier.None);
                        break;
                }
            }
            catch
            {
                NM.SetDoubleClickTime(DoubleClickTimer);    //Reset double click timer
                Input.Reset();                              //Reset the mouse blocking
                throw;
            }
            finally
            {
                TimerResolution.UnsetMaxTimerResolution();
                NM.SetDoubleClickTime(DoubleClickTimer);
                Input.Unblock();
            }
        }

        private string SplitterOrientation()
        {
            //Get the orientation
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Orientation", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string orientation = GUI.m_APE.GetValueFromMessage();
            return orientation;
        }

        private Rectangle GetSplitterRectangle()
        {
            //Get the panel dimension
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SplitterRectangle", MemberTypes.Property);
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

            Rectangle itemRectangle = new Rectangle(x, y, width, height);
            return itemRectangle;
        }

        private int GetPanelMinimumDimension(SplitterPanel panel)
        {
            //Get the panel dimension
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            switch (panel)
            {
                case SplitterPanel.Panel1:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panel1MinSize", MemberTypes.Property);
                    break;
                case SplitterPanel.Panel2:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panel2MinSize", MemberTypes.Property);
                    break;
                default:
                    throw new Exception("Unknown splitter panel: " + panel.ToString());
            }
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int dimensionSize = GUI.m_APE.GetValueFromMessage();
            return dimensionSize;
        }

        private int GetPanelDimension(SplitterPanel panel, SplitterDimension dimension)
        {
            //Get the panel dimension
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            switch (panel)
            {
                case SplitterPanel.Panel1:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panel1", MemberTypes.Property);
                    break;
                case SplitterPanel.Panel2:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panel2", MemberTypes.Property);
                    break;
                default:
                    throw new Exception("Unknown splitter panel: " + panel.ToString());
            }
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ClientSize", MemberTypes.Property);
            switch (dimension)
            {
                case SplitterDimension.Width:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Width", MemberTypes.Property);
                    break;
                case SplitterDimension.Height:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Height", MemberTypes.Property);
                    break;
                default:
                    throw new Exception("Unknown splitter panel: " + panel.ToString());
            }
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int dimensionSize = GUI.m_APE.GetValueFromMessage();
            return dimensionSize;
        }

        /// <summary>
        /// The top panel in the split container
        /// </summary>
        public class TopPanelInternal
        {
            GUISplitContainer Container;

            internal TopPanelInternal(GUISplitContainer container)
            {
                Container = container;
            }

            /// <summary>
            /// Resizes the top panel
            /// </summary>
            /// <param name="newHeight">The new height of the top panel in pixels</param>
            public void Resize(int newHeight)
            {
                if (Container.SplitterOrientation() != "Horizontal")
                {
                    throw new Exception("Top panel does not exist");
                }

                // Resize the panel
                int actualNewHeight = newHeight;
                int bottomPanelMinimumHeight = Container.GetPanelMinimumDimension(SplitterPanel.Panel2);
                int topPanelMinimumHeight = Container.GetPanelMinimumDimension(SplitterPanel.Panel1);
                int splitterHeight = Container.GetSplitterRectangle().Height;

                if (actualNewHeight > Container.Height - bottomPanelMinimumHeight - splitterHeight)
                {
                    actualNewHeight = Container.Height - bottomPanelMinimumHeight - splitterHeight;
                }
                if (actualNewHeight < topPanelMinimumHeight)
                {
                    actualNewHeight = topPanelMinimumHeight;
                }

                int currentHeight = Container.GetPanelDimension(SplitterPanel.Panel1, SplitterDimension.Height);
                int adjustment = 0 - (actualNewHeight - currentHeight);

                if (actualNewHeight != newHeight)
                {
                    GUI.Log("Warning new height adjusted to " + actualNewHeight + " due to size restrictions of the " + Container.Description, LogItemType.Warning);
                }

                if (actualNewHeight == currentHeight)
                {
                    GUI.Log("Make sure the height of the top panel in the " + Container.Description + " is " + actualNewHeight.ToString(), LogItemType.Action);
                    return;
                }

                GUI.Log("Resize the top panel of the " + Container.Identity.Description + " to have a height of " + actualNewHeight.ToString() + " pixels ", LogItemType.Action);
                Container.Resize(adjustment, Axis.Y);
            }

            /// <summary>
            /// The current height of the top panel in pixels
            /// </summary>
            /// <returns>Panel height in pixels</returns>
            public int Height()
            {
                int currentHeight = Container.GetPanelDimension(SplitterPanel.Panel1, SplitterDimension.Height);
                return currentHeight;
            }
        }

        /// <summary>
        /// The bottom panel in the split container
        /// </summary>
        public class BottomPanelInternal
        {
            GUISplitContainer Container;

            internal BottomPanelInternal(GUISplitContainer container)
            {
                Container = container;
            }

            /// <summary>
            /// Resizes the bottom panel
            /// </summary>
            /// <param name="newHeight">The new height of the bottom panel in pixels</param>
            public void Resize(int newHeight)
            {
                if (Container.SplitterOrientation() != "Horizontal")
                {
                    throw new Exception("Bottom panel does not exist");
                }

                // Resize the panel
                int actualNewHeight = newHeight;
                int bottomPanelMinimumHeight = Container.GetPanelMinimumDimension(SplitterPanel.Panel2);
                int topPanelMinimumHeight = Container.GetPanelMinimumDimension(SplitterPanel.Panel1);
                int splitterHeight = Container.GetSplitterRectangle().Height;

                if (actualNewHeight > Container.Height - topPanelMinimumHeight - splitterHeight)
                {
                    actualNewHeight = Container.Height - topPanelMinimumHeight - splitterHeight;
                }
                if (actualNewHeight < bottomPanelMinimumHeight)
                {
                    actualNewHeight = bottomPanelMinimumHeight;
                }

                int currentHeight = Container.GetPanelDimension(SplitterPanel.Panel2, SplitterDimension.Height);
                int adjustment = actualNewHeight - currentHeight;

                if (actualNewHeight != newHeight)
                {
                    GUI.Log("Warning new height adjusted to " + actualNewHeight + " due to size restrictions of the " + Container.Description, LogItemType.Warning);
                }
                    
                if (actualNewHeight == currentHeight)
                {
                    GUI.Log("Make sure the height of the bottom panel in the " + Container.Description + " is " + actualNewHeight.ToString(), LogItemType.Action);
                    return;
                }

                GUI.Log("Resize the bottom panel of the " + Container.Identity.Description + " to have a height of " + actualNewHeight.ToString() + " pixels ", LogItemType.Action);
                Container.Resize(adjustment, Axis.Y);
            }

            /// <summary>
            /// The current height of the bottom panel in pixels
            /// </summary>
            /// <returns>Panel height in pixels</returns>
            public int Height()
            {
                int currentHeight = Container.GetPanelDimension(SplitterPanel.Panel2, SplitterDimension.Height);
                return currentHeight;
            }
        }

        /// <summary>
        /// The left panel in the split container
        /// </summary>
        public class LeftPanelInternal
        {
            GUISplitContainer Container;

            internal LeftPanelInternal(GUISplitContainer container)
            {
                Container = container;
            }

            /// <summary>
            /// Resizes the left panel
            /// </summary>
            /// <param name="newWidth">The new width of the left panel in pixels</param>
            public void Resize(int newWidth)
            {
                if (Container.SplitterOrientation() != "Vertical")
                {
                    throw new Exception("Left panel does not exist");
                }

                // Resize the panel
                int actualNewWidth = newWidth;
                int rightPanelMinimumWidth = Container.GetPanelMinimumDimension(SplitterPanel.Panel2);
                int leftPanelMinimumWidth = Container.GetPanelMinimumDimension(SplitterPanel.Panel1);
                int splitterWidth = Container.GetSplitterRectangle().Width;

                if (actualNewWidth > Container.Width - rightPanelMinimumWidth - splitterWidth)
                {
                    actualNewWidth = Container.Width - rightPanelMinimumWidth - splitterWidth;
                }
                if (actualNewWidth < leftPanelMinimumWidth)
                {
                    actualNewWidth = leftPanelMinimumWidth;
                }

                int currentWidth = Container.GetPanelDimension(SplitterPanel.Panel1, SplitterDimension.Width);
                int adjustment = 0 - (actualNewWidth - currentWidth);
                    
                if (actualNewWidth != newWidth)
                {
                    GUI.Log("Warning new width adjusted to " + actualNewWidth + " due to size restrictions of the " + Container.Description, LogItemType.Warning);
                }

                if (actualNewWidth == currentWidth)
                {

                    GUI.Log("Make sure the width of the left panel in the " + Container.Description + " is " + actualNewWidth.ToString(), LogItemType.Action);
                    return;
                }

                GUI.Log("Resize the left panel of the " + Container.Identity.Description + " to have a width of " + actualNewWidth.ToString() + " pixels ", LogItemType.Action);
                Container.Resize(adjustment, Axis.X);
            }

            /// <summary>
            /// The current width of the left panel in pixels
            /// </summary>
            /// <returns>Panel width in pixels</returns>
            public int Width()
            {
                int currentWidth = Container.GetPanelDimension(SplitterPanel.Panel1, SplitterDimension.Width);
                return currentWidth;
            }
        }

        /// <summary>
        /// The right panel in the split container
        /// </summary>
        public class RightPanelInternal
        {
            GUISplitContainer Container;

            internal RightPanelInternal(GUISplitContainer container)
            {
                Container = container;
            }

            /// <summary>
            /// Resizes the right panel
            /// </summary>
            /// <param name="newWidth">The new width of the right panel in pixels</param>
            public void Resize(int newWidth)
            {
                if (Container.SplitterOrientation() != "Vertical")
                {
                    throw new Exception("Right panel does not exist");
                }

                // Resize the panel
                int actualNewWidth = newWidth;
                int rightPanelMinimumWidth = Container.GetPanelMinimumDimension(SplitterPanel.Panel2);
                int leftPanelMinimumWidth = Container.GetPanelMinimumDimension(SplitterPanel.Panel1);
                int splitterWidth = Container.GetSplitterRectangle().Width;

                if (actualNewWidth > Container.Width - leftPanelMinimumWidth - splitterWidth)
                {
                    actualNewWidth = Container.Width - leftPanelMinimumWidth - splitterWidth;
                }
                if (actualNewWidth < rightPanelMinimumWidth)
                {
                    actualNewWidth = rightPanelMinimumWidth;
                }

                int currentWidth = Container.GetPanelDimension(SplitterPanel.Panel2, SplitterDimension.Width);
                int adjustment = actualNewWidth - currentWidth;

                if (actualNewWidth != newWidth)
                {
                    GUI.Log("Warning new width adjusted to " + actualNewWidth + " due to size restrictions of the " + Container.Description, LogItemType.Warning);
                }
                
                if (actualNewWidth == currentWidth)
                {
                    GUI.Log("Make sure the width of the right panel in the " + Container.Description + " is " + actualNewWidth.ToString(), LogItemType.Action);
                    return;
                }

                GUI.Log("Resize the right panel of the " + Container.Identity.Description + " to have a width of " + actualNewWidth.ToString() + " pixels ", LogItemType.Action);
                Container.Resize(adjustment, Axis.X);
            }

            /// <summary>
            /// The current width of the right panel in pixels
            /// </summary>
            /// <returns>Panel width in pixels</returns>
            public int Width()
            {
                int currentWidth = Container.GetPanelDimension(SplitterPanel.Panel2, SplitterDimension.Width);
                return currentWidth;
            }
        }
    }
}
