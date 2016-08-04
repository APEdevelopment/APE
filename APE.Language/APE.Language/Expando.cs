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
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// XPExplorerBar.Expando
    /// </summary>
    public sealed class GUIExpando : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIExpando(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Expand the expando control
        /// </summary>
        public void Expand()
        {
            if (!IsCollapsed())
            {
                throw new Exception("Expando is already collapsed");
            }

            if (!CanCollapse())
            {
                throw new Exception("Expando can not be collapsed");
            }

            GUI.Log("Single Left click on the " + m_DescriptionOfControl + " to expand it", LogItemTypeEnum.Action);
            ClickTitleBar(MouseButton.Left);

            if (IsCollapsed())
            {
                throw new Exception("Failed to expand expando");
            }
        }

        /// <summary>
        /// Collapse the expando control
        /// </summary>
        public void Collapse()
        {
            if (IsCollapsed())
            {
                throw new Exception("Expando is already collapsed");
            }

            if (!CanCollapse())
            {
                throw new Exception("Expando can not be collapsed");
            }

            GUI.Log("Single Left click on the " + m_DescriptionOfControl + " to collapse it", LogItemTypeEnum.Action);
            ClickTitleBar(MouseButton.Left);

            if (!IsCollapsed())
            {
                throw new Exception("Failed to collapse expando");
            }
        }

        /// <summary>
        /// Determines if the expando is currently collapsed or not
        /// </summary>
        /// <returns>True if the expando is collapsed otherwise false</returns>
        public bool IsCollapsed()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Collapsed", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            bool isCollapsed = GUI.m_APE.GetValueFromMessage();

            return isCollapsed;
        }

        /// <summary>
        /// Determines if the expando can be collapsed
        /// </summary>
        /// <returns>True if the expando can be collapsed otherwise false</returns>
        public bool CanCollapse()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CanCollapse", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            bool canCollapse = GUI.m_APE.GetValueFromMessage();

            return canCollapse;
        }

        private bool Animating()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Animating", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            bool animating = GUI.m_APE.GetValueFromMessage();

            return animating;
        }

        private Rectangle TitleBarRectangle()
        {
            Rectangle titleBarRectangle = new Rectangle();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TitleBarRectangle", MemberTypes.Property);
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
            //get the values returned
            titleBarRectangle.X = GUI.m_APE.GetValueFromMessage();
            titleBarRectangle.Y = GUI.m_APE.GetValueFromMessage();
            titleBarRectangle.Width = GUI.m_APE.GetValueFromMessage();
            titleBarRectangle.Height = GUI.m_APE.GetValueFromMessage();

            return titleBarRectangle;
        }

        private void ClickTitleBar(MouseButton button)
        {
            Rectangle titleBarRectangle = TitleBarRectangle();

            base.MouseSingleClickInternal(titleBarRectangle.X + (titleBarRectangle.Width / 2), titleBarRectangle.Y + (titleBarRectangle.Height / 2), button, MouseKeyModifier.None);

            if (button == MouseButton.Left)
            {
                Stopwatch timer = Stopwatch.StartNew();
                while (true)
                {
                    if (!Animating())
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Expando failed to stop animating");
                    }

                    Thread.Sleep(15);
                }
            }
        }
    }
}
