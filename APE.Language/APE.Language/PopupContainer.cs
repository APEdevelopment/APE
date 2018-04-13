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
using APE.Communication;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// TD.SandDock.PopupContainer
    /// </summary>
    public sealed class GUIPopupContainer : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIPopupContainer(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Checks if the pin button is currently visible in the container
        /// </summary>
        /// <returns>True if visible otherwise false</returns>
        public bool PinButtonVisible()
        {
            bool visible = GetButtonVisibleInternal(ButtonToGet.Pin);
            return visible;
        }

        /// <summary>
        /// Checks if the pin button is currently enabled in the container
        /// </summary>
        /// <returns>True if enabled otherwise false</returns>
        public bool PinButtonEnabled()
        {
            bool enabled = GetButtonEnabledInternal(ButtonToGet.Pin);
            return enabled;
        }

        /// <summary>
        /// Clicks on the pin button in the container
        /// </summary>
        public void PinButtonClick()
        {
            GUI.Log("Single Left click on the pin button in the " + Identity.Description, LogItemType.Action);
			if (!PinButtonVisible())
            {
                throw new Exception("Pin button is not visible in the " + Description);
            }
            if (!PinButtonEnabled())
            {
                throw new Exception("Pin button is not enabled in the " + Description);
            }
            ClickButtonInternal(ButtonToGet.Pin);
        }

        /// <summary>
        /// Checks if the window button is currently visible in the container
        /// </summary>
        /// <returns>True if visible otherwise false</returns>
        public bool WindowButtonVisible()
        {
            bool visible = GetButtonVisibleInternal(ButtonToGet.Pin);
            return visible;
        }

        /// <summary>
        /// Checks if the window button is currently enabled in the container
        /// </summary>
        /// <returns>True if enabled otherwise false</returns>
        public bool WindowButtonEnabled()
        {
            bool enabled = GetButtonEnabledInternal(ButtonToGet.Pin);
            return enabled;
        }

        /// <summary>
        /// Clicks on the window button in the container
        /// </summary>
        public void WindowButtonClick()
        {
            GUI.Log("Single Left click on the window button in the " + Identity.Description, LogItemType.Action);
            if (!WindowButtonVisible())
            {
                throw new Exception("Window button is not visible in the " + Description);
            }
            if (!WindowButtonEnabled())
            {
                throw new Exception("Window button is not enabled in the " + Description);
            }
            ClickButtonInternal(ButtonToGet.Window);
        }

        internal enum ButtonToGet
        {
            Close,
            Window,
            Pin,
        }
        internal bool GetButtonEnabledInternal(ButtonToGet button)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystem", MemberTypes.Property);
            switch (button)
            {
                case ButtonToGet.Close:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "CloseButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "enabled", MemberTypes.Field);
                    break;
                case ButtonToGet.Window:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "WindowButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "enabled", MemberTypes.Field);
                    break;
                case ButtonToGet.Pin:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "PinButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "enabled", MemberTypes.Field);
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

        internal bool GetButtonVisibleInternal(ButtonToGet button)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystem", MemberTypes.Property);
            switch (button)
            {
                case ButtonToGet.Close:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "CloseButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "visible", MemberTypes.Field);
                    break;
                case ButtonToGet.Window:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "WindowButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "visible", MemberTypes.Field);
                    break;
                case ButtonToGet.Pin:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "PinButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "visible", MemberTypes.Field);
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

        internal Rectangle GetButtonBoundsInternal(ButtonToGet button)
        {
            Rectangle buttonBounds = new Rectangle();
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LayoutSystem", MemberTypes.Property);
            switch (button)
            {
                case ButtonToGet.Close:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "CloseButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                case ButtonToGet.Window:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "WindowButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                case ButtonToGet.Pin:
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "PinButton", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "bounds", MemberTypes.Field);
                    break;
                default:
                    throw new Exception("Unsupported button bounds: " + button.ToString());
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

        internal void ClickButtonInternal(ButtonToGet button)
        {
            Rectangle buttonBounds = GetButtonBoundsInternal(button);
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
    }
}
