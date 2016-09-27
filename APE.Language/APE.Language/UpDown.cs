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
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using APE.Capture;
using APE.Communication;
using System.Threading;
using System.Drawing.Imaging;
using System.Security.Principal;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.NumericUpDown
    /// </summary>
    public sealed class GUIUpDown : GUIFocusableObject, IGUITextBox
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIUpDown(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the text portion of the up down control to the specified text
        /// </summary>
        /// <param name="text">The text to set the control to</param>
        public void SetText(string text)
        {
            //Get the selectedText property
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "upDownEdit", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic UpDownEditControl = GUI.m_APE.GetValueFromMessage();

            if (UpDownEditControl == null)
            {
                throw new Exception("Failed to find the updown edit control");
            }

            IntPtr UpDownEditHandle = UpDownEditControl;

            GUITextBox UpDownEdit = new GUITextBox(m_ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, UpDownEditHandle));
            
            UpDownEdit.SetText(text);
        }

        /// <summary>
        /// Clicks on the up portion of the up down control
        /// </summary>
        /// <param name="Button">The button with which to click</param>
        public void SingleClickUp(MouseButton Button)
        {
            //Get the selectedText property
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "upDownButtons", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic UpDownButtonsControl = GUI.m_APE.GetValueFromMessage();

            if (UpDownButtonsControl == null)
            {
                throw new Exception("Failed to find the updown buttons control");
            }

            IntPtr UpDownButtonHandle = UpDownButtonsControl;

            GUIButton UpDownButton = new GUIButton(m_ParentForm, Identity.Description + " up button", new Identifier(Identifiers.Handle, UpDownButtonHandle));

            UpDownButton.MouseSingleClick(UpDownButton.Width / 2, (UpDownButton.Height / 4) * 1, Button);
        }

        /// <summary>
        /// Clicks on the down portion of the up down control
        /// </summary>
        /// <param name="Button">>The button with which to click</param>
        public void SingleClickDown(MouseButton Button)
        {
            //Get the selectedText property
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "upDownButtons", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic UpDownButtonsControl = GUI.m_APE.GetValueFromMessage();

            if (UpDownButtonsControl == null)
            {
                throw new Exception("Failed to find the updown buttons control");
            }

            IntPtr UpDownButtonHandle = UpDownButtonsControl;

            GUIButton UpDownButton = new GUIButton(m_ParentForm, Identity.Description + " down button", new Identifier(Identifiers.Handle, UpDownButtonHandle));

            UpDownButton.MouseSingleClick(UpDownButton.Width / 2, (UpDownButton.Height / 4) * 3, Button);
        }
    }
}
