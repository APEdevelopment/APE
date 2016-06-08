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
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using NM = APE.Native.NativeMethods;
using APE.Communication;
using System.Reflection;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Capstone.Controls.TextValueWithTypeControl.TextValueWithTypeControl
    /// </summary>
    public sealed class GUITextValueWithTypeControl : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITextValueWithTypeControl(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the type portion of the TextValueWithTypeControl to the specified text
        /// </summary>
        /// <param name="typeText">The text to set the type portion of the TextValueWithTypeControl to</param>
        public void SetType(string typeText)
        {
            //click on the right side of the control to display the popup list form
            GUI.Log("Left click on the right side of the text value with type control to display the popup list", LogItemTypeEnum.Action);
            base.MouseSingleClickInternal(Width - 5, 5, MouseButton.Left, MouseKeyModifier.None);

            //find the form
            GUIForm popupListForm = new GUIForm("Popup list form", new Identifier(Identifiers.Name, "PopupListForm"));

            //find the popuplist
            GUIPopupList popupList = new GUIPopupList(popupListForm, m_DescriptionOfControl + " popup List", new Identifier(Identifiers.Name, "popupList1"));
            //Select the type in the popup list
            popupList.Select(typeText);
        }

        /// <summary>
        /// Sets the text portion of the TextValueWithTypeControl to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text portion of the TextValueWithTypeControl to</param>
        public void SetText(string text)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "TextBox", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            IntPtr textboxHandle = GUI.m_APE.GetValueFromMessage();

            GUITextBox textBox = new GUITextBox(m_ParentForm, m_DescriptionOfControl + " textbox", new Identifier(Identifiers.Handle, textboxHandle));
            // Select the item
            textBox.SetText(text);
        }
    }
}
