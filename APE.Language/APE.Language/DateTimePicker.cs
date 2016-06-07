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

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.DateTimePicker
    /// </summary>
    public sealed class GUIDateTimePicker : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIDateTimePicker(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the text portion of the datetimepicker to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text portion of the datetimepicker to</param>
        public void SetText(string text)
        {
            //TODO with really narrow picker it may click the button...
            GUITextBox dateTimePickerTextBox = new GUITextBox(m_ParentForm, m_DescriptionOfControl + " textbox", new Identifier(Identifiers.Handle, Identity.Handle));
            dateTimePickerTextBox.SetText(text);
        }
    }
}
