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
