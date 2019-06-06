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
using System.ComponentModel;
using APE.Communication;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate focusable controls
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public class GUIFocusableObject : GUIObject
    {
        /// <summary>
        /// Constructor used for form controls
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIFocusableObject(string descriptionOfControl, params Identifier[] identParams)
            : base(descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIFocusableObject(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Whether the control currently has focus
        /// </summary>
        public bool HasFocus
        {
            get
            {
                return Input.HasFocus(Identity.ParentHandle, Identity.Handle);
            }
        }

        /// <summary>
        /// Sets focus to the control if it or a child of it doesn't currently have focus
        /// </summary>
        public void SetFocus()
        {
            Input.SetFocus(Identity.ParentHandle, Identity.Handle, Identity.Description);
        }

        /// <summary>
        /// Send the specified text to the currently focused control, logging the action
        /// </summary>
        /// <param name="textToSend">The text to send to the control</param>
        protected void SendKeys(string textToSend)
        {
            GUI.Log("Type [" + Unescape(textToSend) + "] into the " + Identity.Description, LogItemType.Action);
            SendKeysInternal(textToSend);
        }

        /// <summary>
        /// Send the specified text to the currently focused control
        /// </summary>
        /// <param name="textToSend">The text to send to the control</param>
        protected void SendKeysInternal(string textToSend)
        {
            Input.SendKeys(this, textToSend);
        }

        internal string Unescape(string escapedText)
        {
            if (escapedText == null)
            {
                return null;
            }
            string unescapedText = escapedText;
            unescapedText = unescapedText.Replace("{{}", "{");
            unescapedText = unescapedText.Replace("{}}", "}");
            unescapedText = unescapedText.Replace("{(}", "(");
            unescapedText = unescapedText.Replace("{)}", ")");
            unescapedText = unescapedText.Replace("{{}", "{");
            unescapedText = unescapedText.Replace("{+}", "+");
            unescapedText = unescapedText.Replace("{^}", "^");
            unescapedText = unescapedText.Replace("{%}", "%");
            unescapedText = unescapedText.Replace("{~}", "~");
            unescapedText = unescapedText.Replace("{[}", "[");
            unescapedText = unescapedText.Replace("{]}", "]");
            return unescapedText;
        }
    }
}
