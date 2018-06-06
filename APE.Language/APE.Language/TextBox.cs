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
    /// System.Windows.Forms.TextBox
    /// </summary>
    public sealed class GUITextBox : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITextBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the text of the textbox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text of the textbox to</param>
        public void SetText(string text)
        {
            SetText(text, null);
        }

        /// <summary>
        /// Sets the text of the textbox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text of the textbox to</param>
        /// <param name="submitKey">The key to press after setting the text, for instance {Enter}</param>
        public void SetText(string text, string submitKey)
        {
            Stopwatch timer;
            string currentText;

            currentText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);
            if (currentText == text)
            {
                GUI.Log("Ensure the " + Identity.Description + " is set to " + text, LogItemType.Action);
                return;
            }

            Input.Block();
            try
            {
                if (text != null)
                {
                    currentText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);

                    if (currentText != "")
                    {
                        base.SingleClick(MouseButton.Left);

                        //Select everything in the textbox
                        base.SendKeys("{HOME}+{END}");

                        string selectedText;

                        //wait for .selectedText to = Text
                        timer = Stopwatch.StartNew();
                        while (true)
                        {
                            selectedText = SelectedText();
                            if (currentText == selectedText)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw GUI.ApeException("Failed to select all the text in the TextBox");
                            }

                            Thread.Sleep(15);
                        }

                        if (text == "")
                        {
                            base.SendKeys("{BACKSPACE}");

                            //wait for .Text to == ""
                            timer = Stopwatch.StartNew();
                            while (true)
                            {
                                currentText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);
                                if (currentText == "")
                                {
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw GUI.ApeException("Failed to set the text of the TextBox");
                                }

                                Thread.Sleep(15);
                            }
                        }
                    }

                    if (text != "")
                    {
                        base.SendKeys(text);

                        //wait for .Text to == text
                        timer = Stopwatch.StartNew();
                        while (true)
                        {
                            currentText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);
                            if (currentText == text)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw GUI.ApeException("Failed to set the text of the TextBox");
                            }

                            Thread.Sleep(15);
                        }
                    }
                }

                if (!string.IsNullOrEmpty(submitKey))
                {
                    base.SendKeys(submitKey);
                }
            }
            catch when (Input.ResetInputFilter())
            {
                // Will never be reached as ResetInputFilter always returns false
            }
            finally
            {
                Input.Unblock();
            }
        }

        /// <summary>
        /// Gets the text which is currently selected
        /// </summary>
        /// <returns>The selected text</returns>
        public string SelectedText()
        {
            string selectedText;
            if (Identity.TechnologyType == "Windows Forms (WinForms)")
            {
                //Get the selectedText property
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedText", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                selectedText = GUI.m_APE.GetValueFromMessage();
            }
            else
            {
                IntPtr messageResult;
                IntPtr sendResult;
                int start = 0;
                int end = 0;

                string currentText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);
                sendResult = NM.SendMessageTimeout(Identity.Handle, NM.EM_GETSEL, ref start, ref end, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
                
                if (sendResult != IntPtr.Zero)  //Succeeded
                {
                    selectedText = currentText.Substring(start, end - start);
                }
                else
                {
                    throw GUI.ApeException("Failed to query the " + Description);
                }
            }
            return selectedText;
        }
    }
}
