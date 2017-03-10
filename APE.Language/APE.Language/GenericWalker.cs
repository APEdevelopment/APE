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
using System.Collections;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Capstone.Controls.Walker.WalkerControl
    /// LzGenericWalker.GenericWalker
    /// </summary>
    public sealed class GUIGenericWalker : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIGenericWalker(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the text of the generic walker control
        /// </summary>
        /// <param name="text">The text to set the control to</param>
        public void SetText(string text)
        {
            Stopwatch timer;
            string currentText;

            Input.Block();
            try
            {
                IntPtr textboxHandle;

                if (Identity.TypeNameSpace == "LzGenericWalker")
                {
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "txtText", MemberTypes.Field);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    textboxHandle = GUI.m_APE.GetValueFromMessage();
                }
                else
                {
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "WalkerTextBox", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    textboxHandle = GUI.m_APE.GetValueFromMessage();
                }               

                GUITextBox textbox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, textboxHandle));

                currentText = textbox.Text;

                if (text == currentText)
                {
                    GUI.Log("Ensure " + Identity.Description + " is set to " + text, LogItemType.Action);
                }
                else
                {
                    if (currentText != "")
                    {
                        textbox.SingleClick(MouseButton.Left);
                        //textbox.MouseDoubleClick(MouseButton.Left);

                        //Select everything in the textbox
                        base.SendKeys("{HOME}+{END}");

                        string selectedText;

                        //wait for .selectedText to = Text
                        timer = Stopwatch.StartNew();
                        while (true)
                        {
                            //Get the selectedText property
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, textboxHandle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedText", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            selectedText = GUI.m_APE.GetValueFromMessage();

                            if (currentText == selectedText)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw new Exception("Failed to select all the text of the " + Description);
                            }

                            Thread.Sleep(15);
                        }
                    }

                    // Get an array of each character token for example a or {+}
                    string[] tokens = Tokenise(text);
                    string unescapedText = Unescape(text);

                    switch (tokens.Length)
                    {
                        case 0:
                            break;
                        case 1:
                            GUI.Log("Type [" + unescapedText.Substring(0, 1) + "] into the " + Identity.Description, LogItemType.Action);
                            base.SendKeysInternal(tokens[0]);
                            break;
                        default:
                            //Send first 2 characters
                            GUI.Log("Type [" + unescapedText.Substring(0, 2) + "] into the " + Identity.Description, LogItemType.Action);
                            base.SendKeysInternal(tokens[0] + tokens[1]);

                            //Wait for popup
                            GUI.Log("Wait for the generic walker popup to appear", LogItemType.Action);
                            timer = Stopwatch.StartNew();
                            bool isDropped = false;
                            while (true)
                            {
                                //Get the state of the popup control
                                if (Identity.TypeNameSpace == "LzGenericWalker")
                                {
                                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsDropped", MemberTypes.Property);
                                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                                    GUI.m_APE.SendMessages(EventSet.APE);
                                    GUI.m_APE.WaitForMessages(EventSet.APE);
                                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                                    isDropped = GUI.m_APE.GetValueFromMessage();
                                }
                                else
                                {
                                    string PopupState;
                                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "PopupState", MemberTypes.Property);
                                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                                    GUI.m_APE.SendMessages(EventSet.APE);
                                    GUI.m_APE.WaitForMessages(EventSet.APE);
                                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                                    PopupState = GUI.m_APE.GetValueFromMessage();
                                    if (PopupState == "Open")
                                    {
                                        isDropped = true;
                                    }
                                }

                                if (isDropped)
                                {
                                    if (Identity.TypeNameSpace == "LzGenericWalker")
                                    {
                                        WaitForTimer("tmrResize");
                                    }
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw new Exception("Failed to find the " + Description + " dropdown");
                                }

                                Thread.Sleep(15);
                            }

                            //Send rest of characters
                            if (tokens.Length > 2)
                            {
                                // Clear the first 2 characters as they have already been sent
                                tokens[0] = "";
                                tokens[1] = "";
                                GUI.Log("Type [" + unescapedText.Substring(2) + "] into the " + Identity.Description, LogItemType.Action);
                                base.SendKeysInternal(string.Join("", tokens));
                            }
                            break;
                    }

                    //wait for .Text to == text
                    timer = Stopwatch.StartNew();
                    while (true)
                    {
                        currentText = textbox.Text;

                        if (currentText == unescapedText)
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to set the text of the " + Description);
                        }

                        Thread.Sleep(15);
                    }

                    GUI.Log("Press Enter to set the value", LogItemType.Action);
                    base.SendKeysInternal("{Enter}");

                    if (Identity.TypeNameSpace == "LzGenericWalker")
                    {
                        WaitForTimer("tmrDelayedEvent");
                    }
                }
            }
            finally
            {
                Input.Unblock();
            }
        }

        private void WaitForTimer(string timerName)
        {
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                // Wait for the tmrDelayedEvent to be disabled
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, timerName, MemberTypes.Field);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Enabled", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                bool timerEnabled = GUI.m_APE.GetValueFromMessage();

                if (!timerEnabled)
                {
                    GUI.WaitForInputIdle(this);
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception(timerName + " failed to become disabled for the " + Description);
                }

                Thread.Sleep(15);
            }
        }

        private string[] Tokenise(string escapedText)
        {
            ArrayList tokens = new ArrayList(escapedText.Length);
            string character;
            for (int characterIndex = 0; characterIndex < escapedText.Length; characterIndex++)
            {
                if (escapedText.Substring(characterIndex, 1) == "{")
                {
                    if (escapedText.Substring(characterIndex + 1, 1) == "}")
                    {
                        character = escapedText.Substring(characterIndex, 3);
                        characterIndex = characterIndex + 2;
                    }
                    else
                    {
                        int endIndex = escapedText.IndexOf("}", characterIndex);
                        character = escapedText.Substring(characterIndex, endIndex - characterIndex + 1);
                        characterIndex = endIndex;
                    }
                }
                else
                {
                    character = escapedText.Substring(characterIndex, 1);
                }

                tokens.Add(character);
            }

            tokens.TrimToSize();
            return (string[])tokens.ToArray(typeof(string));
        }

        private string Unescape(string escapedText)
        {
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
