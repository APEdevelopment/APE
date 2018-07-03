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
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using APE.Capture;
using APE.Communication;
using System.Threading;
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

        private int GetCurrentDropAfter()
        {
            if (Identity.TypeNameSpace == "LzGenericWalker" || Identity.TypeName == "AxGenericWalker")
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Driver", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "WalkerStateDropAfter", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call;
                int dropAfter = GUI.m_APE.GetValueFromMessage();

                if (dropAfter == 0)
                {
                    return 999999999;   //Don't yet have a drop after value so set it to a really large value
                }
                else
                {
                    return dropAfter;
                }
            }
            else
            {
                throw GUI.ApeException("Not supported");
            }
        }

        private string GetCurrentState()
        {
            if (Identity.TypeNameSpace == "LzGenericWalker")
            {
                throw GUI.ApeException("Not supported");
            }
            else
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Driver", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "State", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ToString", MemberTypes.Method);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call;
                string state = GUI.m_APE.GetValueFromMessage();
                return state;
            }
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
                GUITextBox textbox;
                if (Identity.TypeNameSpace == "LzGenericWalker" || Identity.TypeName == "AxGenericWalker")
                {
                    if (Identity.TechnologyType == "Windows Forms (WinForms)" && Identity.TypeName != "AxGenericWalker")
                    {
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "txtText", MemberTypes.Field);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        IntPtr textboxHandle = GUI.m_APE.GetValueFromMessage();
                        textbox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, textboxHandle));
                    }
                    else
                    {
                        textbox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Name, "txtText"), new Identifier(Identifiers.ChildOf, this));
                    }
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
                    IntPtr textboxHandle = GUI.m_APE.GetValueFromMessage();
                    textbox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, textboxHandle));
                }               

                currentText = textbox.Text;
                string unescapedText = Unescape(text);

                if (unescapedText == currentText)
                {
                    GUI.Log("Ensure " + Identity.Description + " is set to " + unescapedText, LogItemType.Action);
                }
                else
                {
                    if (currentText != "")
                    {
                        textbox.SingleClick(MouseButton.Left);

                        //Select everything in the textbox
                        base.SendKeys("{HOME}+{END}");

                        //wait for SelectedText == Text
                        timer = Stopwatch.StartNew();
                        while (true)
                        {
                            string selectedText = textbox.SelectedText();
                            if (currentText == selectedText)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw GUI.ApeException("Failed to select all the text of the " + Description);
                            }

                            Thread.Sleep(15);
                        }
                    }

                    // Get an array of each character token for example a or {+}
                    string[] tokens = Tokenise(text);

                    switch (tokens.Length)
                    {
                        case 0:
                            break;
                        default:
                            int token = 0;
                            int searchAfter;

                            while (true)
                            {
                                if (Identity.TypeNameSpace == "LzGenericWalker" || Identity.TypeName == "AxGenericWalker")
                                {
                                    searchAfter = GetCurrentDropAfter();
                                }
                                else
                                {
                                    string state = GetCurrentState();

                                    //PrefixState has 4 states: Unknown, Incomplete, Valid, Invalid
                                    //The last 2 states cause the drop down to appear
                                    if (token == 0 || state == "Unknown" || state == "Incomplete")
                                    {
                                        searchAfter = token + 1;
                                    }
                                    else
                                    {
                                        searchAfter = token;
                                    }
                                }

                                if (token < searchAfter)
                                {
                                    GUI.Log("Type [" + unescapedText.Substring(token, 1) + "] into the " + Identity.Description, LogItemType.Action);
                                    base.SendKeysInternal(tokens[token]);
                                    tokens[token] = "";
                                }
                                else
                                {
                                    break;
                                }

                                token++;

                                if (token == tokens.Length)
                                {
                                    break;
                                }
                            }

                            if (token == tokens.Length)
                            {
                                break;
                            }

                            //Wait for popup
                            GUI.Log("Wait for the generic walker popup to appear", LogItemType.Action);
                            timer = Stopwatch.StartNew();
                            bool isDropped = false;
                            while (true)
                            {
                                //Get the state of the popup control
                                if (Identity.TypeNameSpace == "LzGenericWalker" || Identity.TypeName == "AxGenericWalker")
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
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw GUI.ApeException("Failed to find the " + Description + " dropdown");
                                }

                                Thread.Sleep(15);
                            }

                            //Send rest of characters
                            if (tokens.Length > searchAfter)
                            {
                                GUI.Log("Type [" + unescapedText.Substring(searchAfter) + "] into the " + Identity.Description, LogItemType.Action);
                                base.SendKeysInternal(string.Join("", tokens));
                            }
                            break;
                    }

                    //wait for current text == unescaped text
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
                            throw GUI.ApeException("Failed to set the text of the " + Description);
                        }

                        Thread.Sleep(15);
                    }

                    if (Identity.TypeNameSpace == "LzGenericWalker" && Identity.TechnologyType == "Windows Forms (WinForms)")
                    {
                        // Add event handler
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageAddGenericWalkerSelectedHandler(DataStores.Store0, Identity.ParentHandle);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                    }

                    bool ok = false;
                    try
                    {
                        GUI.Log("Press Enter to set the value", LogItemType.Action);
                        base.SendKeysInternal("{Enter}");
                        ok = true;
                    }
                    finally
                    {
                        if (Identity.TypeNameSpace == "LzGenericWalker" && Identity.TechnologyType == "Windows Forms (WinForms)")
                        {
                            if (ok)
                            {
                                //Wait for the event handler then remove it
                                GUI.m_APE.AddFirstMessageWaitForAndRemoveGenericWalkerSelectedHandler();
                                GUI.m_APE.SendMessages(EventSet.APE);
                                GUI.m_APE.WaitForMessages(EventSet.APE);
                            }
                            else
                            {
                                //Remove the event handler
                                GUI.m_APE.AddFirstMessageRemoveGenericWalkerSelectedHandler();
                                GUI.m_APE.SendMessages(EventSet.APE);
                                GUI.m_APE.WaitForMessages(EventSet.APE);
                            }
                        }
                    }
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
