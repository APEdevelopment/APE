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
            string CurrentText;

            Input.Block(Identity.ParentHandle, Identity.Handle);
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

                GUITextBox textbox = new GUITextBox(m_ParentForm, m_DescriptionOfControl + " textbox", new Identifier(Identifiers.Handle, textboxHandle));

                CurrentText = GUI.m_APE.GetWindowTextViaWindowMessage(textboxHandle);

                if (CurrentText != "")
                {
                    textbox.MouseSingleClick(MouseButton.Left);
                    //textbox.MouseDoubleClick(MouseButton.Left);

                    //Select everything in the textbox
                    base.SendKeys("{HOME}+{END}");

                    string selectedText;

                    //wait for .selectedText to = Text
                    timer = Stopwatch.StartNew();
                    do
                    {
                        //Get the selectedText property
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, textboxHandle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedText", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        selectedText = GUI.m_APE.GetValueFromMessage();

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to select all the text in the TextBox");
                        }

                        Thread.Sleep(15);
                    }
                    while (CurrentText != selectedText);
                    timer.Stop();
                }

                //Send first 2 characters
                base.SendKeys(text.Substring(0, 2));
                GUI.Log("Wait for the generic walker popup to appear", LogItemTypeEnum.Action);

                //Wait for popup
                timer = Stopwatch.StartNew();
                bool isDropped = false;
                do
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
                }
                while (!isDropped);
                timer.Stop();

                //Send rest of characters
                base.SendKeys(text.Substring(2));

                //wait for .Text to == text
                timer = Stopwatch.StartNew();
                do
                {
                    CurrentText = GUI.m_APE.GetWindowTextViaWindowMessage(textboxHandle);

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to set the text of the TextBox");
                    }

                    Thread.Sleep(15);
                }
                while (CurrentText != text);
                timer.Stop();

                GUI.Log("Press Enter to set the value", LogItemTypeEnum.Action);
                base.SendKeysInternal("{Enter}");
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
