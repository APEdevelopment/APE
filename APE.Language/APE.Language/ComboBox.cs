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
using System.Reflection;
using System.Diagnostics;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.ComboBox
    /// LatentZero.Utility.Controls.GUIComboBox
    /// </summary>
    public sealed class GUIComboBox : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIComboBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Checks if the specified item exists in the combobox
        /// </summary>
        /// <param name="item">The item to check if it exists</param>
        /// <returns></returns>
        public bool ItemExists(string item)
        {
            int Index = ItemIndex(item);

            if (Index == NM.CB_ERR)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Selects the specified item in the combobox
        /// </summary>
        /// <param name="item">The item to select</param>
        public void ItemSelect(string item)
        {
            Stopwatch timer;

            //Check if its already selected
            if (this.Text == item)
            {
                GUI.Log("Ensure " + m_DescriptionOfControl + " is set to " + item, LogItemType.Action);
                return;
            }
            
            GUI.Log("Select [" + item + "] from " + m_DescriptionOfControl, LogItemType.Action);

            //Get the style
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "DropDownStyle", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "DroppedDown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string Style = GUI.m_APE.GetValueFromMessage();
            bool DroppedDown = GUI.m_APE.GetValueFromMessage();

            IntPtr ListBox = IntPtr.Zero;

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                if (Style == "Simple")
                {
                    //get the Simple mode listbox child window
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "childListBox", MemberTypes.Field);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    ListBox = (IntPtr)GUI.m_APE.GetValueFromMessage();
                }
                else
                {
                    if (!DroppedDown)
                    {
                        //show the dropdown
                        base.MouseSingleClickInternal(Width - 5, -1, MouseButton.Left, MouseKeyModifier.None);
                    }
                    
                    //find the dropdown
                    Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "dropDownHandle", MemberTypes.Field);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    dynamic DroppedDownHandle = GUI.m_APE.GetValueFromMessage();

                    if (DroppedDownHandle == null)
                    {
                        throw new Exception("Failed to find the ComboBox dropdown");
                    }

                    ListBox = DroppedDownHandle;
                }

                //locate the item
                int Index = ItemIndex(item);
                if (Index == NM.CB_ERR)
                {
                    throw new Exception("Failed to find the ComboBox item");
                }

                IntPtr MessageResult;
                IntPtr SendResult;
                NM.tagRect ItemRect = new NM.tagRect();

                //Locate the rect of the item
                SendResult = NM.SendMessageTimeout(ListBox, NM.ListBoxMessages.LB_GETITEMRECT, new IntPtr(Index), ref ItemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                if (SendResult == IntPtr.Zero || unchecked((int)MessageResult.ToInt64()) == NM.LB_ERR)  //Failed
                {
                    throw new Exception("Failed to access the combobox");
                }

                NM.tagRect ClientRect;
                NM.GetClientRect(ListBox, out ClientRect);

                //scroll the item into view if needed
                if (((ItemRect.bottom - ItemRect.top) / 2) + ItemRect.top > ClientRect.bottom || ((ItemRect.bottom - ItemRect.top) / 2) + ItemRect.top < ClientRect.top)
                {
                    SendResult = NM.SendMessageTimeout(ListBox, NM.ListBoxMessages.LB_SETTOPINDEX, new IntPtr(Index), ref ItemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                    if (SendResult == IntPtr.Zero || unchecked((int)MessageResult.ToInt64()) == NM.LB_ERR)  //Failed
                    {
                        throw new Exception("Failed to access the combobox");
                    }

                    //Locate the rect of the item
                    SendResult = NM.SendMessageTimeout(ListBox, NM.ListBoxMessages.LB_GETITEMRECT, new IntPtr(Index), ref ItemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                    if (SendResult == IntPtr.Zero || unchecked((int)MessageResult.ToInt64()) == NM.LB_ERR)  //Failed
                    {
                        throw new Exception("Failed to access the combobox");
                    }
                }

                //click the item
                IntPtr temp = Identity.Handle;
                try
                {
                    Identity.Handle = ListBox;
                    base.MouseSingleClickInternal(-1, ((ItemRect.bottom - ItemRect.top) / 2) + ItemRect.top, MouseButton.Left, MouseKeyModifier.None);
                }
                finally
                {
                    Identity.Handle = temp;
                }

                //wait for .Text to == text
                string CurrentText;
                timer = Stopwatch.StartNew();
                do
                {
                    CurrentText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);

                    if (CurrentText == item)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to set the text of the ComboBox");
                    }

                    Thread.Sleep(15);
                }
                while (true);
                timer.Stop();
            }
            finally
            {
                Input.Unblock();
            }
        }

        /// <summary>
        /// Sets the text portion of the combobox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text portion of the combobox to</param>
        public void SetText(string text)
        {
            //Get the style
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "DropDownStyle", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string Style = GUI.m_APE.GetValueFromMessage();

            if (Style == "DropDownList")
            {
                throw new Exception("ComboBox Edit control is not editable");
            }

            //get the editbox child window
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "childEdit", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            IntPtr EditBox = (IntPtr)GUI.m_APE.GetValueFromMessage();

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                GUITextBox comboboxTextBox = new GUITextBox(m_ParentForm, m_DescriptionOfControl + " textbox", new Identifier(Identifiers.Handle, EditBox), new Identifier(Identifiers.TechnologyType, "Windows Native"));
                comboboxTextBox.SetText(text);
            }
            finally
            {
                Input.Unblock();
            }
        }

        private int ItemIndex(string Item)
        {
            IntPtr MessageResult;
            IntPtr SendResult;

            SendResult = NM.SendMessageTimeout(Identity.Handle, NM.ComboBoxMessages.CB_FINDSTRINGEXACT, new IntPtr(-1), Item, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
            if (SendResult == IntPtr.Zero) //Failed
            {
                throw new Exception("Failed to access the combobox");
            }
            else
            {
                return unchecked((int)MessageResult.ToInt64());
            }
        }
    }
}
