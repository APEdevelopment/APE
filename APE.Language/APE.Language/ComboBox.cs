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
using System.Drawing;

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
        /// <returns>True if the item exists otherwise false</returns>
        public bool ItemExists(string item)
        {
            return ItemExists(item, CaseSensitivity.Sensitive);
        }

        /// <summary>
        /// Checks if the specified item exists in the combobox
        /// </summary>
        /// <param name="itemText">The item to check if it exists</param>
        /// <param name="caseSensitivity">Whether to include the case of the item in the comparison</param>
        /// <returns>True if the item exists otherwise false</returns>
        public bool ItemExists(string itemText, CaseSensitivity caseSensitivity)
        {
            int index = ItemIndex(itemText, caseSensitivity);
            if (index == NM.CB_ERR)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Gets the text of the item at the specified index
        /// </summary>
        /// <param name="itemIndex">Item index</param>
        /// <returns>The text of the item</returns>
        public string ItemText(int itemIndex)
        {
            IntPtr messageResult;
            IntPtr sendResult;
            string itemText;

            sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ComboBoxMessages.CB_GETLBTEXTLEN, new IntPtr(itemIndex), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || messageResult == new IntPtr(NM.CB_ERR)) //Failed
            {
                throw GUI.ApeException("Failed to access the " + Description);
            }
            
            int itemCharacterCount = messageResult.ToInt32();
            if (itemCharacterCount == 0)
            {
                return "";
            }

            itemText = new string(' ', itemCharacterCount);
            
            sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ComboBoxMessages.CB_GETLBTEXT, new IntPtr(itemIndex), itemText, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || messageResult == new IntPtr(NM.CB_ERR)) //Failed
            {
                throw GUI.ApeException("Failed to access the " + Description);
            }
            else
            {
                return itemText;
            }
        }

        private string GetComboBoxStyle(out dynamic droppedDown)
        {
            string style;
            switch (Identity.TechnologyType)
            {
                case "Windows Forms (WinForms)":
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "DropDownStyle", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "DroppedDown", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    style = GUI.m_APE.GetValueFromMessage();
                    droppedDown = GUI.m_APE.GetValueFromMessage();
                    break;
                default:
                    long windowStyle = this.Style;
                    if ((windowStyle & (long)NM.ComboBoxStyle.DropDownList) == (long)NM.ComboBoxStyle.DropDownList)
                    {
                        style = NM.ComboBoxStyle.DropDownList.ToString();
                    }
                    else if ((windowStyle & (long)NM.ComboBoxStyle.DropDown) == (long)NM.ComboBoxStyle.DropDown)
                    {
                        style = NM.ComboBoxStyle.DropDown.ToString();
                    }
                    else if ((windowStyle & (long)NM.ComboBoxStyle.Simple) == (long)NM.ComboBoxStyle.Simple)
                    {
                        style = NM.ComboBoxStyle.Simple.ToString();
                    }
                    else
                    {
                        throw GUI.ApeException("Unsupported combobox style for the " + Description);
                    }

                    IntPtr messageResult;
                    IntPtr sendResult;
                    sendResult = NM.SendMessageTimeout(Identity.Handle, NM.CB_GETDROPPEDSTATE, IntPtr.Zero, IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, (uint)GUI.GetTimeOut(), out messageResult);
                    if (sendResult != IntPtr.Zero)  //Succeeded
                    {
                        if (messageResult == IntPtr.Zero)
                        {
                            droppedDown = false;
                        }
                        else
                        {
                            droppedDown = true;
                        }
                    }
                    else
                    {
                        throw GUI.ApeException("Failed to query the " + Description);
                    }
                    break;
            }
            return style;
        }

        /// <summary>
        /// Selects the specified item in the combobox by clicking on it
        /// </summary>
        /// <param name="itemText">The item to select</param>
        public void SingleClickItem(string itemText)
        {
            SingleClickItem(itemText, CaseSensitivity.Sensitive);
        }

        /// <summary>
        /// Selects the specified item in the combobox by clicking on it
        /// </summary>
        /// <param name="itemText">The item to select</param>
        /// <param name="caseSensitivity">Whether to include the case of the item in the comparison</param>
        public void SingleClickItem(string itemText, CaseSensitivity caseSensitivity)
        {
            Stopwatch timer;

            //Check if its already selected
            switch (caseSensitivity)
            {
                case CaseSensitivity.Sensitive:
                    if (this.Text == itemText)
                    {
                        GUI.Log("Ensure " + Identity.Description + " is set to " + itemText, LogItemType.Action);
                        return;
                    }
                    break;
                case CaseSensitivity.Insensitive:
                    if (this.Text.ToLower() == itemText.ToLower())
                    {
                        GUI.Log("Ensure " + Identity.Description + " is set to " + itemText, LogItemType.Action);
                        return;
                    }
                    break;
                default:
                    throw GUI.ApeException("Unsupported CaseSensitivity value: " + caseSensitivity.ToString());
            }
            
            GUI.Log("Select [" + itemText + "] from " + Identity.Description, LogItemType.Action);

            //Get the style
            Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);
            string style = GetComboBoxStyle(out dynamic droppedDown);

            IntPtr listBox = IntPtr.Zero;
            NM.COMBOBOXINFO cbi = new NM.COMBOBOXINFO();

            Input.Block();
            try
            {
                if (style == "Simple")
                {
                    //get the Simple mode listbox child window
                    NM.GetComboBoxInfo(Identity.Handle, ref cbi);
                    listBox = cbi.hwndList;
                }
                else
                {
                    if (droppedDown == null)
                    {
                        throw GUI.ApeException("Failed to determine the dropdown state of the " + Description);
                    }

                    if (!droppedDown)
                    {
                        //show the dropdown
                        base.SingleClickInternal(Width - 5, -1, MouseButton.Left, MouseKeyModifier.None);
                    }
                    
                    //find the dropdown
                    Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);

                    NM.GetComboBoxInfo(Identity.Handle, ref cbi);
                    listBox = cbi.hwndList;
                    if (listBox == IntPtr.Zero)
                    {
                        throw GUI.ApeException("Failed to find the " + Description + " dropdown");
                    }
                }

                //locate the item
                int index = ItemIndex(itemText, caseSensitivity);
                if (index == NM.CB_ERR)
                {
                    throw GUI.ApeException("Failed to find the " + Description + " item");
                }

                NM.tagRect ClientRect;
                NM.GetClientRect(listBox, out ClientRect);

                //Locate the rectangle of the item
                Rectangle itemRectangle = GetItemRectangle(listBox, index);

                //scroll the item into view if needed and then locate the rectangle
                if ((itemRectangle.Height / 2) + itemRectangle.Top > ClientRect.bottom || (itemRectangle.Height / 2) + itemRectangle.Top < ClientRect.top)
                {
                    SetTopIndex(listBox, index);
                    itemRectangle = GetItemRectangle(listBox, index);
                }

                //click the item
                GUIForm comboBoxDropdown = new GUIForm(ParentForm, Description + " dropdown", new Identifier(Identifiers.Handle, listBox), new Identifier(Identifiers.TechnologyType, "Windows Native"));
                comboBoxDropdown.SingleClickInternal(-1, (itemRectangle.Height / 2) + itemRectangle.Top, MouseButton.Left, MouseKeyModifier.None);
                
                //wait for .Text to == text
                bool selected = false;
                timer = Stopwatch.StartNew();
                
                while (true)
                {
                    switch (caseSensitivity)
                    {
                        case CaseSensitivity.Sensitive:
                            if (this.Text == itemText)
                            {
                                selected = true;
                            }
                            break;
                        case CaseSensitivity.Insensitive:
                            if (this.Text.ToLower() == itemText.ToLower())
                            {
                                selected = true;
                            }
                            break;
                        default:
                            throw GUI.ApeException("Unsupported CaseSensitivity value: " + caseSensitivity.ToString());
                    }

                    if (selected)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw GUI.ApeException("Failed to select the item in the " + Description);
                    }

                    Thread.Sleep(15);
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
        /// Sets the text portion of the combobox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text portion of the combobox to</param>
        public void SetText(string text)
        {
            //Get the style
            string style = GetComboBoxStyle(out dynamic droppedDown);
            if (style == "DropDownList")
            {
                throw GUI.ApeException("Style of the " + Description + " is not an editable style");
            }

            //get the editbox child window
            NM.COMBOBOXINFO cbi = new NM.COMBOBOXINFO();
            NM.GetComboBoxInfo(Identity.Handle, ref cbi);
            IntPtr EditBox = cbi.hwndEdit;
            if (EditBox == IntPtr.Zero)
            {
                throw GUI.ApeException("Failed to find the " + Description + " textbox");
            }

            Input.Block();
            try
            {
                GUITextBox comboboxTextBox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, EditBox), new Identifier(Identifiers.TechnologyType, "Windows Native"));
                comboboxTextBox.SetText(text);
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

        private int ItemIndex(string itemText, CaseSensitivity caseSensitivity)
        {
            IntPtr messageResult;
            IntPtr sendResult;

            int startIndex = -1;
            while (true)
            {
                sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ComboBoxMessages.CB_FINDSTRINGEXACT, new IntPtr(startIndex), itemText, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
                if (sendResult == IntPtr.Zero)
                {
                    throw GUI.ApeException("Failed to access the " + Description);
                }
                else if (messageResult == new IntPtr(NM.CB_ERR))
                {
                    return NM.CB_ERR;
                }
                else
                {
                    int index = unchecked((int)messageResult.ToInt64());

                    // looped around through 0
                    if (index < startIndex)
                    {
                        return NM.CB_ERR;
                    }

                    switch (caseSensitivity)
                    {
                        case CaseSensitivity.Insensitive:
                            return index;
                        case CaseSensitivity.Sensitive:
                            string foundItemText = ItemText(index);
                            if (foundItemText == itemText)
                            {
                                return index;
                            }
                            break;
                        default:
                            throw GUI.ApeException("Unsupported CaseSensitivity value: " + caseSensitivity.ToString());
                    }
                    startIndex = index + 1;
                }
            }
        }

        private Rectangle GetItemRectangle(IntPtr listBox, int itemIndex)
        {
            IntPtr messageResult;
            IntPtr sendResult;
            NM.tagRect itemRect = new NM.tagRect();

            //Locate the rect of the item
            sendResult = NM.SendMessageTimeout(listBox, NM.ListBoxMessages.LB_GETITEMRECT, new IntPtr(itemIndex), ref itemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || unchecked((int)messageResult.ToInt64()) == NM.LB_ERR)  //Failed
            {
                if (sendResult == IntPtr.Zero)
                {
                    throw GUI.ApeException("Failed to access the " + Description + " to locate the rectangle containing the item");
                }
                else
                {
                    throw GUI.ApeException("Failed to locate the rectangle containing the item in the " + Description);
                }
            }
            return new Rectangle(itemRect.left, itemRect.top, itemRect.right - itemRect.left, itemRect.bottom - itemRect.top);
        }

        private void SetTopIndex(IntPtr listBox, int itemIndex)
        {
            IntPtr messageResult;
            IntPtr sendResult;

            sendResult = NM.SendMessageTimeout(listBox, NM.ListBoxMessages.LB_SETTOPINDEX, new IntPtr(itemIndex), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || unchecked((int)messageResult.ToInt64()) == NM.LB_ERR)  //Failed
            {
                if (sendResult == IntPtr.Zero)
                {
                    throw GUI.ApeException("Failed to access the " + Description + " to scroll the item");
                }
                else
                {
                    throw GUI.ApeException("Failed to scroll the item in the " + Description);
                }
            }
            Input.WaitForInputIdle(this.Handle, (uint)GUI.GetTimeOut());
        }
    }
}
