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
    /// System.Windows.Forms.ListBox
    /// </summary>
    public sealed class GUIListBox : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIListBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Checks if the specified item exists in the listbox
        /// </summary>
        /// <param name="itemText">The item to check if it exists</param>
        /// <returns>Whether item exists</returns>
        public bool ItemExists(string itemText)
        {
            return ItemExists(itemText, CaseSensitivity.Sensitive);
        }

        /// <summary>
        /// Checks if the specified item exists in the listbox
        /// </summary>
        /// <param name="itemText">The item to check if it exists</param>
        /// <param name="caseSensitivity">Whether to include the case of the item in the comparison</param>
        /// <returns>Whether item exists</returns>
        public bool ItemExists(string itemText, CaseSensitivity caseSensitivity)
        {
            int index = ItemIndex(itemText, caseSensitivity);
            if (index == NM.LB_ERR)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Returns the number of items in the listbox
        /// </summary>
        /// <returns>The number of items</returns>
        public int ItemCount()
        {
            IntPtr sendResult;
            IntPtr messageResult;
            int itemCount;

            sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_GETCOUNT, IntPtr.Zero, IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || messageResult.ToInt64() == NM.LB_ERR)  //Failed
            {
                throw GUI.ApeException("Failed to get the number of items in the " + Description);
            }
            itemCount = messageResult.ToInt32();
            return itemCount;
        }

        internal int ItemIndex(string itemText, CaseSensitivity caseSensitivity)
        {
            //Get the index
            IntPtr sendResult;
            IntPtr messageResult;
            int itemIndex;

            int startIndex = -1;
            while (true)
            {
                sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_FINDSTRINGEXACT, new IntPtr(startIndex), itemText, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
                if (sendResult == IntPtr.Zero)  //Failed
                {
                    throw GUI.ApeException("Failed to find query the " + Description);
                }
                itemIndex = messageResult.ToInt32();

                switch (caseSensitivity)
                {
                    case CaseSensitivity.Insensitive:
                        return itemIndex;
                    case CaseSensitivity.Sensitive:
                        string foundItemText = ItemText(itemIndex);
                        if (foundItemText == itemText)
                        {
                            return itemIndex;
                        }
                        break;
                    default:
                        throw GUI.ApeException("Unsupported CaseSensitivity value: " + caseSensitivity.ToString());
                }
                startIndex = itemIndex + 1;
            }
        }

        /// <summary>
        /// Returns the text of the item at the specified index
        /// </summary>
        /// <param name="itemIndex">The index of the item</param>
        /// <returns>The text of the item</returns>
        public string ItemText(int itemIndex)
        {
            IntPtr sendResult;
            IntPtr messageResult;
            string itemText = null;

            //Get the length of the item text
            sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_GETTEXT, new IntPtr(itemIndex), itemText, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            int itemCharacterCount = messageResult.ToInt32();
            if (itemCharacterCount > 0)
            {
                itemText = new string(' ', itemCharacterCount);
            }

            //Get the item text
            sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_GETTEXT, new IntPtr(itemIndex), itemText, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || messageResult.ToInt64() == NM.LB_ERR || messageResult.ToInt64() != itemCharacterCount)  //Failed
            {
                throw GUI.ApeException("Failed to get the item text of the " + Description);
            }
            return itemText;
        }

        /// <summary>
        /// Selects the specified item in the listbox by clicking on it
        /// </summary>
        /// <param name="itemText">The item to select</param>
        public void SingleClickItem(string itemText)
        {
            SingleClickItem(itemText, CaseSensitivity.Sensitive);
        }

        /// <summary>
        /// Selects the specified item in the listbox by clicking on it
        /// </summary>
        /// <param name="itemText">The item to select</param>
        /// <param name="caseSensitivity">Whether to include the case of the item in the comparison</param>
        public void SingleClickItem(string itemText, CaseSensitivity caseSensitivity)
        {
            GUI.Log("Single Left click on the item " + itemText + " from the " + Identity.Description, LogItemType.Action);
            SingleClickItemInternal(itemText, caseSensitivity);
        }

        private Rectangle GetItemRectangle(int itemIndex)
        {
            IntPtr sendResult;
            IntPtr messageResult;
            NM.tagRect rectangle = new NM.tagRect();

            //Get the rectangle of the item
            sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_GETITEMRECT, new IntPtr(itemIndex), ref rectangle, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero || messageResult.ToInt64() == NM.LB_ERR)  //Failed
            {
                throw GUI.ApeException("Failed to get the rectangle of the " + Description);
            }
            return new Rectangle(rectangle.left, rectangle.top, (rectangle.right - rectangle.left), (rectangle.bottom - rectangle.top));
        }

        internal void SingleClickItemInternal(string itemText, CaseSensitivity caseSensitivity)
        {
            //locate the item
            int itemIndex = ItemIndex(itemText, caseSensitivity);
            if (itemIndex == NM.LB_ERR)
            {
                throw GUI.ApeException("Failed to find the item in the " + Description);
            }

            //Locate the rect of the item
            Rectangle itemRectangle = GetItemRectangle(itemIndex);

            //scroll the item into view if needed
            IntPtr sendResult;
            IntPtr messageResult;
            NM.tagRect ClientRect;
            NM.GetClientRect(Identity.Handle, out ClientRect);

            if ((itemRectangle.Height / 2) + itemRectangle.Top > ClientRect.bottom || (itemRectangle.Height / 2) + itemRectangle.Top < ClientRect.top)
            {
                sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_SETTOPINDEX, new IntPtr(itemIndex), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
                if (sendResult == IntPtr.Zero || messageResult.ToInt64() == NM.LB_ERR)  //Failed
                {
                    throw GUI.ApeException("Failed to access the listbox of the " + Description);
                }

                //Wait for animation to finish
                //base.WaitForAnimation(Identity.Handle, false);
                itemRectangle = GetItemRectangle(itemIndex);
            }

            //click the item
            base.SingleClickInternal(-1, (itemRectangle.Height / 2) + itemRectangle.Top, MouseButton.Left, MouseKeyModifier.None);

            //wait for selected == item
            int selectedIndex = -1;
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                sendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_GETCURSEL, IntPtr.Zero, IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out messageResult);
                if (sendResult == IntPtr.Zero || messageResult.ToInt64() == NM.LB_ERR)  //Failed
                {
                    throw GUI.ApeException("Failed to access the listbox of the " + Description);
                }
                selectedIndex = messageResult.ToInt32();

                if (selectedIndex == itemIndex)
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
    }
}
