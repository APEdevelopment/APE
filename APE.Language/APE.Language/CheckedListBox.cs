using System.Runtime.InteropServices;
using APE.Capture;
using APE.Communication;
using System.Threading;
using System.Drawing.Imaging;
using System.Security.Principal;
using NM = APE.Native.NativeMethods;
using System;
using System.Reflection;
using System.Diagnostics;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.CheckedListBox
    /// </summary>
    public sealed class GUICheckedListBox : GUIFocusableObject
    {
        private GUIListBox ListBox;

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUICheckedListBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
            ListBox = new GUIListBox(parentForm, descriptionOfControl, new Identifier(Identifiers.Handle, Handle));
        }

        /// <summary>
        /// Checks if the specified item exists in the checkedlistbox
        /// </summary>
        /// <param name="itemText">The item to check if it exists</param>
        /// <returns>Whether item exists</returns>
        public bool ItemExists(string itemText)
        {
            return ListBox.ItemExists(itemText);
        }

        /// <summary>
        /// Returns the number of items in the checkedlistbox
        /// </summary>
        /// <returns>The number of items</returns>
        public int ItemCount()
        {
            return ListBox.ItemCount();
        }

        internal int ItemIndex(string itemText)
        {
            return ListBox.ItemIndex(itemText);
        }

        /// <summary>
        /// Returns the text of the item at the specified index
        /// </summary>
        /// <param name="itemIndex">The index of the item</param>
        /// <returns>The text of the item</returns>
        public string ItemText(int itemIndex)
        {
            return ListBox.ItemText(itemIndex);
        }

        /// <summary>
        /// Selects the specified item in the checkedlistbox by clicking on it
        /// </summary>
        /// <param name="itemText">The item to select</param>
        public void SingleClickItem(string itemText)
        {
            ListBox.SingleClickItem(itemText);
        }

        /// <summary>
        /// Get whether the item in the checkedlistbox is checked or not
        /// </summary>
        /// <param name="itemText">The text of the item to determine if checked</param>
        /// <returns>True if the item is checked otherwise false</returns>
        public bool ItemChecked(string itemText)
        {
            int itemIndex = ItemIndex(itemText);
            return (ItemChecked(itemIndex));
        }

        /// <summary>
        /// Get whether the item in the checkedlistbox is checked or not
        /// </summary>
        /// <param name="itemIndex">The index of the item to determine if checked</param>
        /// <returns>True if the item is checked otherwise false</returns>
        public bool ItemChecked(int itemIndex)
        {
            //Get if the item is checked
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetItemChecked", MemberTypes.Method, new Parameter(GUI.m_APE, itemIndex));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            return GUI.m_APE.GetValueFromMessage();
        }

        /// <summary>
        /// Checks the specified item in the checkedlistbox
        /// </summary>
        /// <param name="itemText">The item to check</param>
        public void ItemCheck(string itemText)
        {
            if (ItemChecked(itemText) == true)
            {
                GUI.Log("Ensure " + itemText + " in the " + Description + " is checked", LogItemType.Action);
            }
            else
            {
                GUI.Log("Check " + itemText + " in the " + Description, LogItemType.Action);
                ListBox.SingleClickItemInternal(itemText);
                PollForState(itemText, true);
            }
        }

        /// <summary>
        /// Unchecks the specified item in the checkedlistbox
        /// </summary>
        /// <param name="itemText">The item to uncheck</param>
        public void ItemUncheck(string itemText)
        {
            if (ItemChecked(itemText) == false)
            {
                GUI.Log("Ensure " + itemText + " in the " + Description + " is unchecked", LogItemType.Action);
            }
            else
            {
                GUI.Log("Uncheck " + itemText + " in the " + Description, LogItemType.Action);
                ListBox.SingleClickItemInternal(itemText);
                PollForState(itemText, false);
            }
        }

        /// <summary>
        /// Send the specified text to the control
        /// </summary>
        /// <param name="text">The text to send to the control</param>
        public void Type(string text)
        {
            SetFocus();
            SendKeys(text);
        }

        private void PollForState(string itemText, bool state)
        {
            int itemIndex = ItemIndex(itemText);
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (ItemChecked(itemIndex) == state)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    string stateText;
                    if (state)
                    {
                        stateText = "check";
                    }
                    else
                    {
                        stateText = "uncheck";
                    }
                    throw GUI.ApeException("Failed to " + stateText + " item " + itemText + " of the " + Description);
                }

                Thread.Sleep(15);
            }
        }
    }
}