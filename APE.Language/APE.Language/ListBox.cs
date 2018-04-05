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
            int Index = ItemIndex(itemText);

            if (Index == NM.LB_ERR)
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int itemCount = GUI.m_APE.GetValueFromMessage();
            return itemCount;
        }

        internal int ItemIndex(string itemText)
        {
            //Get the index
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "FindStringExact", MemberTypes.Method, new Parameter(GUI.m_APE, itemText));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Index = GUI.m_APE.GetValueFromMessage();

            return Index;
        }

        /// <summary>
        /// Returns the text of the item at the specified index
        /// </summary>
        /// <param name="itemIndex">The index of the item</param>
        /// <returns>The text of the item</returns>
        public string ItemText(int itemIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, itemIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ToString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string itemText = GUI.m_APE.GetValueFromMessage();
            return itemText;
        }

        /// <summary>
        /// Selects the specified item in the listbox by clicking on it
        /// </summary>
        /// <param name="itemText">The item to select</param>
        public void SingleClickItem(string itemText)
        {
            GUI.Log("Single Left click on the item " + itemText + " from the " + Identity.Description, LogItemType.Action);
            SingleClickItemInternal(itemText);
        }

        internal void SingleClickItemInternal(string itemText)
        {
            //locate the item
            int Index = ItemIndex(itemText);
            if (Index == NM.LB_ERR)
            {
                throw new Exception("Failed to find the item in the " + Description);
            }

            //Locate the rect of the item
            int Left;
            int Top;
            int Right;
            int Bottom;
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetItemRectangle", MemberTypes.Method, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Left", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Top", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store4, "Right", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Bottom", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            Left = GUI.m_APE.GetValueFromMessage();
            Top = GUI.m_APE.GetValueFromMessage();
            Right = GUI.m_APE.GetValueFromMessage();
            Bottom = GUI.m_APE.GetValueFromMessage();

            //scroll the item into view if needed
            NM.tagRect ClientRect;
            NM.GetClientRect(Identity.Handle, out ClientRect);

            if (((Bottom - Top) / 2) + Top > ClientRect.bottom || ((Bottom - Top) / 2) + Top < ClientRect.top)
            {
                // Currently APE doesn't support setting parameters or fields and I don't really want to add support as we should be doing things like a real user would
                //GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                //GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "TopIndex", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
                //GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store2, "TopIndex", MemberTypes.);
                //GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                //GUI.m_APE.SendMessages(EventSet.APE);
                //GUI.m_APE.WaitForMessages(EventSet.APE);
                ////get the values returned
                //int CurrentIndex = GUI.m_APE.GetValueFromMessage();
                IntPtr SendResult;
                IntPtr MessageResult;

                SendResult = NM.SendMessageTimeout(Identity.Handle, NM.ListBoxMessages.LB_SETTOPINDEX, new IntPtr(Index), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                if (SendResult == IntPtr.Zero || MessageResult.ToInt64() == NM.LB_ERR)  //Failed
                {
                    throw new Exception("Failed to access the listbox of the " + Description);
                }
                //Wait for animation to finish
                //base.WaitForAnimation(Identity.Handle, false);

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetItemRectangle", MemberTypes.Method, new Parameter(GUI.m_APE, Index));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Top", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store4, "Right", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //get the values returned
                Left = GUI.m_APE.GetValueFromMessage();
                Top = GUI.m_APE.GetValueFromMessage();
                Right = GUI.m_APE.GetValueFromMessage();
                Bottom = GUI.m_APE.GetValueFromMessage();
            }

            //click the item
            base.SingleClickInternal(-1, ((Bottom - Top) / 2) + Top, MouseButton.Left, MouseKeyModifier.None);

            //wait for selected == item
            int SelectedIndex;
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedIndex", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                SelectedIndex = GUI.m_APE.GetValueFromMessage();

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to select the item in the " + Description);
                }

                Thread.Sleep(15);
            }
            while (SelectedIndex != Index);
            timer.Stop();
        }
    }
}
