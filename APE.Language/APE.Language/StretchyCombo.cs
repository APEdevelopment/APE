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

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LzStretchyCombo.StretchyCombo
    /// </summary>
    public sealed class GUIStretchyCombo : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIStretchyCombo(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
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
            int itemIndex = ItemIndex(item);
            if (itemIndex == -1)
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

            // Check if its already selected
            if (SelectedItemText() == item)
            {
                GUI.Log("Ensure " + m_DescriptionOfControl + " is set to " + item, LogItemTypeEnum.Action);
                return;
            }

            GUI.Log("Select [" + item + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            int itemIndex = ItemIndex(item);

            if (itemIndex == -1)
            {
                throw new Exception("Item " + item + " does not exist in the " + Description);
            }

            if (!IsDropped())
            {
                base.MouseSingleClickInternal(Width - 5, 5, MouseButton.Left, MouseKeyModifier.None);
            }

            GUIForm dropDownForm = new GUIForm("drop down form", new Identifier(Identifiers.Name, "frmList"));

            //TODO scroll the item into view if need be

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, IntPtr.Zero, dropDownForm.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_sngItemHeight", MemberTypes.Field);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int itemHeight = (int)GUI.m_APE.GetValueFromMessage();

            int x = dropDownForm.Width / 2;
            int y = (itemHeight * itemIndex) + (itemHeight / 2);

            dropDownForm.MouseSingleClickInternal(x, y, MouseButton.Left, MouseKeyModifier.None);

            timer = Stopwatch.StartNew();
            while (true)
            {
                if (SelectedItemText() == item)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > 5000)
                {
                    throw new Exception("Failed to select item " + item + " in the " + Description);
                }

                Thread.Sleep(50);
            }
        }

        private bool IsDropped()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_bPopupDown", MemberTypes.Field);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isDropped = GUI.m_APE.GetValueFromMessage();
            return isDropped;
        }

        private string SelectedItemText()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_sSelectedCaption", MemberTypes.Field);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string selectedCaption = GUI.m_APE.GetValueFromMessage();
            return selectedCaption;
        }

        private int ItemIndex(string item)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_colItems", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int itemCount = GUI.m_APE.GetValueFromMessage();
            
            for (int itemIndex = 0; itemIndex < itemCount; itemIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_colItems", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, itemIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Caption", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemText = GUI.m_APE.GetValueFromMessage();

                if (itemText == item)
                {
                    return itemIndex;
                }
            }

            return -1;
        }
    }
}
