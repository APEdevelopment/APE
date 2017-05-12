//
//Copyright 2016-2017 David Beales
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
using System.Reflection;
using APE.Communication;
using System.Drawing;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Common.ModelControls.ElementStrip.Grouping.GroupingBar
    /// </summary>
    public class GUIGroupingBar : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIGroupingBar(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Clicks on the specified item using the specified button
        /// </summary>
        /// <param name="itemIndex">The index of the item</param>
        /// <param name="button">The button with which to click</param>
        public virtual void SingleClickItem(int itemIndex, MouseButton button)
        {
            SingleClickItemInternal(itemIndex, itemIndex.ToString(), button);
        }

        /// <summary>
        /// Clicks on the specified item using the specified button
        /// </summary>
        /// <param name="itemText">The text of the item to click on</param>
        /// <param name="button">The button with which to click</param>
        public virtual void SingleClickItem(string itemText, MouseButton button)
        {
            int itemIndex = ItemIndex(itemText);
            SingleClickItemInternal(itemIndex, itemText, button);
        }

        private void SingleClickItemInternal(int itemIndex, string ItemName, MouseButton button)
        {
            Rectangle tabRectangle = ItemRectangle(itemIndex);
            NM.tagRect tabRect;
            NM.GetClientRect(Identity.Handle, out tabRect);
            GUI.Log("Single " + button.ToString() + " click on the " + Identity.Description + " item " + ItemName, LogItemType.Action);
            int x = tabRectangle.X + (tabRectangle.Width / 2);
            int y = tabRectangle.Y + (tabRectangle.Height / 2);
            base.SingleClickInternal(x, y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Returns the number of items in the group bar
        /// </summary>
        /// <returns>The number of items</returns>
        public int ItemCount()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GroupingSpecification", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int itemCount = GUI.m_APE.GetValueFromMessage();
            return itemCount;
        }

        /// <summary>
        /// Returns the index of the item with the specified text
        /// </summary>
        /// <param name="itemText">The text of the item</param>
        /// <returns>The index of the item</returns>
        public int ItemIndex(string itemText)
        {
            int itemCount = ItemCount();

            for (int item = 0; item < itemCount; item++)
            {
                if (ItemText(item) == itemText)
                {
                    return item;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the text of the item at the specified index
        /// </summary>
        /// <param name="itemIndex">The index of the item</param>
        /// <returns>The text of the item</returns>
        public string ItemText(int itemIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GroupingSpecification", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, itemIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "FieldInstance", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Title", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string itemText = GUI.m_APE.GetValueFromMessage();
            return itemText;
        }

        private Rectangle ItemRectangle(int itemIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetGroupPositionRectangle", MemberTypes.Method, new Parameter(GUI.m_APE, itemIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "X", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Y", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store4, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Height", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int x = GUI.m_APE.GetValueFromMessage();
            int y = GUI.m_APE.GetValueFromMessage();
            int width = GUI.m_APE.GetValueFromMessage();
            int height = GUI.m_APE.GetValueFromMessage();

            Rectangle itemRectangle = new Rectangle(x, y, width, height);
            return itemRectangle;
        }
    }
}
