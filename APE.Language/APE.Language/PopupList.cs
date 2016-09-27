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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using NM = APE.Native.NativeMethods;
using APE.Communication;
using System.Reflection;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Utility.Controls.PopupControlHelper
    /// </summary>
    public sealed class GUIPopupList : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIPopupList(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the text portion of the TextValueWithTypeControl to the specified text by sending keystrokes
        /// </summary>
        /// <param name="item">The text to set the text portion of the TextValueWithTypeControl to</param>
        public void Select(string item)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "PopupControlHelper", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "InternalList", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            int count = GUI.m_APE.GetValueFromMessage();

            for (int itemIndex = 0; itemIndex < count; itemIndex++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "PopupControlHelper", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "InternalList", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, itemIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Text", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Top", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store8, "Right", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store9, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call;
                string textOfitem = GUI.m_APE.GetValueFromMessage();
                int Left = GUI.m_APE.GetValueFromMessage();
                int Top = GUI.m_APE.GetValueFromMessage();
                int Right = GUI.m_APE.GetValueFromMessage();
                int Bottom = GUI.m_APE.GetValueFromMessage();
                
                if (textOfitem == item)
                {
                    // we only do a mouse down as the control disappears before any up event can be sent
                    GUI.Log("Select item " + item + " in the " + Identity.Description, LogItemType.Action);
                    base.MouseDownInternal(5, Top + ((Bottom - Top) / 2), MouseButton.Left, MouseKeyModifier.None);
                    return;
                }
            }

            throw new Exception("Failed to find item " + item + " in popup list");
        }
    }
}
