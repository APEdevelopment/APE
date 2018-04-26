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
using APE.Communication;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.StatusBar
    /// </summary>
    public sealed class GUIStatusBar : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIStatusBar(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Returns the number of panels
        /// </summary>
        /// <returns>The number of panels</returns>
        public int PanelCount()
        {
            //Get the number of items
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            int items = GUI.m_APE.GetValueFromMessage();
            return items;
        }

        /// <summary>
        /// Returns the index of the specified panel name
        /// </summary>
        /// <param name="PanelName">The name of the panel to get the index of</param>
        /// <returns>The index of the panel</returns>
        public int PanelIndex(String PanelName)
        {
            int items = PanelCount();

            //Loop through looking for the item we want
            for (int Item = 0; Item < items; Item++)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, Item));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //get the values returned
                string CurrentPanelName = GUI.m_APE.GetValueFromMessage();

                if (CurrentPanelName == PanelName)
                {
                    //found it
                    return Item;
                }
            }

            //Failed to find it
            throw GUI.ApeException("Failed to find panel with name [" + PanelName + "]");
        }

        /// <summary>
        /// Returns the name of the specified panel index
        /// </summary>
        /// <param name="PanelIndex">The index of the panel to get the name of</param>
        /// <returns>The name of the panel</returns>
        public string PanelName(int PanelIndex)
        {
            //query the panel text 0 based index
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, PanelIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            return GUI.m_APE.GetValueFromMessage();
        }

        /// <summary>
        /// Polls for the panel with the specified name to have the specified text
        /// </summary>
        /// <param name="panelName">The name of the panel to poll</param>
        /// <param name="text">The text to wait for the panel to have</param>
        public void PanelPollForText(string panelName, string text)
        {
            int panelIndex = PanelIndex(panelName);
            PanelPollForText(panelIndex, text);
        }

        /// <summary>
        /// Polls for the panel at the specified index to have the specified text
        /// </summary>
        /// <param name="panelIndex">The index of the panel to poll</param>
        /// <param name="text">The text to wait for the panel to have</param>
        public void PanelPollForText(int panelIndex, string text)
        {
            //query the panel text 0 based index
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, panelIndex));
            GUI.m_APE.AddMessagePollMember(DataStores.Store2, "Text", MemberTypes.Property, new Parameter(GUI.m_APE, text), new Parameter(GUI.m_APE, true));
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
        }

        /// <summary>
        /// Returns the text of the panel with the specified name
        /// </summary>
        /// <param name="panelName">The name of the panel to get the text of</param>
        /// <returns>The text of the panel</returns>
        public string PanelText(string panelName)
        {
            int panelIndex = PanelIndex(panelName);
            return PanelText(panelIndex);
        }

        /// <summary>
        /// Returns the text of the panel at the specified index
        /// </summary>
        /// <param name="panelIndex">The index of the panel to get the text of</param>
        /// <returns>The text of the panel</returns>
        public string PanelText(int panelIndex)
        {
            //query the panel text 0 based index
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, panelIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);

            GUI.m_APE.SendMessages(EventSet.APE);

            GUI.m_APE.WaitForMessages(EventSet.APE);

            //get the values returned
            return GUI.m_APE.GetValueFromMessage();
        }
    }
}
