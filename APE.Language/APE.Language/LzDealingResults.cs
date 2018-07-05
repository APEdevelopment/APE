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
using System;
using System.Reflection;
using APE.Communication;

namespace APE.Language
{

    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LzcDealingResults.lzDealingResults
    /// </summary>
    public sealed class GUILzDealingResults : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUILzDealingResults(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Gets whether all results have been returned
        /// </summary>
        /// <returns>True if all results have been returned or there are no results otherwise false</returns>
        public bool AllResultsBack()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "AllResultsBack", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool allResultsBack = GUI.m_APE.GetValueFromMessage();
            return allResultsBack;
        }

        /// <summary>
        /// Gets the results grid related to the lzDealingResults object
        /// </summary>
        /// <returns>The results grid</returns>
        public GUIAxLZResultsGrid GetResultsGrid()
        {
            GUIAxLZResultsGrid resultsGrid = new GUIAxLZResultsGrid(ParentForm, Description + " results grid", new Identifier(Identifiers.Name, "lzcResultsGrid"), new Identifier(Identifiers.ChildOf, this));
            return resultsGrid;
        }
    }
}
