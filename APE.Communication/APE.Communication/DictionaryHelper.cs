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
using System.Reflection;
using System.Text;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to improve performance of accessing flexgrid controls
    /// </summary>
    public partial class APEIPC
    {
        /// <summary>
        /// Initialise the flexgrid helper delegates for use
        /// </summary>
        private void SetupDictionaryHelperDelegates()
        {
            m_DictionaryContainsKeyDelegater = new DictionaryContainsKeyDelegate(DictionaryContainsKeyInternal);
        }

        //
        //  DictionaryContainsKey
        //

        private delegate bool DictionaryContainsKeyDelegate(object dictionary, string key);
        private DictionaryContainsKeyDelegate m_DictionaryContainsKeyDelegater;

        /// <summary>
        /// Checks if the dictionary contains the specified key
        /// </summary>
        /// <param name="dictionary">The grid object</param>
        /// <param name="key">The start row of the range</param>
        /// <returns>True if the dictionary contains the key</returns>
        private bool DictionaryContainsKeyInternal(object dictionary, string key)
        {
            //dynamic doesn't work here so we have to use reflection, so this is rather expensive
            MethodInfo methodInfoContainsKey = dictionary.GetType().GetMethod("ContainsKey");
            object[] parametersArray = new object[] { key };
            bool result = (bool)methodInfoContainsKey.Invoke(dictionary, parametersArray);
            return result;
        }

        /// <summary>
        /// Calls into the AUT to check if the dictionary contains the specified key
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="key">The key to check for</param>
        unsafe public void AddQueryMessageDictionaryContainsKey(DataStores sourceStore, DataStores destinationStore, string key)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the dictionary before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;
            ptrMessage->Action = MessageAction.DictionaryContainsKey;

            Parameter keyParameter = new Parameter(this, key);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls the DictionaryContainsKeyInternal method 
        /// on the correct thread storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void DictionaryContainsKey(Message* ptrMessage)
        {
            string key = GetParameterString(ptrMessage, 0);
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, key };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_DictionaryContainsKeyDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }
    }
}
