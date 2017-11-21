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
using WF = System.Windows.Forms;
using System.Runtime.InteropServices;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to help interact with SentinelGrids
    /// </summary>
    public partial class APEIPC
    {
        private delegate object SentinelGridsGetUnderlyingGridDelegate(object resultsGrid);
        private SentinelGridsGetUnderlyingGridDelegate m_SentinelGridsGetUnderlyingGridDelegater;

        /// <summary>
        /// Initialise the sentinel grids helper delegates for use
        /// </summary>
        private void SetupSentinelGridsHelperDelegates()
        {
            m_SentinelGridsGetUnderlyingGridDelegater = new SentinelGridsGetUnderlyingGridDelegate(SentinelGridsGetUnderlyingGridInternal);
        }

        /// <summary>
        /// Casts the provided object as an ILZGrid and returns the underlying grid
        /// </summary>
        /// <param name="resultsGrid">The object to cast to an ILZGrid</param>
        /// <returns>The underlying grid</returns>
        private object SentinelGridsGetUnderlyingGridInternal(object resultsGrid)
        {
            //runtime late binding for COM uses the dispatch interface but we don't want the default (ILZResultsGrid)
            //interface we want the ILZGrid interface which means we need to use early binding and cast the object.
            ILZGrid resultsGridAsLzGrid = (ILZGrid)resultsGrid;
            object theGrid = resultsGridAsLzGrid.UnderlyingGrid;
            return theGrid;
        }

        /// <summary>
        /// Calls into the AUT to cast the specified object as an ILZGrid
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the object to cast to an ILZGrid</param>
        /// <param name="destinationStore">The datastore to put the underlying grid into</param>
        unsafe public void AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores sourceStore, DataStores destinationStore)
        {
            if (m_DoneFind == false)
            {
                throw new Exception("Need to find the control before querying it");
            }

            if (m_DoneGet == true)
            {
                throw new Exception("Can not query control after getting values from it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;
            ptrMessage->Action = MessageAction.SentinelGridsGetUnderlyingGrid;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls the SentinelGridsGetUnderlyingGridInternal method 
        /// on the correct thread storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void SentinelGridsGetUnderlyingGrid(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_SentinelGridsGetUnderlyingGridDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// The ILZGrid interface
        /// </summary>
        [ComImport]
        [Guid("FB3D6B62-B8A5-41EC-98E8-ED01D2E023E5")]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [TypeLibType(TypeLibTypeFlags.FNonExtensible)]
        private interface ILZGrid
        {
            object UnderlyingGrid
            {
                [return: MarshalAs(UnmanagedType.IUnknown)]get;
            }
        }
    }
}
