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
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;
using NM = APE.Native.NativeMethods;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    public partial class APEIPC
    {

        /// <summary>
        /// Initialise the data grid view helper delegates for use
        /// </summary>
        private void SetupDataGridViewHelperDelegates()
        {
            DataGridViewShowCellDelegater = new DataGridViewShowCellDelegate(DataGridViewShowCellInternal);
        }

        //
        //  DataGridViewShowCell
        //

        private delegate void DataGridViewShowCellDelegate(dynamic grid, int rowIndexData, int columnIndexData);
        private DataGridViewShowCellDelegate DataGridViewShowCellDelegater;

        /// <summary>
        /// Calls into the AUT to add a selected handler to the specific control
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the control object</param>
        unsafe public void AddQueryMessageDataGridViewShowCell(DataStores sourceStore, int rowIndexData, int columnIndexData)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the data grid view before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->Action = MessageAction.DataGridViewShowCell;

            Parameter rowIndexDataParameter = new Parameter(this, rowIndexData);
            Parameter columnIndexDataParameter = new Parameter(this, columnIndexData);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Shows a data column / row in the data grid view
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void DataGridViewShowCell(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            int rowIndexData = GetParameterInt32(ptrMessage, 0);
            int columnIndexData = GetParameterInt32(ptrMessage, 1);

            object[] theParameters = { sourceObject, rowIndexData, columnIndexData };
            ((WF.Control)tempStore0).Invoke(DataGridViewShowCellDelegater, theParameters);

            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Shows a data column / row in the data grid view
        /// </summary>
        /// <param name="grid">The grid</param>
        /// <param name="rowIndexData">The data row index</param>
        /// <param name="columnIndexData">The data column index</param>
        unsafe private void DataGridViewShowCellInternal(dynamic grid, int rowIndexData, int columnIndexData)
        {
            grid.FirstDisplayedScrollingRowIndex = rowIndexData;
            grid.FirstDisplayedScrollingColumnIndex = columnIndexData;
        }
    }
}
