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
using System.Reflection;
using System.Text;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to improve performance of accessing controls derived from the grid control such as the elementstripgrid
    /// </summary>
    public partial class APEIPC
    {
        private PropertyInfo ColumnHeaderLevelPropertyInfo = null;

        /// <summary>
        /// Initialise the grid control helper delegates for use
        /// </summary>
        private void SetupGridControlHelperDelegates()
        {
            GridControlGetTitleRowsDelegater = new GridControlGetTitleRowsDelegate(GridControlGetTitleRowsInternal);
            GridControlGetTitleRowCountDelegater = new GridControlGetTitleRowCountDelegate(GridControlGetTitleRowCountInternal);
            GridControlGetAllColumnsVisibleDelegater = new GridControlGetAllColumnsVisibleDelegate(GridControlGetAllColumnsVisibleInternal);
        }

        //
        //   GridControlGetAllColumnsVisible
        //

        private delegate string GridControlGetAllColumnsVisibleDelegate(dynamic grid);
        private GridControlGetAllColumnsVisibleDelegate GridControlGetAllColumnsVisibleDelegater;

        /// <summary>
        /// Calls into the AUT to iterate over every column in the grid returning a comma separated string of whether
        /// the column is hidden (True) or visible (False)
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        unsafe public void AddQueryMessageGridControlGetAllColumnsVisible(DataStores sourceStore, DataStores destinationStore)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Need to find the control before querying it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            ptrMessage->Action = MessageAction.GridControlGetAllColumnsHidden;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Calls the GridControlGetAllColumnsHiddenInternal method on the correct thread storing the results
        /// in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void GridControlGetAllColumnsVisible(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject };
                destinationObject = ((WF.Control)tempStore0).Invoke(GridControlGetAllColumnsVisibleDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Iterates over every column in the grid returning a comma separated string of whether the column is
        /// visible (True) or hidden (False)
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <returns>A comma delimited string of whether the column is visible</returns>
        private string GridControlGetAllColumnsVisibleInternal(dynamic grid)
        {
            StringBuilder columnsVisibleState = new StringBuilder(10240);

            // Get the total number of columns
            int columns = grid.Columns.DataColumns.Count;
            
            // Get which column are visible
            for (int column = 0; column < columns; column++)
            {
                bool visible = grid.Columns.DataColumns[column].Visible;

                if (column > 0)
                {
                    columnsVisibleState.Append(",");
                }

                if (visible)
                {
                    columnsVisibleState.Append("True");
                }
                else
                {
                    columnsVisibleState.Append("False");
                }
            }

            return columnsVisibleState.ToString();
        }

        //
        //  GridControlGetTitleRowCount
        //

        private delegate int GridControlGetTitleRowCountDelegate(dynamic grid);
        private GridControlGetTitleRowCountDelegate GridControlGetTitleRowCountDelegater;

        /// <summary>
        /// Calls into the AUT to get the number of title rows in the grid
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant int into</param>
        unsafe public void AddQueryMessageGridControlGetTitleRowCount(DataStores sourceStore, DataStores destinationStore)
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
            ptrMessage->Action = MessageAction.GridControlGetTitleRowCount;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls the GridControlGetTitleRowCountInternal method 
        /// on the correct thread storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void GridControlGetTitleRowCount(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject };
                destinationObject = ((WF.Control)tempStore0).Invoke(GridControlGetTitleRowCountDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Gets the number of title rows in the grid
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <returns>The number of rows in the grid title (header)</returns>
        private int GridControlGetTitleRowCountInternal(dynamic grid)
        {
            // Get the max level
            int rows = -1;
            foreach (dynamic column in grid.Columns)
            {
                if (ColumnHeaderLevelPropertyInfo == null)
                {
                    Type columnType = column.GetType();
                    ColumnHeaderLevelPropertyInfo = columnType.GetProperty("ColumnHeaderLevel", BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                }

                int columnHeaderLevel = ColumnHeaderLevelPropertyInfo.GetValue(column);
                if (columnHeaderLevel > rows)
                {
                    rows = columnHeaderLevel;
                }
            }
            rows++;
            return rows;
        }

        //
        //  GridControlGetTitleRows
        //

        private delegate string[,] GridControlGetTitleRowsDelegate(dynamic grid);
        private GridControlGetTitleRowsDelegate GridControlGetTitleRowsDelegater;

        /// <summary>
        /// Calls into the AUT to get the title rows in the grid
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        unsafe public void AddQueryMessageGridControlGetTitleRows(DataStores sourceStore, DataStores destinationStore)
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
            ptrMessage->Action = MessageAction.GridControlGetTitleRows;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls the GridControlGetTitleRowsInternal method 
        /// on the correct thread storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void GridControlGetTitleRows(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject };
                destinationObject = ((WF.Control)tempStore0).Invoke(GridControlGetTitleRowsDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Gets the title (header) rows in the grid
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <returns>The grid title (header)</returns>
        private string[,] GridControlGetTitleRowsInternal(dynamic grid)
        {
            // Get the total number of columns and rows in the title
            int columns = grid.Columns.DataColumns.Count;
            int rows = GridControlGetTitleRowCountInternal(grid);

            // Create the array
            string[,] titles = new string[rows, columns];

            // Populate the array
            int columnIndex = 0;
            BuildTitleArray(grid.Columns, ref titles, ref columnIndex);

            return titles;
        }

        /// <summary>
        /// Builds a 2d array of the grid title
        /// </summary>
        /// <param name="Columns">The grid internal column collection</param>
        /// <param name="titles">The array to populate</param>
        /// <param name="columnIndex">The current column index in the array</param>
        private void BuildTitleArray(dynamic Columns, ref string[,] titles, ref int columnIndex)
        {
            foreach (dynamic column in Columns)
            {
                int rowIndex = titles.GetLength(0) - ColumnHeaderLevelPropertyInfo.GetValue(column) - 1;
                titles[rowIndex, columnIndex] = column.ToString();
                if ((column.Columns != null) && (column.Columns.Count > 0))
                {
                    BuildTitleArray(column.Columns, ref titles, ref columnIndex);
                }
                else
                {
                    columnIndex++;
                }
            }
        }
    }
}