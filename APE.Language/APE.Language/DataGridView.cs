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
using System.Drawing;
using System.Reflection;
using APE.Communication;
using System.Threading;
using System.Diagnostics;
using NM = APE.Native.NativeMethods;
using System.Text;
using System.ComponentModel;
using System.Collections.Generic;
using System.Windows.Forms.VisualStyles;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.DataGridView
    /// </summary>
    public sealed class GUIDataGridView : GUIGridObject
    {
        /// <summary>
        /// Initialises a new instance of the GUIDataGridView class
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIDataGridView(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Returns true if the specified row is hidden in the grid
        /// </summary>
        /// <param name="rowText">Row to check if hidden</param>
        /// <returns>True or False</returns>
        public bool IsRowHidden(string rowText)
        {
            int rowIndex = FindRow(rowText);
            return IsRowHidden(rowIndex);
        }

        /// <summary>
        /// Returns true if the specified row is hidden in the grid
        /// </summary>
        /// <param name="rowIndex">Row index to check if hidden</param>
        /// <returns>True or False</returns>
        public bool IsRowHidden(int rowIndex)
        {
            if (rowIndex < TitleRows())
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ColumnHeadersVisible", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            }
            else
            {
                int rowIndexData = rowIndex - TitleRows();

                // TODO include where height = 0 or 1?
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            }
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool visible = GUI.m_APE.GetValueFromMessage();
            return !visible;
        }

        /// <summary>
        /// Returns true if the specified column is hidden in the grid
        /// </summary>
        /// <param name="columnText">Column to check if hidden delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns>True or False</returns>
        public bool IsColumnHidden(string columnText)
        {
            int columnIndex = FindColumn(columnText);
            return IsColumnHidden(columnIndex);
        }

        /// <summary>
        /// Returns true if the specified column is hidden in the grid
        /// </summary>
        /// <param name="columnIndex">Column index to check if hidden</param>
        public bool IsColumnHidden(int columnIndex)
        {
            if (columnIndex < TitleColumns())
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "RowHeadersVisible", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            }
            else
            {
                int columnIndexData = columnIndex - TitleColumns();

                // TODO include where width = 0 or 1?
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            }
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool visible = GUI.m_APE.GetValueFromMessage();
            return !visible;
        }

        /// <summary>
        /// Determines if the specified cell is currently viewable without scrolling
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to check</param>
        /// <param name="columnIndex">The column index of the cell to check</param>
        /// <returns>True if the cell is visible without scrolling otherwise false</returns>
        public bool IsCellVisible(int rowIndex, int columnIndex)
        {
            bool rowDisplayed;
            bool columnDisplayed;

            //Check row
            if (rowIndex < TitleRows())
            {
                rowDisplayed = true;   //title rows always shown?
            }
            else
            {
                int rowIndexData = rowIndex - TitleRows();
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Displayed", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                rowDisplayed = GUI.m_APE.GetValueFromMessage();
            }

            //Check column
            if (columnIndex < TitleRows())
            {
                columnDisplayed = true;   //title rows always shown?
            }
            else
            {
                int columnIndexData = columnIndex - TitleColumns();
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Displayed", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                columnDisplayed = GUI.m_APE.GetValueFromMessage();
            }
            
            if (rowDisplayed && columnDisplayed)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Returns the first visible (non-hidden) row in the grid
        /// </summary>
        /// <returns>The first visible row</returns>
        public int FirstVisibleRow()
        {
            int rowCount = Rows();
            for (int rowIndex = 0; rowIndex < rowCount; rowIndex++)
            {
                if (!IsRowHidden(rowIndex))
                {
                    return rowIndex;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the first visible (non-hidden) column in the grid
        /// </summary>
        /// <returns>The first visible column</returns>
        public int FirstVisibleColumn()
        {
            int columnCount = Columns();
            for (int columnIndex = 0; columnIndex < columnCount; columnIndex++)
            {
                if (!IsColumnHidden(columnIndex))
                {
                    return columnIndex;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the selected row
        /// </summary>
        /// <returns>The selected row index</returns>
        public int SelectedRow()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CurrentCell", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "RowIndex", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int selectedRowIndex = GUI.m_APE.GetValueFromMessage() + TitleRows();
            return selectedRowIndex;
        }

        /// <summary>
        /// Returns the selected column
        /// </summary>
        /// <returns>The selected column index</returns>
        public int SelectedColumn()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CurrentCell", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ColumnIndex", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int selectedColumnIndex = GUI.m_APE.GetValueFromMessage() + TitleColumns();
            return selectedColumnIndex;
        }

        /// <summary>
        /// Returns the number of fixed rows, that is a row which makes up the header of the grid (rows may or may not be hidden)
        /// </summary>
        /// <returns>The number of fixed rows</returns>
        public int FixedRows()
        {
            return TitleRows();
        }

        /// <summary>
        /// Returns the number of frozen rows, that is a row which doesn't scroll, in the grid (rows may or may not be hidden)
        /// </summary>
        /// <returns>The number of frozen rows</returns>
        public int FrozenRows()
        {
            int dataRows = Rows() - TitleRows();
            for (int rowIndexData = dataRows - 1; rowIndexData >= 0; rowIndexData--)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Frozen", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                bool frozen = GUI.m_APE.GetValueFromMessage();
                
                if (frozen)
                {
                    return rowIndexData + TitleRows();
                }
            }
            return TitleRows();
        }

        /// <summary>
        /// Returns the number of fixed columns, that is a column which doesn't scroll, in the grid (columns may or may not be hidden)
        /// </summary>
        /// <returns>The number of fixed columns</returns>
        public int FixedColumns()
        {
            return TitleColumns();
        }

        /// <summary>
        /// Returns the number of frozen columns, that is a column which doesn't scroll, in the grid (columns may or may not be hidden)
        /// </summary>
        /// <returns>The number of frozen columns</returns>
        public int FrozenColumns()
        {
            int dataColumns = Columns() - TitleColumns();
            for (int columnIndexData = dataColumns - 1; columnIndexData >= 0; columnIndexData--)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Frozen", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                bool frozen = GUI.m_APE.GetValueFromMessage();

                if (frozen)
                {
                    return columnIndexData + TitleColumns();
                }
            }
            return TitleColumns();
        }

        /// <summary>
        /// Returns the number of title (column header) rows in the grid (the rows may or may not be visible)
        /// </summary>
        /// <returns>The number of title rows</returns>
        public int TitleRows()
        {
            return 1;
        }

        /// <summary>
        /// Returns the number of title (row header) columns in the grid (the columns may or may not be visible)
        /// </summary>
        /// <returns>The number of title columns</returns>
        public int TitleColumns()
        {
            return 1;
        }

        /// <summary>
        /// Returns the number of rows in the grid, including those which are hidden
        /// </summary>
        /// <returns>The number of rows</returns>
        public int Rows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int rows = GUI.m_APE.GetValueFromMessage() + TitleRows();
            return rows;
        }

        /// <summary>
        /// Returns the number of columns, including those which are hidden
        /// </summary>
        /// <returns>The number of columns</returns>
        public int Columns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int columns = GUI.m_APE.GetValueFromMessage() + TitleColumns();
            return columns;
        }

        /// <summary>
        /// Returns whether at the grid level it is editable
        /// </summary>
        /// <returns>True if it is editable otherwise false</returns>
        public bool IsEditable()
        {
            if (IsEnabled)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ReadOnly", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                bool readOnly = GUI.m_APE.GetValueFromMessage();
                return !readOnly;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Returns whether the specified cell is editable
        /// </summary>
        /// <param name="rowText">The row of the cell to check if its editable</param>
        /// <param name="columnText">The column of the cell to check if its editable</param>
        /// <returns>True if the cell is editable otherwise false</returns>
        public bool IsCellEditable(string rowText, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);
            return IsCellEditable(rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns whether the specified cell is editable
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to check if its editable</param>
        /// <param name="columnText">The column of the cell to check if its editable</param>
        /// <returns>True if the cell is editable otherwise false</returns>
        public bool IsCellEditable(int rowIndex, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            return IsCellEditable(rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns whether the specified cell is editable
        /// </summary>
        /// <param name="rowText">The row of the cell to check if its editable</param>
        /// <param name="columnIndex">The column index of the cell to check if its editable</param>
        /// <returns>True if the cell is editable otherwise false</returns>
        public bool IsCellEditable(string rowText, int columnIndex)
        {
            int rowIndex = FindRow(rowText, columnIndex);
            return IsCellEditable(rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns whether the specified cell is editable
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to check if its editable</param>
        /// <param name="columnIndex">The column index of the cell to check if its editable</param>
        /// <returns>True if the cell is editable otherwise false</returns>
        public bool IsCellEditable(int rowIndex, int columnIndex)
        {
            if (rowIndex < TitleRows())
            {
                return false;
            }

            if (columnIndex < TitleColumns())
            {
                return false;
            }

            int rowIndexData = rowIndex - TitleRows();
            int columnIndexData = columnIndex - TitleColumns();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cells", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ReadOnly", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool readOnly = GUI.m_APE.GetValueFromMessage();
            return !readOnly;
        }

        /// <summary>
        /// Polls for the specified row index to be the selected row
        /// </summary>
        /// <param name="rowIndex">The row index to wait to be selected</param>
        public void SelectedRowPollForIndex(int rowIndex)
        {
            int rowIndexData = rowIndex - TitleRows();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CurrentCell", MemberTypes.Property);
            GUI.m_APE.AddMessagePollMember(DataStores.Store1, "RowIndex", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData), new Parameter(GUI.m_APE, false));
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
        }

        private string ColumnType(int columnIndex)
        {
            int columnIndexData = columnIndex - TitleColumns();
            string ColumnType = null;

            if (columnIndexData >= 0)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "CellType", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Name", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                ColumnType = GUI.m_APE.GetValueFromMessage();
            }

            if (ColumnType == null)
            {
                throw GUI.ApeException("Unable to determine column type for column " + columnIndex.ToString() + " of the " + Description);
            }
            return ColumnType;
        }

        internal override bool SetCellValueInternal<T>(string rowText, string columnText, int rowIndex, int columnIndex, T value, T expectedValue, string submitKey, ComparisonMethod compareMethod)
        {
            string rowFriendlyText;
            string columnFriendlyText;
            Stopwatch timer;

            if (rowText == null)
            {
                rowFriendlyText = rowIndex.ToString();
            }
            else
            {
                rowFriendlyText = rowText;
            }

            if (columnText == null)
            {
                columnFriendlyText = columnIndex.ToString();
            }
            else
            {
                columnFriendlyText = columnText;
            }

            if (submitKey == null)
            {
                submitKey = "{Enter}";
            }

            // Couple of sanity checks
            if (rowIndex < 0)
            {
                throw GUI.ApeException("Must supply a row index greater than 0 in the " + Description);
            }

            if (columnIndex < 0)
            {
                throw GUI.ApeException("Must supply a column index greater than 0 in the " + Description);
            }

            string currentValueAsString = this.GetCell(rowIndex, columnIndex);
            T currentValue = (T)Convert.ChangeType(currentValueAsString, typeof(T));

            // Check if the cell is already set to the correct value
            if (compareMethod == ComparisonMethod.DoNotCompare)
            {
                string expectedValueAsString = expectedValue.ToString();
                decimal currentValueAsDecimal;
                decimal expectedValueAsDecimal;

                if (decimal.TryParse(currentValueAsString, out currentValueAsDecimal) && decimal.TryParse(expectedValueAsString, out expectedValueAsDecimal))
                {
                    //numeric value in the string so comvert to a decimal and compare
                    if (currentValueAsDecimal == expectedValueAsDecimal)
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                        return false;
                    }
                }
                else
                {
                    //non numeric so just do a check if its equal
                    if (currentValueAsString == expectedValueAsString)
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                        return false;
                    }
                }
            }
            else
            {
                if (EqualityComparer<T>.Default.Equals(currentValue, expectedValue))
                {
                    GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                    return false;
                }
            }

            if (IsEnabled)
            {
                // Set focus to the grid, we can't use SetFocus() here as we want the grid to get focus regardless
                // of whether a child of the grid has it or not
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Focus", MemberTypes.Method);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
            }

            // Get the type of column
            string columnType = ColumnType(columnIndex);
            string valueText = value.ToString();

            switch (columnType)
            {
                case "DataGridViewTextBoxCell":
                    // Select the cell if its not selected
                    if (this.SelectedRow() == rowIndex && this.SelectedColumn() == columnIndex)
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is selected", LogItemType.Action);
                    }
                    else
                    {
                        GUI.Log("Single " + MouseButton.Left.ToString() + " click on the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                        this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                        SelectedRowPollForIndex(rowIndex);
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemType.Action);
                    base.SendKeysInternal("{F2}");

                    // wait for the DataGridViewTextBoxEditingControl
                    GUITextBox textBox = new GUITextBox(ParentForm, Description + " texbox", new Identifier(Identifiers.ChildOf, this), new Identifier(Identifiers.TypeName, "DataGridViewTextBoxEditingControl"));
                    textBox.SetText(valueText);
                    break;
                case "DataGridViewComboBoxCell":
                    // Select the cell if its not selected
                    if (this.SelectedRow() == rowIndex && this.SelectedColumn() == columnIndex)
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is selected", LogItemType.Action);
                    }
                    else
                    {
                        GUI.Log("Single " + MouseButton.Left.ToString() + " click on the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                        this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                        SelectedRowPollForIndex(rowIndex);
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemType.Action);
                    base.SendKeysInternal("{F2}");

                    // wait for the DataGridViewComboBoxEditingControl
                    GUIComboBox comboBox = new GUIComboBox(ParentForm, Description + " combobox", new Identifier(Identifiers.ChildOf, this), new Identifier(Identifiers.TypeName, "DataGridViewComboBoxEditingControl"));
                    comboBox.SingleClickItem(valueText);
                    break;
                case "DataGridViewCheckBoxCell":
                    // Click on the checkbox
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                    // May need to determine location to click at some point
                    this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                //case "DataGridViewLinkCell":
                //case "DataGridViewButtonCell":
                //case "DataGridViewImageCell":
                default:
                    throw GUI.ApeException("Unsupported data grid view column type: " + columnType.ToString());
            }

            GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
            base.SendKeysInternal(submitKey);

            if (IsEnabled && ParentForm.IsEnabled)
            {
                // We have changed a value so make sure it has painted
                GUI.m_APE.AddFirstMessagePeakMessage(Identity.Handle);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);

                Input.WaitForInputIdle(this.Handle, (uint)GUI.GetTimeOut());

                // Set focus to the grid, we can't use SetFocus() here as we want the grid to get focus regardless
                // of whether a child of the grid has it or not
                GUI.m_APE.AddFirstMessageSetFocusAsync(Identity.Handle);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
            }

            switch (compareMethod)
            {
                case ComparisonMethod.CompareUsingDefaultEqualityComparer:
                    // Check the value was set
                    timer = Stopwatch.StartNew();
                    int loop = 0;
                    int sleep = 15;
                    while (true)
                    {
                        currentValueAsString = this.GetCell(rowIndex, columnIndex, CellProperty.TextDisplay);
                        currentValue = (T)Convert.ChangeType(currentValueAsString, typeof(T));
                        if (EqualityComparer<T>.Default.Equals(currentValue, expectedValue))
                        {
                            break;
                        }

                        // If the form it belongs to isn't enabled then there is likely a modal form displayed
                        // so exit the loop so the code can continue (its up to the caller to validate the
                        // value is set in these cases)
                        if (!ParentForm.IsEnabled)
                        {
                            break;
                        }

                        if (columnText != null)
                        {
                            // Look for the column again as the act of setting the value may have changed its position
                            columnIndex = FindColumn(columnText);
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw GUI.ApeException("Failed to set the " + Description + " cell value");
                        }

                        // Increase the sleep time at certain loop thresholds
                        switch (loop)
                        {
                            case 100:
                                sleep = 50;
                                break;
                            case 200:
                                sleep = 250;
                                break;
                        }

                        loop++;
                        Thread.Sleep(sleep);
                    }
                    break;
                case ComparisonMethod.DoNotCompare:
                    break;
                default:
                    throw GUI.ApeException("Unsupported compare method: " + compareMethod.ToString());
            }

            return true;
        }

        private bool HeadersVisualStylesEnabled()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "EnableHeadersVisualStyles", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool enableHeadersVisualStyles = GUI.m_APE.GetValueFromMessage();
            return enableHeadersVisualStyles;
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        internal override dynamic GetCellInternal(int rowIndex, int columnIndex, CellProperty property)
        {
            //If visual styles are supported and EnableHeadersVisualStyles is set to true then the header back and fore colours come
            //from the visual style instead of what the grid is set to

            bool headersVisualStylesEnabled;
            bool visualStyleSupported;
            VisualStyleRenderer visualStyleRenderer;

            if (rowIndex < TitleRows() && columnIndex < TitleColumns())
            {
                switch (property)
                {
                    case CellProperty.TextDisplay:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TopLeftHeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "FormattedValue", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string formattedValue = GUI.m_APE.GetValueFromMessage();
                        return formattedValue;
                    case CellProperty.BackColourName:
                        headersVisualStylesEnabled = HeadersVisualStylesEnabled();
                        visualStyleSupported = GUI.AttachedProcessVisualStyleSupported();

                        if (headersVisualStylesEnabled && visualStyleSupported)
                        {
                            visualStyleRenderer = new VisualStyleRenderer(VisualStyleElement.Header.Item.Normal);
                            Bitmap background = new Bitmap(1, 1);
                            Graphics backgroundGraphics = Graphics.FromImage(background);
                            Rectangle bounds = Rectangle.Truncate(backgroundGraphics.ClipBounds);
                            visualStyleRenderer.DrawBackground(backgroundGraphics, bounds, bounds);
                            Color backColour = background.GetPixel(0, 0);
                            return backColour.Name;
                        }
                        else
                        {
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TopLeftHeaderCell", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "InheritedStyle", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "BackColor", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            string backColourName = GUI.m_APE.GetValueFromMessage();
                            return backColourName;
                        }
                    case CellProperty.ForeColourName:
                        headersVisualStylesEnabled = HeadersVisualStylesEnabled();
                        visualStyleSupported = GUI.AttachedProcessVisualStyleSupported();

                        if (headersVisualStylesEnabled && visualStyleSupported)
                        {
                            visualStyleRenderer = new VisualStyleRenderer(VisualStyleElement.Header.Item.Normal);
                            Color foreColour = visualStyleRenderer.GetColor(ColorProperty.TextColor);
                            return foreColour.Name;
                        }
                        else
                        {
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TopLeftHeaderCell", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "InheritedStyle", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ForeColor", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            string foreColourName = GUI.m_APE.GetValueFromMessage();
                            return foreColourName;
                        }
                    case CellProperty.DataType:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TopLeftHeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "FormattedValueType", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Namespace", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "Name", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string cellTypeNamespace = GUI.m_APE.GetValueFromMessage();
                        string cellTypeName = GUI.m_APE.GetValueFromMessage();
                        return cellTypeNamespace + "." + cellTypeName;
                    case CellProperty.FontStyle:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TopLeftHeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "InheritedStyle", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Font", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Style", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string fontStyleText = GUI.m_APE.GetValueFromMessage();
                        FontStyle fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), fontStyleText);
                        return fontStyle;
                    default:
                        throw GUI.ApeException("Implement support for getting cell property " + property.ToString());
                }
            }
            else if (rowIndex < TitleRows())
            {
                int columnIndexData = columnIndex - TitleColumns();

                switch (property)
                {
                    case CellProperty.TextDisplay:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "FormattedValue", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string formattedValue = GUI.m_APE.GetValueFromMessage();
                        return formattedValue;
                    case CellProperty.BackColourName:
                        headersVisualStylesEnabled = HeadersVisualStylesEnabled();
                        visualStyleSupported = GUI.AttachedProcessVisualStyleSupported();

                        if (headersVisualStylesEnabled && visualStyleSupported)
                        {
                            visualStyleRenderer = new VisualStyleRenderer(VisualStyleElement.Header.Item.Normal);
                            Bitmap background = new Bitmap(1, 1);
                            Graphics backgroundGraphics = Graphics.FromImage(background);
                            Rectangle bounds = Rectangle.Truncate(backgroundGraphics.ClipBounds);
                            visualStyleRenderer.DrawBackground(backgroundGraphics, bounds, bounds);
                            Color backColour = background.GetPixel(0, 0);
                            return backColour.Name;
                        }
                        else
                        {
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "InheritedStyle", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "BackColor", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            string backColourName = GUI.m_APE.GetValueFromMessage();
                            return backColourName;
                        }
                    case CellProperty.ForeColourName:
                        headersVisualStylesEnabled = HeadersVisualStylesEnabled();
                        visualStyleSupported = GUI.AttachedProcessVisualStyleSupported();

                        if (headersVisualStylesEnabled && visualStyleSupported)
                        {
                            visualStyleRenderer = new VisualStyleRenderer(VisualStyleElement.Header.Item.Normal);
                            Color foreColour = visualStyleRenderer.GetColor(ColorProperty.TextColor);
                            return foreColour.Name;
                        }
                        else
                        {
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "InheritedStyle", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ForeColor", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            string foreColourName = GUI.m_APE.GetValueFromMessage();
                            return foreColourName;
                        }
                    case CellProperty.DataType:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "FormattedValueType", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Namespace", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store6, "Name", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string cellTypeNamespace = GUI.m_APE.GetValueFromMessage();
                        string cellTypeName = GUI.m_APE.GetValueFromMessage();
                        return cellTypeNamespace + "." + cellTypeName;
                    case CellProperty.FontStyle:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "InheritedStyle", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Font", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Style", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string fontStyleText = GUI.m_APE.GetValueFromMessage();
                        FontStyle fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), fontStyleText);
                        return fontStyle;
                    default:
                        throw GUI.ApeException("Implement support for getting cell property " + property.ToString());
                }
            }
            else if (columnIndex < TitleColumns())
            {
                int rowIndexData = rowIndex - TitleRows();

                switch (property)
                {
                    case CellProperty.TextDisplay:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "FormattedValue", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string formattedValue = GUI.m_APE.GetValueFromMessage();
                        return formattedValue;
                    case CellProperty.BackColourName:
                        headersVisualStylesEnabled = HeadersVisualStylesEnabled();
                        visualStyleSupported = GUI.AttachedProcessVisualStyleSupported();

                        if (headersVisualStylesEnabled && visualStyleSupported)
                        {
                            visualStyleRenderer = new VisualStyleRenderer(VisualStyleElement.Header.Item.Normal);
                            Bitmap background = new Bitmap(1, 1);
                            Graphics backgroundGraphics = Graphics.FromImage(background);
                            Rectangle bounds = Rectangle.Truncate(backgroundGraphics.ClipBounds);
                            visualStyleRenderer.DrawBackground(backgroundGraphics, bounds, bounds);
                            Color backColour = background.GetPixel(0, 0);
                            return backColour.Name;
                        }
                        else
                        {
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "InheritedStyle", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "BackColor", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            string backColourName = GUI.m_APE.GetValueFromMessage();
                            return backColourName;
                        }
                    case CellProperty.ForeColourName:
                        headersVisualStylesEnabled = HeadersVisualStylesEnabled();
                        visualStyleSupported = GUI.AttachedProcessVisualStyleSupported();

                        if (headersVisualStylesEnabled && visualStyleSupported)
                        {
                            visualStyleRenderer = new VisualStyleRenderer(VisualStyleElement.Header.Item.Normal);
                            Color foreColour = visualStyleRenderer.GetColor(ColorProperty.TextColor);
                            return foreColour.Name;
                        }
                        else
                        {
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "InheritedStyle", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ForeColor", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            //Get the value(s) returned MUST be done straight after the WaitForMessages call
                            string foreColourName = GUI.m_APE.GetValueFromMessage();
                            return foreColourName;
                        }
                    case CellProperty.DataType:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "FormattedValueType", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Namespace", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store6, "Name", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string cellTypeNamespace = GUI.m_APE.GetValueFromMessage();
                        string cellTypeName = GUI.m_APE.GetValueFromMessage();
                        return cellTypeNamespace + "." + cellTypeName;
                    case CellProperty.FontStyle:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "HeaderCell", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "InheritedStyle", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Font", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Style", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string fontStyleText = GUI.m_APE.GetValueFromMessage();
                        FontStyle fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), fontStyleText);
                        return fontStyle;
                    default:
                        throw GUI.ApeException("Implement support for getting cell property " + property.ToString());
                }
            }
            else
            {
                int rowIndexData = rowIndex - TitleRows();
                int columnIndexData = columnIndex - TitleColumns();

                switch (property)
                {
                    case CellProperty.TextDisplay:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cells", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "FormattedValue", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string formattedValue = GUI.m_APE.GetValueFromMessage();
                        return formattedValue;
                    case CellProperty.BackColourName:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cells", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "InheritedStyle", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "BackColor", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "Name", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string backColourName = GUI.m_APE.GetValueFromMessage();
                        return backColourName;
                    case CellProperty.ForeColourName:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cells", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "InheritedStyle", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "ForeColor", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "Name", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string foreColourName = GUI.m_APE.GetValueFromMessage();
                        return foreColourName;
                    case CellProperty.DataType:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cells", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "FormattedValueType", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Namespace", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Name", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string cellTypeNamespace = GUI.m_APE.GetValueFromMessage();
                        string cellTypeName = GUI.m_APE.GetValueFromMessage();
                        return cellTypeNamespace + "." + cellTypeName;
                    case CellProperty.FontStyle:
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cells", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndexData));
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "InheritedStyle", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Font", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "Style", MemberTypes.Property);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "ToString", MemberTypes.Method);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string fontStyleText = GUI.m_APE.GetValueFromMessage();
                        FontStyle fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), fontStyleText);
                        return fontStyle;
                    default:
                        throw GUI.ApeException("Implement support for getting cell property " + property.ToString());
                }
            }
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Index">The start column of the range</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Index">The end column of the range</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>A string containing the range of values</returns>
        internal override string GetCellRangeInternal(int row1Index, int column1Index, int row2Index, int column2Index, CellProperty property)
        {
            //TODO make this more efficient by doing it in process
            StringBuilder grid = new StringBuilder();
            for (int row = row1Index; row <= row2Index; row++)
            {
                for (int column = column1Index; column <= column2Index; column++)
                {
                    grid.Append(GetCell(row, column, property).ToString());
                    if (column == column2Index)
                    {
                        grid.Append("\r");
                    }
                    else
                    {
                        grid.Append("\t");
                    }
                }
            }
            if (grid.Length > 0)
            {
                grid.Length--;
            }
            return grid.ToString();
        }

        /// <summary>
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        internal override void ShowCellInternal(int rowIndex, int columnIndex)
        {
            //if (!IsCellVisible(rowIndex, columnIndex))    //Commented out as partial visible cells cause issues
            //{
                if (IsRowHidden(rowIndex))
                {
                    throw GUI.ApeException("Row " + rowIndex.ToString() + " in the " + Description + " is hidden");
                }

                if (IsColumnHidden(columnIndex))
                {
                    throw GUI.ApeException("Column " + columnIndex.ToString() + " in the " + Description + " is hidden");
                }

                int rowIndexData = rowIndex - TitleRows();
                int columnIndexData = columnIndex - TitleColumns();

                if (rowIndexData < 0)
                {
                    rowIndexData = 0;
                }

                if (columnIndexData < 0)
                {
                    columnIndexData = 0;
                }

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageDataGridViewShowCell(DataStores.Store0, rowIndexData, columnIndexData);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
            //}
        }

        internal override void MoveToCellInternal(int X, int Y)
        {
            MoveTo(X, Y);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        internal override void SingleClickCellInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            SingleClickInternal(Location.X, Location.Y, button, keyModifier);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        internal override void DoubleClickCellInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            DoubleClickInternal(Location.X, Location.Y, button, keyModifier);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        internal override void MouseDownInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            MouseDownInternal(Location.X, Location.Y, button, keyModifier);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        internal override void MouseUpInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            MouseUpInternal(Location.X, Location.Y, button, keyModifier);
        }

        /// <summary>
        /// Gets the location given a row and column and position in a cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <param name="locationInCell">The location in the cell</param>
        /// <returns>The points location</returns>
        internal override Point GetLocationInCell(int rowIndex, int columnIndex, CellClickLocation locationInCell)
        {
            if (rowIndex < 0)
            {
                throw GUI.ApeException("Must supply a row index greater than 0 in the " + Description);
            }

            if (columnIndex < 0)
            {
                throw GUI.ApeException("Must supply a column index greater than 0 in the " + Description);
            }

            //Check to make sure the row isn't hidden
            if (IsRowHidden(rowIndex))
            {
                throw GUI.ApeException("Row is hidden in the " + Description);
            }

            //Check to make sure the column isn't hidden
            if (IsColumnHidden(columnIndex))
            {
                throw GUI.ApeException("Column is hidden in the " + Description);
            }

            //Scroll the cell into view
            ShowCell(rowIndex, columnIndex);

            Rectangle CellRectangle = GetCellRectangle(rowIndex, columnIndex);

            Point Location = new Point();

            //Adjust for where we want to click in the cell
            switch (locationInCell)
            {
                case CellClickLocation.LeftSideOfCell:
                    Location.X = CellRectangle.Left + 5;
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                case CellClickLocation.CentreOfCell:
                    Location.X = CellRectangle.Left + (CellRectangle.Width / 2);
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                case CellClickLocation.RightSideOfCell:
                    Location.X = CellRectangle.Left + CellRectangle.Width - 5;
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                default:
                    throw GUI.ApeException("Implement for CellClickLocation: " + Location.ToString());
            }

            return Location;
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <returns>The bound rectangle</returns>
        internal override Rectangle GetCellRectangleInternal(int rowIndex, int columnIndex)
        {
            int x = 0;
            int y = 0;
            int width = 0;
            int height = 0;

            if (rowIndex < TitleRows() && columnIndex < TitleColumns())
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "RowHeadersWidth", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "ColumnHeadersHeight", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                width = GUI.m_APE.GetValueFromMessage();
                height = GUI.m_APE.GetValueFromMessage();
            }
            else if (rowIndex < TitleRows())
            {
                int columnIndexData = columnIndex - TitleColumns();

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetColumnDisplayRectangle", MemberTypes.Method, new Parameter(GUI.m_APE, columnIndexData), new Parameter(GUI.m_APE, true));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "X", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Y", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store4, "Width", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store5, "ColumnHeadersHeight", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                x = GUI.m_APE.GetValueFromMessage();
                y = GUI.m_APE.GetValueFromMessage();
                width = GUI.m_APE.GetValueFromMessage();
                height = GUI.m_APE.GetValueFromMessage();
            }
            else if (columnIndex < TitleColumns())
            {
                int rowIndexData = rowIndex - TitleRows();

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRowDisplayRectangle", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndexData), new Parameter(GUI.m_APE, true));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "X", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Y", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store4, "RowHeadersWidth", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Height", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                x = GUI.m_APE.GetValueFromMessage();
                y = GUI.m_APE.GetValueFromMessage();
                width = GUI.m_APE.GetValueFromMessage();
                height = GUI.m_APE.GetValueFromMessage();
            }
            else
            {
                int rowIndexData = rowIndex - TitleRows();
                int columnIndexData = columnIndex - TitleColumns();

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellDisplayRectangle", MemberTypes.Method, new Parameter(GUI.m_APE, columnIndexData), new Parameter(GUI.m_APE, rowIndexData), new Parameter(GUI.m_APE, true));
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
                x = GUI.m_APE.GetValueFromMessage();
                y = GUI.m_APE.GetValueFromMessage();
                width = GUI.m_APE.GetValueFromMessage();
                height = GUI.m_APE.GetValueFromMessage();
            }

            Rectangle CellRectangle = new Rectangle(x, y, width, height);
            return CellRectangle;
        }

        // TODO Need to move this to GridObject at some point
        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnIndex">The column to look for the value in</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row or -1</returns>
        internal override int FindRowTemp(string rowText, int columnIndex, int startAtRow)
        {
            int rowIndex;
            if (columnIndex == -1)
            {
                columnIndex = FirstVisibleColumn();
                rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
            }
            else
            {
                rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
            }
            return rowIndex;
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnIndex">The column to look for the value in</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row or -1</returns>
        internal override int FindRowInternal(string rowText, int columnIndex, int startAtRow)
        {
            int rows = Rows();
            for (int row = 0; row < rows; row++)
            {
                if (rowText == GetCell(row, columnIndex, CellProperty.TextDisplay))
                {
                    return row;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns the index of the specified column in the grid
        /// </summary>
        /// <param name="columnHeader">The column to check</param>
        /// <returns>The index of the column</returns>
        internal override int FindColumnInternal(string[] columnHeader)
        {
            string columnText = columnHeader[0];
            if (columnHeader.GetLength(0) > 1)
            {
                throw GUI.ApeException("Implement support for more than one 1 title row");
            }
            int columns = Columns();
            for (int column = 0; column < columns; column++)
            {
                if (columnText == GetCell(0, column, CellProperty.TextDisplay))
                {
                    return column;
                }
            }
            return -1;
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r where
        /// each cell has a width and height greater than 1 pixel and the row or column is not
        /// hidden.  Collapsed nodes are also excluded.
        /// </summary>
        /// <returns>A string containing the range of values</returns>
        public string GetAllVisibleCells()
        {
            return GetAllVisibleCells(CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r where
        /// each cell has a width and height greater than 1 pixel and the row or column is not
        /// hidden.  Collapsed nodes are also excluded.
        /// </summary>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>A string containing the range of values</returns>
        public string GetAllVisibleCells(CellProperty property)
        {
            //TODO make this more efficient by doing it in process
            int row1Index = 0;
            int row2Index = Rows() - 1;
            int column1Index = 0;
            int column2Index = Columns() - 1;
            StringBuilder grid = new StringBuilder();

            for (int row = row1Index; row <= row2Index; row++)
            {
                if (!IsRowHidden(row))
                {
                    for (int column = column1Index; column <= column2Index; column++)
                    {
                        if (!IsColumnHidden(column))
                        {
                            grid.Append(GetCell(row, column, property).ToString());
                            if (column == column2Index)
                            {
                                grid.Append("\r");
                            }
                            else
                            {
                                grid.Append("\t");
                            }
                        }
                    }
                }
            }
            if (grid.Length > 0)
            {
                grid.Length--;
            }
            return grid.ToString();
        }
    }
}