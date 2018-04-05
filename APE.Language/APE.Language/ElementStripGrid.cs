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
using System.Text;
using System.Reflection;
using APE.Communication;
using System.Threading;
using System.Diagnostics;
using System.Collections.Generic;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    // TODO this should probably extend GridControl or EditorStrip with most of the methods in there?

    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Common.ModelControls.ElementStrip.ElementStripGrid
    /// </summary>
    public sealed class GUIElementStripGrid : GUIGridObject
    {
        /// <summary>
        /// Initialises a new instance of the GUIElementStripGrid class
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIElementStripGrid(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Returns the number of title (column header) rows in the grid
        /// </summary>
        /// <returns>The number of title rows</returns>
        public int TitleRows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageGridControlGetTitleRowCount(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int titleRows = GUI.m_APE.GetValueFromMessage();

            return titleRows;
        }

        /// <summary>
        /// Returns the number of rows in the grid
        /// </summary>
        /// <returns>The number of rows</returns>
        public int Rows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NumRows", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Rows = GUI.m_APE.GetValueFromMessage();

            return Rows + TitleRows();
        }

        /// <summary>
        /// Returns the number of fixed / frozen rows, that is a row which doesn't scroll, in the grid (rows may or may not be hidden)
        /// </summary>
        /// <returns>The number of fixed / frozen rows</returns>
        public int FixedRows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NumFrozenRows", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int frozenRows = GUI.m_APE.GetValueFromMessage();
            return frozenRows + TitleRows();
        }

        /// <summary>
        /// Returns the number of fixed / frozen columns, that is a row which doesn't scroll, in the grid (columns may or may not be hidden)
        /// </summary>
        /// <returns>The number of fixed / frozen columns</returns>
        public int FixedColumns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "DataColumns", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "LastFrozenColumnIndex", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int frozenColumns = GUI.m_APE.GetValueFromMessage();
            return frozenColumns;
        }

        /// <summary>
        /// Returns the number of selected rows
        /// </summary>
        /// <returns>The number of selected rows</returns>
        public int SelectedRows()
        {
            int result;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetSelectedRowIndexesAsString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string SelectedRowIndexesAsString = GUI.m_APE.GetValueFromMessage();

            if (SelectedRowIndexesAsString == "")
            {
                result = 0;
            }
            else
            {
                char[] separator = { ',' };
                string[] SelectedRowIndexesAsStringArray = SelectedRowIndexesAsString.Split(separator);
                result = SelectedRowIndexesAsStringArray.GetLength(0);
            }

            return result;
        }

        /// <summary>
        /// Returns whether the sepecified rows is in the current selection
        /// </summary>
        /// <param name="row">The row to check</param>
        /// <returns>True or False</returns>
        public bool SelectedRowsContains(int row)
        {
            row -= TitleRows();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedRows", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Contains", MemberTypes.Method, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool result = GUI.m_APE.GetValueFromMessage();

            return result;
        }

        /// <summary>
        /// Returns the column the of the cursor cell
        /// </summary>
        /// <returns>The cursor column number</returns>
        public int CursorCellColumn()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedCell", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Col", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int column = GUI.m_APE.GetValueFromMessage();
            return column;
        }

        /// <summary>
        /// Returns the row the of the cursor cell
        /// Note this cell may not currently be a selected row.  Use SelectedRow if you need to know if it is selected 
        /// </summary>
        /// <returns>The cursor row number</returns>
        public int CursorCellRow()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedCell", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Row", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int row = GUI.m_APE.GetValueFromMessage();
            row += TitleRows();
            return row;
        }

        /// <summary>
        /// Returns the selected row
        /// </summary>
        /// <returns>The selected row number</returns>
        public int SelectedRow()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "SelectedRow", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int row = GUI.m_APE.GetValueFromMessage();
            row += TitleRows();
            return row;
        }

        /// <summary>
        /// Polls for the specified row index to be the selected row
        /// </summary>
        /// <param name="rowIndex">The row index to wait to be selected</param>
        public void SelectedRowPollForIndex(int rowIndex)
        {
            int row = rowIndex - TitleRows();

            if (row < 0)
            {
                throw new Exception("Title rows are not currently supported");
            }

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessagePollMember(DataStores.Store0, "SelectedRow", MemberTypes.Property, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, false));
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
        }

        /// <summary>
        /// Returns an array of indexes of the currently selected rows
        /// </summary>
        /// <returns>An array of indexes</returns>
        public int[] GetSelectedRows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetSelectedRowIndexesAsString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string selectedRowIndexesString = GUI.m_APE.GetValueFromMessage();

            int titleRows = TitleRows();
            int[] selectedRowIndexesintArray;

            if (selectedRowIndexesString == "")
            {
                selectedRowIndexesintArray = new int[0];
            }
            else
            {
                char[] splitSeparator = { ',' };
                string[] selectedRowIndexesStringArray = selectedRowIndexesString.Split(splitSeparator);
                selectedRowIndexesintArray = Array.ConvertAll(selectedRowIndexesStringArray, s => int.Parse(s) + titleRows);
            }
            
            return selectedRowIndexesintArray;
        }

        /// <summary>
        /// Returns the number of columns, including those which are hidden
        /// </summary>
        /// <returns>The number of columns</returns>
        public int Columns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NumDataColumns", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int columns = GUI.m_APE.GetValueFromMessage();
            return columns;
        }

        private string[,] GetColumnTitles()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageGridControlGetTitleRows(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            string titlesText = GUI.m_APE.GetValueFromMessage();

            string[,] titlesArray = ConvertStringToString2dArray(titlesText);

            return titlesArray;
        }

        /// <summary>
        /// Returns the index of the specified column in the grid
        /// </summary>
        /// <param name="columnHeader">The column to check</param>
        /// <returns>The index of the column</returns>
        internal override int FindColumnInternal(string[] columnHeader)
        {
            string[,] titles = GetColumnTitles();
            bool[] visibleColumns = GetColumnsVisibilityState();

            //Search for the column
            bool Found = false;
            int columns = titles.GetLength(1);
            int row = 0;
            int column = -1;
            for (column = 0; column < columns; column++)
            {
                if (visibleColumns[column])
                {
                    int startHeaderRow = columnHeader.GetLength(0) - 1;

                    if (titles[startHeaderRow, column] == columnHeader[startHeaderRow] || titles.GetLength(0) > columnHeader.GetLength(0)) 
                    {
                        if (titles.GetLength(0) > columnHeader.GetLength(0))
                        {
                            startHeaderRow++;
                        }
                        Found = true;
                        for (row = startHeaderRow - 1; row > -1; row--)
                        {
                            int tempColumn = column;

                            if (columnHeader[row] != "")
                            {
                                while (titles[row, tempColumn] == "")
                                {
                                    if (tempColumn == 0)
                                    {
                                        break;
                                    }
                                    else
                                    {
                                        tempColumn--;
                                    }
                                }
                            }

                            if (titles[row, tempColumn] == columnHeader[row])
                            {
                                Found = true;
                            }
                            else
                            {
                                Found = false;
                                break;
                            }
                        }

                        if (Found)
                        {
                            break;
                        }
                    }
                }
            }

            if (Found)
            {
                return column;
            }

            return -1;
        }

        /// <summary>
        /// Gets the tree view column
        /// </summary>
        /// <returns>The tree view column index or -1</returns>
        public int TreeViewColumn()
        {
            int treeViewColumn = 0;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "DrawTree", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            ////Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool drawTree = GUI.m_APE.GetValueFromMessage();

            if (!drawTree)
            {
                treeViewColumn = -1;
            }
            return treeViewColumn;
        }

        /// <summary>
        /// Determines if the grid has a tree view column
        /// </summary>
        /// <returns>True if the grid has a tree view column otherwise false</returns>
        public bool IsTreeView()
        {
            int TreeColumn = TreeViewColumn();

            if (TreeColumn == -1)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private bool[] GetColumnsVisibilityState()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageGridControlGetAllColumnsVisible(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string columnsVisibleText = GUI.m_APE.GetValueFromMessage();

            bool[] columnsVisibleArray;

            if (columnsVisibleText == "")
            {
                columnsVisibleArray = new bool[0];
            }
            else
            {
                char[] splitSeparator = { ',' };
                string[] columnsVisibleTextArray = columnsVisibleText.Split(splitSeparator, StringSplitOptions.None);
                columnsVisibleArray = Array.ConvertAll(columnsVisibleTextArray, item => bool.Parse(item));
            }
            return columnsVisibleArray;
        }

        /// <summary>
        /// Returns whether at the grid level it is editable
        /// </summary>
        /// <returns>True if it is editable otherwise false</returns>
        public bool IsEditable()
        {
            if (IsEnabled)
            {
                return true;    // Doesn't have a concept of at the grid level being editable or not
            }
            return false;
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
            int titleRows = TitleRows();
            rowIndex -= titleRows;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "LatentZero.Capstone.Controls.GridControl.CellPosition", MemberTypes.Constructor, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetEditDescriptor", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store3));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "IsReadOnly", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store3));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            bool isReadOnly = GUI.m_APE.GetValueFromMessage();
            return !isReadOnly;
        }

        /// <summary>
        /// Returns the first visible (non-hidden) row in the grid
        /// </summary>
        /// <returns>The first visible row</returns>
        public int FirstVisibleRow()
        {
            //TODO check with simon if the grid has a concept of hidden rows, if it does then this needs fixing
            return 0;
        }

        /// <summary>
        /// Returns the first visible (non-hidden) column in the grid
        /// </summary>
        /// <returns>The first visible column</returns>
        public int FirstVisibleColumn()
        {
            bool[] visibleColumns = GetColumnsVisibilityState();

            for (int i = 0; i < visibleColumns.GetLength(0); i++)
            {
                if (visibleColumns[i] == true)
                {
                    return i;
                }
            }

            return -1;
        }

        /// <summary>
        /// Returns true if the specified column is hidden in the grid
        /// </summary>
        /// <param name="columnText">Column to check if hidden delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns>True or False</returns>
        public bool IsColumnHidden(string columnText)
        {
            int column = FindColumn(columnText);
            return IsColumnHidden(column);
        }

        /// <summary>
        /// Returns true if the specified column is hidden in the grid
        /// </summary>
        /// <param name="columnIndex">Column index to check if hidden</param>
        /// <returns>True or False</returns>
        public bool IsColumnHidden(int columnIndex)
        {
            bool[] visibleColumns = GetColumnsVisibilityState();

            if (visibleColumns[columnIndex])
            {
                return false;
            }
            else
            {
                return true;
            }
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
            //TODO do we want to do this?
            //GetRow(row - titleRows).NodeInformation.AreParentsExpanded
            return false;
        }

        /// <summary>
        /// Returns the index of the top row currently visible in the scrollable area
        /// </summary>
        /// <returns>Index of top visible row</returns>
        public int TopVisibleRow()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "TopVisibleRow", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int row = GUI.m_APE.GetValueFromMessage();
            row = row + TitleRows();
            return row;
        }

        /// <summary>
        /// Returns the index of the bottom row currently visible in the scrollable area
        /// </summary>
        /// <returns>Index of bottom visible row</returns>
        public int BottomVisibleRow()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "BottomVisibleRow", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int row = GUI.m_APE.GetValueFromMessage();
            row = row + TitleRows();
            return row;
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
            switch (property)
            {
                case CellProperty.TextDisplay:
                    string text = "";

                    //Get how many column title rows we have
                    int titleRows = TitleRows();

                    //If the range we want includes the title rows
                    if (row1Index < titleRows)
                    {
                        //Get the title rows
                        string[,] columnTitle = GetColumnTitles();

                        for (int row = row1Index; row < titleRows; row++)
                        {
                            if (row > row2Index)
                            {
                                break;
                            }

                            for (int column = column1Index; column <= column2Index; column++)
                            {
                                if (column1Index == column2Index)
                                {
                                    text += columnTitle[row, column] + "\r";
                                }
                                else
                                {
                                    if (column == column1Index)
                                    {
                                        text += columnTitle[row, column];
                                    }
                                    else if (column == column2Index)
                                    {
                                        text += "\t" + columnTitle[row, column] + "\r";
                                    }
                                    else
                                    {
                                        text += "\t" + columnTitle[row, column];
                                    }
                                }
                            }
                        }
                    }

                    row1Index -= titleRows;
                    row2Index -= titleRows;

                    if (row1Index < 0)
                    {
                        row1Index = 0;
                    }

                    //Get the data rows
                    if (row2Index > -1)
                    {
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellsDisplayTextString", MemberTypes.Method, new Parameter(GUI.m_APE, row1Index), new Parameter(GUI.m_APE, column1Index), new Parameter(GUI.m_APE, row2Index), new Parameter(GUI.m_APE, column2Index));
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        string dataText = GUI.m_APE.GetValueFromMessage();

                        text += dataText;
                    }

                    if (text.EndsWith("\r"))
                    {
                        text = text.Remove(text.Length - 1);
                    }
                    return text;
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }
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
            int cellRowIndex;
            switch (property)
            {
                case CellProperty.TextDisplay:
                    return GetCellRangeInternal(rowIndex, columnIndex, rowIndex, columnIndex, property);
                case CellProperty.BackColourName:
                case CellProperty.ForeColourName:
                case CellProperty.FontStyle:
                    // Cache is only for displayed rows
                    ShowCell(rowIndex, columnIndex);

                    if (rowIndex < TitleRows())
                    {
                        // Title row
                        cellRowIndex = -1;
                        // TODO Styles.ColHeadingStyle or ColumnHeadingStyle for v18+
                    }
                    else if (rowIndex < FixedRows())
                    {
                        // Fixed row
                        cellRowIndex = rowIndex - TitleRows();
                    }
                    else
                    {
                        // Other rows
                        cellRowIndex = rowIndex - TitleRows() - TopVisibleRow() + FixedRows();
                    }

                    switch (property)
                    {
                        case CellProperty.BackColourName:
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "View", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "RowContentsCache", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, cellRowIndex), new Parameter(GUI.m_APE, columnIndex));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "MergedCellStyle", MemberTypes.Field);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "BackColour", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            string backColourName = GUI.m_APE.GetValueFromMessage();
                            return backColourName;
                        case CellProperty.ForeColourName:
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "View", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "RowContentsCache", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, cellRowIndex), new Parameter(GUI.m_APE, columnIndex));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "MergedCellStyle", MemberTypes.Field);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ForeColour", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            string foreColourName = GUI.m_APE.GetValueFromMessage();
                            return foreColourName;
                        case CellProperty.FontStyle:
                            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "View", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "RowContentsCache", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, cellRowIndex), new Parameter(GUI.m_APE, columnIndex));
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "MergedCellStyle", MemberTypes.Field);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Font", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Style", MemberTypes.Property);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "ToString", MemberTypes.Method);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store8, "RowStyleCharacteristic", MemberTypes.Field);
                            GUI.m_APE.AddQueryMessageReflect(DataStores.Store8, DataStores.Store9, "ToString", MemberTypes.Method);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                            GUI.m_APE.SendMessages(EventSet.APE);
                            GUI.m_APE.WaitForMessages(EventSet.APE);
                            string fontStyleText = GUI.m_APE.GetValueFromMessage();
                            string rowStyleCharacteristic = GUI.m_APE.GetValueFromMessage();
                            FontStyle fontStyle;
                            switch (rowStyleCharacteristic)
                            {
                                case "StrikeThrough":
                                    fontStyle = FontStyle.Strikeout;
                                    break;
                                case "Italic": 
                                    fontStyle = FontStyle.Italic;
                                    break;
                                case "None":
                                    fontStyle = (FontStyle)Enum.Parse(typeof(FontStyle), fontStyleText);
                                    break;
                                default:
                                    throw new Exception("Implement support for RowStyleCharacteristic " + rowStyleCharacteristic);
                            }
                            return fontStyle;
                        default:
                            throw new Exception("Implement support for getting cell property " + property.ToString());
                    }
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }
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
            int treeColumn = TreeViewColumn();
            if (columnIndex == -1)
            {
                if (IsTreeView())
                {
                    rowIndex = FindNodeRow(rowText);
                }
                else
                {
                    columnIndex = FirstVisibleColumn();
                    rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
                }
            }
            else
            {
                if (columnIndex == treeColumn)
                {
                    rowIndex = FindNodeRow(rowText);
                }
                else
                {
                    rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
                }
            }
            return rowIndex;
        }

        private int FindNodeRow(string NodeText)
        {
            //TODO this is bit inefficent, it would be better to find from parent to child rather than child to parent
            string[] SplitNodeText = NodeText.Split(GUI.GridDelimiterAsArray, StringSplitOptions.None);

            int column = TreeViewColumn();
            int currentRow = FixedRows() - 1;
            string childNodeText = SplitNodeText[SplitNodeText.GetUpperBound(0)];
            string currentNodeText;

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                //Find a row which matches the child node we are after
                currentRow = FindRowInternal(childNodeText, column, currentRow + 1);

                if (currentRow == -1)
                {
                    break;
                }

                //Check if its parents match the node we are after
                currentNodeText = GetNodePath(currentRow, column);

                if (currentNodeText == NodeText)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to find the row of the node in the " + Description);
                }
            }

            return currentRow;
        }

        private string GetNodePath(int Row, int Column)
        {
            string NodePath = GetCell(Row, Column, CellProperty.TextDisplay);
            int CurrentRow = GetNodeParentRow(Row);

            while (CurrentRow > -1)
            {
                NodePath = GetCell(CurrentRow, Column, CellProperty.TextDisplay) + GUI.GridDelimiter + NodePath;
                CurrentRow = GetNodeParentRow(CurrentRow);
            }

            return NodePath;
        }

        private int GetNodeParentRow(int rowIndex)
        {
            int titleRows = TitleRows();
            rowIndex -= titleRows;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Parent", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Index", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic parentIndex = GUI.m_APE.GetValueFromMessage();

            if (parentIndex == null)
            {
                parentIndex = -1;
            }
            else
            {
                parentIndex += titleRows;
            }
            
            return parentIndex;
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
            string columnText = GetCellRange(0, columnIndex, this.Rows() - 1, columnIndex);

            char[] separator = { '\r' };
            string[] columnArray = columnText.Split(separator);

            for (int rowIndex = startAtRow; rowIndex < columnArray.GetLength(0); rowIndex++)
            {
                if (rowText == columnArray[rowIndex])
                {
                    return rowIndex;
                }
            }

            return -1;
        }

        /// <summary>
        /// Returns whether the specified row is a child row
        /// </summary>
        /// <param name="rowIndex">The row to check</param>
        /// <returns>True or False</returns>
        public bool RowIsChild(int rowIndex)
        {
            rowIndex -= TitleRows();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "IsChild", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isChild = GUI.m_APE.GetValueFromMessage();

            return isChild;
        }

        /// <summary>
        /// Returns whether the total row is visible or not
        /// </summary>
        /// <returns>True or False</returns>
        public bool TotalRowVisible()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsTotalsRowEnabled", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "IsTotalsRowShown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isEnabled = GUI.m_APE.GetValueFromMessage();
            bool isShown = GUI.m_APE.GetValueFromMessage();

            return isEnabled && isShown;
        }

        /// <summary>
        /// Returns the index of the total row
        /// </summary>
        /// <returns>The total row index</returns>
        public int TotalRow()
        {
            int row = TitleRows();

            if (TotalRowVisible())
            {
                if (FilterRowVisible())
                {
                    row++;
                }

                return row;
            }
            else
            {
                throw new Exception("Total row is not visible in the " + Description);
            }
        }

        /// <summary>
        /// Returns whether the filter row is visible or not
        /// </summary>
        /// <returns>True or False</returns>
        public bool FilterRowVisible()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsFiltersRowEnabled", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "IsFiltersRowShown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isEnabled = GUI.m_APE.GetValueFromMessage();
            bool isShown = GUI.m_APE.GetValueFromMessage();

            return isEnabled && isShown;
        }

        /// <summary>
        /// Returns the index of the filter row
        /// </summary>
        /// <returns>The filter row index</returns>
        public int FilterRow()
        {
            int row = TitleRows();

            if (FilterRowVisible())
            {
                return row;
            }
            else
            {
                throw new Exception("Filter row is not visible in the " + Description);
            }
        }

        /// <summary>
        /// Returns whether the multi edit row is visible or not
        /// </summary>
        /// <returns>True or False</returns>
        public bool MultiEditRowVisible()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsMultiEditRowEnabled", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "IsMultiEditRowShown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isEnabled = GUI.m_APE.GetValueFromMessage();
            bool isShown = GUI.m_APE.GetValueFromMessage();

            return isEnabled && isShown;
        }

        /// <summary>
        /// Returns the index of the multi edit row
        /// </summary>
        /// <returns>The multi edit row index</returns>
        public int MultiEditRow()
        {
            int row = TitleRows();

            if (MultiEditRowVisible())
            {
                if (FilterRowVisible())
                {
                    row++;
                }

                if (TotalRowVisible())
                {
                    row++;
                }

                return row;
            }
            else
            {
                throw new Exception("Multi edit row is not visible in the " + Description);
            }
        }

        /// <summary>
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        internal override void ShowCellInternal(int rowIndex, int columnIndex)
        {
            rowIndex -= TitleRows();
            if (rowIndex < 0)
            {
                // Title row
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageGridControlEnsureTitleCellVisible(DataStores.Store0, columnIndex);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
            }
            else
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "EnsureCellVisible", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
            }
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
                throw new Exception("Must supply a row index greater than 0 in the " + Description);
            }

            if (columnIndex < 0)
            {
                throw new Exception("Must supply a column index greater than 0 in the " + Description);
            }

            // Check to make sure the row isn't hidden
            if (IsRowHidden(rowIndex))
            {
                throw new Exception("Row is hidden in the " + Description);
            }

            // Check to make sure the column isn't hidden
            if (IsColumnHidden(columnIndex))
            {
                throw new Exception("Column is hidden in the " + Description);
            }

            Input.SetFocus(IntPtr.Zero, Identity.ParentHandle, Identity.Description);

            // Scroll the cell into view
            Rectangle CellRectangle;
            
            ShowCell(rowIndex, columnIndex);
            CellRectangle = GetCellRectangle(rowIndex, columnIndex);

            // Check the cell can be scrolled into view
            if (CellRectangle.Top > Height)
            {
                throw new Exception("Row can not be scrolled into view (below bottom) in the " + Description);
            }

            if (CellRectangle.Left > Width)
            {
                throw new Exception("Column can not be scrolled into view (after right) in the " + Description);
            }

            if (CellRectangle.Top < 0)
            {
                throw new Exception("Row can not be scrolled into view (above top) in the " + Description);
            }

            if (CellRectangle.Left < 0)
            {
                throw new Exception("Column can not be scrolled into view (before left) in the " + Description);
            }
            
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
                case CellClickLocation.ChildColumnPicker:
                    Rectangle childColumnPicker = GetColumnGroupPickChildColsToDisplayIconRect(rowIndex, columnIndex);
                    Location.X = childColumnPicker.Left + (childColumnPicker.Width / 2);
                    Location.Y = childColumnPicker.Top + (childColumnPicker.Height / 2);
                    break;
                //TODO
                //case CellClickLocation.ExpandCollapseIconOfCell:
                //    Location.X = CellRectangle.Left + 5 + 2 + BorderWidth() + (TreeViewIndent() * NodeLevel(Row));
                //    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                //    break;
                //case CellClickLocation.LeftSideOfTreeItem:
                //    Location.X = CellRectangle.Left + 5 + 2 + BorderWidth() + (TreeViewIndent() * (NodeLevel(Row) + 1));
                //    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                //    break;
                default:
                    throw new Exception("Implement for CellClickLocation: " + locationInCell.ToString());
            }

            // Handle partial rows / columns
            if (Location.X >= Width)
            {
                Location.X = Width - 1;
            }

            if (Location.Y >= Height)
            {
                Location.Y = Height - 1;
            }

            return Location;
        }

        private Rectangle GetColumnGroupPickChildColsToDisplayIconRect(int rowIndex, int columnIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "DataColumns", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store4, "View", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "GetColumnGroupPickChildColsToDisplayIconRect", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store3), new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "X", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Y", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store8, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store9, "Height", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int x = GUI.m_APE.GetValueFromMessage();
            int y = GUI.m_APE.GetValueFromMessage();
            int width = GUI.m_APE.GetValueFromMessage();
            int height = GUI.m_APE.GetValueFromMessage();
        
            Rectangle cellRectangle = new Rectangle(x, y, width, height);
            return cellRectangle;
        }

        internal override Rectangle GetCellRectangleInternal(int rowIndex, int columnIndex)
        {
            int titleRows = TitleRows();
            rowIndex -= titleRows;

            int x = -1;
            int y = -1;
            int width = -1;
            int height = -1;

            if (rowIndex < 0)
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Columns", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "DataColumns", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndex));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "ViewWidth", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "IsFrozenColumn", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store7, "View", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "ColHeaderHeight", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store9, "CustomScrollPositionX", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                x = GUI.m_APE.GetValueFromMessage();
                width = GUI.m_APE.GetValueFromMessage();
                bool isFrozen = GUI.m_APE.GetValueFromMessage();
                height = GUI.m_APE.GetValueFromMessage();
                int customScrollPositionX = GUI.m_APE.GetValueFromMessage();
                if (!isFrozen)
                {
                    x = x + customScrollPositionX;
                }
                y = (rowIndex + titleRows) * height;
            }
            else
            {
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellDisplayBounds", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
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

            Rectangle cellRectangle = new Rectangle(x, y, width, height);
            return cellRectangle;
        }

        /// <summary>
        /// Returns the namespace and name of the editor the cell uses
        /// </summary>
        /// <param name="row">The row index of the cell to get the editor of</param>
        /// <param name="column">The column index of the cell to get the editor of</param>
        /// <returns></returns>
        public string GetEdititorType(int row, int column)
        {
            int titleRows = TitleRows();
            row -= titleRows;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "LatentZero.Capstone.Controls.GridControl.CellPosition", MemberTypes.Constructor, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetEditDescriptor", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store3));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Editor", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "GetType", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "Namespace", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store8, "Name", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            string Namespace = GUI.m_APE.GetValueFromMessage();
            string Name = GUI.m_APE.GetValueFromMessage();

            return Namespace + "." + Name;
        }

        private IntPtr GetRenderedControlHandle(int row, int column)
        {
            int titleRows = TitleRows();
            row -= titleRows;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "LatentZero.Capstone.Controls.GridControl.CellPosition", MemberTypes.Constructor, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetEditDescriptor", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store3));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Editor", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "GetControl", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store6, DataStores.Store7, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            IntPtr controlHandle = GUI.m_APE.GetValueFromMessage();
            return controlHandle;
        }

        private IntPtr GetNonRenderedControlHandle(int row, int column)
        {
            int titleRows = TitleRows();
            row -= titleRows;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "LatentZero.Capstone.Controls.GridControl.CellPosition", MemberTypes.Constructor, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetEditDescriptor", MemberTypes.Method, new Parameter(GUI.m_APE, DataStores.Store3));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Editor", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            IntPtr controlHandle = GUI.m_APE.GetValueFromMessage();
            return controlHandle;
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
                throw new Exception("Must supply a row index greater than 0 in the " + Description);
            }

            if (columnIndex < 0)
            {
                throw new Exception("Must supply a column index greater than 0 in the " + Description);
            }

            if (FilterRowVisible() && rowIndex == FilterRow())
            {
                throw new Exception("Use SetFilterCellValue to set a filter value in the " + Description);
            }

            string currentValue = this.GetCell(rowIndex, columnIndex);
            T currentValueT = (T)Convert.ChangeType(currentValue, typeof(T));

            // Check if the cell is already set to the correct value
            if (compareMethod == ComparisonMethod.DoNotCompare)
            {
                //we have a string and don't want to do a strict comparison so do some massaging on it to see if its already set
                if (currentValueT is string)
                {
                    if (Microsoft.VisualBasic.Information.IsNumeric(currentValueT))
                    {
                        //numeric value in the string so comvert to a decimal and compare
                        decimal currentValueAsDecimal = Convert.ToDecimal(currentValueT);
                        decimal expectedValueAsDecimal = Convert.ToDecimal(expectedValue);
                        if (currentValueAsDecimal == expectedValueAsDecimal)
                        {
                            GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                            return false;
                        }
                    }
                    else
                    {
                        //non numeric string so just do a trim and check if its equal
                        if ((currentValueT as string).Trim() == (expectedValue as string).Trim())
                        {
                            GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                            return false;
                        }
                    }
                }
                else
                {
                    if (EqualityComparer<T>.Default.Equals(currentValueT, expectedValue))
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                        return false;
                    }
                }
            }
            else
            {
                if (EqualityComparer<T>.Default.Equals(currentValueT, expectedValue))
                {
                    GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                    return false;
                }
            }

            // Wait for the cell to be editable
            timer = Stopwatch.StartNew();
            while (true)
            {
                bool editable = IsCellEditable(rowIndex, columnIndex);
                if (editable)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
                {
                    throw new Exception("The cell for row " + rowFriendlyText + " column " + columnFriendlyText + " is not editable in the " + Description);
                }

                Thread.Sleep(50);
            }

            // TODO get this fixed
            // Work around a bug in the elementstripgrid
            if (columnIndex > 0)
            {
                ShowCell(rowIndex, columnIndex - 1);
            }
            if (columnIndex + 1 < Columns())
            {
                ShowCell(rowIndex, columnIndex + 1);
            }

            // Get the editor type for the cell
            string editorType = GetEdititorType(rowIndex, columnIndex);
            IntPtr EditorHandle = IntPtr.Zero;

            // Select the cell we are trying to edit
            switch (editorType)
            {
                case "LatentZero.Utility.Controls.CheckStateControl.CheckStateRenderer":
                    // Do nothing special for this type of cell
                    break;
                default:
                    // Select the cell if its not selected
                    if (this.CursorCellRow() == rowIndex && this.CursorCellColumn() == columnIndex)
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is selected", LogItemType.Action);
                    }
                    else
                    {
                        GUI.Log("Single " + MouseButton.Left.ToString() + " click on the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                        SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);

                        //Wait for cell selection and paint to happen
                        timer = Stopwatch.StartNew();
                        while (true)
                        {
                            if (this.CursorCellRow() == rowIndex && this.CursorCellColumn() == columnIndex)
                            {
                                break;
                            }

                            if (timer.ElapsedMilliseconds > 3000)
                            {
                                throw new Exception("row " + rowFriendlyText + " column " + columnFriendlyText + " failed be selected for the " + Description);
                            }

                            Thread.Sleep(15);
                        }

                        // TODO is this actually needed run some tests
                        GUI.m_APE.AddFirstMessagePeakMessage(this.Identity.Handle);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemType.Action);
                    base.SendKeysInternal("{F2}");
                    Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);

                    // Reget the editor type now the cell is in edit mode
                    timer = Stopwatch.StartNew();
                    while (true)
                    {
                        editorType = GetEdititorType(rowIndex, columnIndex);

                        if (editorType != ".")
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > 3000)
                        {
                            throw new Exception("Failed to get editor type for the " + Description);
                        }

                        Thread.Sleep(50);
                    }
                    break;
            }

            // Set the cell value
            string valueText = value.ToString();
            switch (editorType)
            {
                case "LatentZero.Utility.Controls.ComboBoxRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(rowIndex, columnIndex);
                    GUIComboBox comboBox = new GUIComboBox(ParentForm, Identity.Description + " combobox", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    comboBox.SingleClickItem(valueText);
                    break;
                case "LatentZero.Utility.Controls.CheckBoxComboBoxRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(rowIndex, columnIndex);
                    GUIComboBox checkBoxComboBox = new GUIComboBox(ParentForm, Identity.Description + " combobox", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    checkBoxComboBox.SingleClickItem(valueText);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                case "LatentZero.Utility.Controls.DatePickerRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(rowIndex, columnIndex);
                    GUIDateTimePicker datePicker = new GUIDateTimePicker(ParentForm, Identity.Description + " date picker", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    datePicker.SetText(valueText);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                case "LatentZero.Utility.Controls.CheckStateControl.CheckStateRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(rowIndex, columnIndex);
                    GUICheckStateControl checkStateControl = new GUICheckStateControl(ParentForm, Identity.Description + " check state control", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item (it is rendered so we need to click on the grid and it processes mouse down events)
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                    Point Location = GetLocationInCell(rowIndex, columnIndex, CellClickLocation.CentreOfCell);
                    base.MouseDownInternal(Location.X, Location.Y, MouseButton.Left, MouseKeyModifier.None);
                    break;
                case "LatentZero.Capstone.Controls.Walker.WalkerControl":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetNonRenderedControlHandle(rowIndex, columnIndex);
                    GUIGenericWalker genericWalker = new GUIGenericWalker(ParentForm, Identity.Description + " generic walker", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    genericWalker.SetText(valueText);
                    break;
                case "LatentZero.Capstone.Controls.TextValueWithTypeControl.TextValueWithTypeControl":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetNonRenderedControlHandle(rowIndex, columnIndex);
                    GUITextValueWithTypeControl textValueWithTypeControl = new GUITextValueWithTypeControl(ParentForm, Identity.Description + " text value with type control", new Identifier(Identifiers.Handle, EditorHandle));
                    // split the item up into type and text
                    char[] splitSeparator = { '|' };
                    string[] textType = valueText.Split(splitSeparator);
                    // Select the type
                    textValueWithTypeControl.SetType(textType[1]);
                    // Select the item
                    textValueWithTypeControl.SetText(textType[0]);
                    // Submit the change
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                case "System.Windows.Forms.TextBox":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetNonRenderedControlHandle(rowIndex, columnIndex);
                    GUITextBox textBox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    textBox.SetText(valueText);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                default:
                    throw new Exception("Unsupported element strip grid editor: Type: " + editorType);
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
                        currentValue = this.GetCell(rowIndex, columnIndex);
                        currentValueT = (T)Convert.ChangeType(currentValue, typeof(T));
                        if (EqualityComparer<T>.Default.Equals(currentValueT, expectedValue))
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
                            throw new Exception("Failed to set the " + Description + " cell value");
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
                    throw new Exception("Unsupported compare method: " + compareMethod.ToString());
            }

            return true;
        }

        /// <summary>
        /// Sets the filter for the specified column of the grid
        /// </summary>
        /// <param name="columnText">The column to set the filter for</param>
        /// <param name="filterItems">The items to select for the filter</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetFilterCellValue(string columnText, params string[] filterItems)
        {
            int rowIndex = FilterRow();
            int columnIndex = FindColumn(columnText);
            StringBuilder filter = new StringBuilder(1024);

            // Check if it needs to be set
            string currentValue = GetCell(rowIndex, columnIndex);
            bool alreadySet = true;
            for (int i = 0; i < filterItems.GetLength(0); i++)
            {
                // Check for All
                if (filterItems[i] == "[All]")
                {
                    if (currentValue == "")
                    {
                        filter.Clear();
                        filter.Append("[All],");
                        break;
                    }
                }

                if (currentValue == filterItems[i])
                {
                }
                else if (currentValue.StartsWith(filterItems[i] + ","))
                {
                }
                else if (currentValue.EndsWith("," + filterItems[i]))
                {
                }
                else if (currentValue.Contains("," + filterItems[i] + ","))
                {
                }
                else
                {
                    alreadySet = false;
                    break;
                }

                filter.Append(filterItems[i]);
                filter.Append(",");
            }
            if (filter.Length > 0)
            {
                filter.Length--;
            }

            if (alreadySet)
            {
                GUI.Log("Ensure the filter for column " + columnText + " is set to " + filter.ToString() + " in the " + Description, LogItemType.Action);
                return false;
            }

            GUIComboBox filterCombobox;
            GUIForm checkedListBoxForm;
            GUICheckedListBox checkedListBox;
            GUITextBox filterTextBox;
            Point location;

            if (CursorCellRow() == rowIndex && CursorCellColumn() == columnIndex)
            {
                // The filter cell is selected
                GUI.Log("Ensure the filter cell for column " + columnText + " is selected in the " + Description, LogItemType.Action);
            }
            else
            {
                // Select the filter cell
                SingleClickCell(rowIndex, columnText, MouseButton.Left, CellClickLocation.CentreOfCell);
            }

            // Display the drop down
            GUI.Log("Single Left click on the right side of the " + Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            location = GetLocationInCell(rowIndex, columnIndex, CellClickLocation.RightSideOfCell);
            MouseDownInternal(location.X, location.Y, MouseButton.Left, MouseKeyModifier.None);

            filterCombobox = new GUIComboBox(ParentForm, Description + " filter combobox", new Identifier(Identifiers.TypeName, "CheckedComboBox"));
            filterTextBox = new GUITextBox(ParentForm, Description + " filter textbox", new Identifier(Identifiers.ChildOf, filterCombobox), new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.TypeName, "Edit"));
            checkedListBoxForm = new GUIForm(Description + " filter checkedlistbox form", new Identifier(Identifiers.Name, "ccbParent"));
            checkedListBox = new GUICheckedListBox(checkedListBoxForm, Description + " filter checkedlistbox", new Identifier(Identifiers.Name, "cclb"));

            // Clear the current filter
            if (currentValue != "")
            {
                checkedListBox.SingleClickItem("[All]");
                //filterTextBox.SingleClick();  // TODO look into this as it causes 100% cpu (mouse down over textbox then move mouse over listbox and mouse up)
                checkedListBox.Type("{Enter}");

                if (filter.ToString() != "[All]")
                {
                    // Display the drop down
                    GUI.Log("Single Left click on the " + Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
                    location = GetLocationInCell(rowIndex, columnIndex, CellClickLocation.RightSideOfCell);
                    MouseDownInternal(location.X, location.Y, MouseButton.Left, MouseKeyModifier.None);

                    filterCombobox = new GUIComboBox(ParentForm, Description + " filter combobox", new Identifier(Identifiers.TypeName, "CheckedComboBox"));
                    filterTextBox = new GUITextBox(ParentForm, Description + " filter textbox", new Identifier(Identifiers.ChildOf, filterCombobox), new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.TypeName, "Edit"));
                    checkedListBoxForm = new GUIForm(Description + " filter checkedlistbox form", new Identifier(Identifiers.Name, "ccbParent"));
                    checkedListBox = new GUICheckedListBox(checkedListBoxForm, Description + " filter checkedlistbox", new Identifier(Identifiers.Name, "cclb"));
                }
            }

            // Check the filter items
            if (filter.ToString() != "[All]")
            {
                // Check the filter items
                for (int i = 0; i < filterItems.GetLength(0); i++)
                {
                    checkedListBox.ItemCheck(filterItems[i]);
                }
                //filterTextBox.SingleClick();
                checkedListBox.Type("{Enter}");
            }

            return true;
        }

        //public bool SetTotalCellValue()
        //{
        //    return false;
        //}

        //public bool SetMultiEditCellValue()
        //{
        //    return false;
        //}

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r where
        /// each cell has a width and height greater than 1 pixel and the row or column is not
        /// hidden.  Collapsed nodes are also excluded.
        /// </summary>
        /// <returns>A string containing the range of values</returns>
        public string GetAllVisibleCells()
        {
            string[] separatorComma = { "," };
            string[] separatorCr = { "\r" };
            string[] separatorTab = { "\t" };

            bool[] visibleColumns = GetColumnsVisibilityState();

            int columns = this.Columns();
            int rows = this.Rows();

            // Copy the whole grid
            if (rows > 0 && columns > 0)
            {
                string fullGrid = GetCellRange(0, 0, rows - 1, columns - 1);

                StringBuilder grid = new StringBuilder(10240);
                bool doneColumn;

                string[] fullGridRows = fullGrid.Split(separatorCr, StringSplitOptions.None);
                for (int fullGridRow = 0; fullGridRow < rows; fullGridRow++)
                {
                    if (true)   // All rows
                    {
                        string[] fullGridColumns = fullGridRows[fullGridRow].Split(separatorTab, StringSplitOptions.None);
                        doneColumn = false;
                        for (int fullGridColumn = 0; fullGridColumn < columns; fullGridColumn++)
                        {
                            if (visibleColumns[fullGridColumn])
                            {
                                if (doneColumn)
                                {
                                    grid.Append("\t");
                                }
                                grid.Append(fullGridColumns[fullGridColumn]);
                                doneColumn = true;
                            }
                        }
                        grid.Append("\r");
                    }
                }

                // Strip off the final \r
                grid.Length -= 1;

                return grid.ToString();
            }
            else
            {
                return "";
            }
        }

        /// <summary>
        /// Gets the number of cells in the grid which are currently animating
        /// </summary>
        /// <returns>The number of cells currently animating</returns>
        public int AnimationCount()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "View", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "AnimationController", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "m_animationKeyHash", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            int animationCount = GUI.m_APE.GetValueFromMessage();
            return animationCount;
        }

        /// <summary>
        /// Determines if the specified cell is animating
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to check if its animating</param>
        /// <param name="columnText">The column of the cell to check if its animating</param>
        /// <returns>True if the cell is currently animating</returns>
        public bool IsCellAnimating(int rowIndex, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            return IsCellAnimating(rowIndex, columnIndex);
        }

        /// <summary>
        /// Determines if the specified cell is animating
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to check if its animating</param>
        /// <param name="columnIndex">The column index of the cell to check if its animating</param>
        /// <returns>True if the cell is currently animating</returns>
        public bool IsCellAnimating(int rowIndex, int columnIndex)
        {
            int row = rowIndex - TitleRows();

            if (row < 0)
            {
                throw new Exception("Title rows are not currently supported");
            }

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetRow", MemberTypes.Method, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "UniqueId", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            int uniqueRowId = GUI.m_APE.GetValueFromMessage();

            string key = "CVC" + string.Concat(uniqueRowId.ToString(), "_", columnIndex.ToString());

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "View", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "AnimationController", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "m_animationKeyHash", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageDictionaryContainsKey(DataStores.Store3, DataStores.Store4, key);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            bool isCellAnimating = GUI.m_APE.GetValueFromMessage();;
            return isCellAnimating;
        }
    }
}