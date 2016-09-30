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
using System.Xml;

namespace APE.Language
{
    // TODO this should probably extend GridControl or EditorStrip with most of the methods in there?

    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Common.ModelControls.ElementStrip.ElementStripGrid
    /// </summary>
    public sealed class GUIElementStripGrid : GUIFocusableObject
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetColumnInfoXML", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string ColumnInfoXML = GUI.m_APE.GetValueFromMessage();

            XmlDocument columnDocument = new XmlDocument();
            columnDocument.LoadXml(ColumnInfoXML);

            int maxLevels = -1;
            int currentLevel = -1;
            foreach (XmlNode node in columnDocument.SelectSingleNode("Columns").ChildNodes)
            {
                if (int.TryParse(node.Attributes.GetNamedItem("ColumnHeaderLevel").Value, out currentLevel))
                {
                    if (currentLevel > maxLevels)
                    {
                        maxLevels = currentLevel;
                    }
                }
            }

            return maxLevels + 1;
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
        /// Returns the number of fixed / frozen rows, that is a row which doesn't scroll, in the grid
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
            int FixedRows = GUI.m_APE.GetValueFromMessage();

            return FixedRows + TitleRows();
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

            char[] splitSeparator = { ',' };
            string[] selectedRowIndexesStringArray = selectedRowIndexesString.Split(splitSeparator);
            int[] selectedRowIndexesintArray = Array.ConvertAll(selectedRowIndexesStringArray, s => int.Parse(s) + titleRows);

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

        private void ConvertXMLTreeTo2dArray(XmlNodeList nodes, ref int column, ref string[,] titles, int maxLevels, bool visibleOnly)
        {
            foreach (XmlNode node in nodes)
            {
                if (node.Name == "Column")
                {
                    int level = int.Parse(node.Attributes.GetNamedItem("ColumnHeaderLevel").Value);
                    bool visible = bool.Parse(node.Attributes.GetNamedItem("IsVisible").Value);

                    if (visibleOnly)
                    {
                        if (visible)
                        {
                            titles[maxLevels - level, column] = node.Attributes.GetNamedItem("ColumnTitle").Value;
                        }
                    }
                    else
                    {
                        titles[maxLevels - level, column] = node.Attributes.GetNamedItem("ColumnTitle").Value;
                    }
                }

                if (node.HasChildNodes)
                {
                    ConvertXMLTreeTo2dArray(node.ChildNodes, ref column, ref titles, maxLevels, visibleOnly);
                }
                else
                {
                    if (node.Name == "Column")
                    {
                        bool visible = bool.Parse(node.Attributes.GetNamedItem("IsVisible").Value);
                        if (visibleOnly)
                        {
                            if (visible)
                            {
                                column++;
                            }
                        }
                        else
                        {
                            column++;
                        }
                    }
                }
            }
        }

        private string[,] GetColumnTitles(bool visibleOnly)
        {
            //Get an XML representation of the columns
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetColumnInfoXML", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string xmlColumns = GUI.m_APE.GetValueFromMessage();

            //Load the xml into an xml document
            XmlDocument columnsXMLDocument = new XmlDocument();
            columnsXMLDocument.LoadXml(xmlColumns);

            //Work out how many levels we have
            int maxLevels = -1;
            foreach (XmlNode node in columnsXMLDocument.SelectSingleNode("Columns").ChildNodes)
            {
                int level = int.Parse(node.Attributes.GetNamedItem("ColumnHeaderLevel").Value);
                if (level > maxLevels)
                {
                    maxLevels = level;
                }
            }

            //Workout how many columns we have
            int maxColumns = -1;
            foreach (XmlNode node in columnsXMLDocument.SelectNodes(@".//Column"))
            {
                bool visible = bool.Parse(node.Attributes.GetNamedItem("IsVisible").Value);
                int level = int.Parse(node.Attributes.GetNamedItem("ColumnHeaderLevel").Value);

                if (visibleOnly)
                {
                    if (visible)
                    {
                        if (level == 0)
                        {
                            maxColumns++;
                        }
                    }
                }
                else
                {
                    if (level == 0)
                    {
                        maxColumns++;
                    }
                }
            }

            //Build an array of the correct size
            string[,] titles = new string[maxLevels + 1, maxColumns + 1];

            for (int y = 0; y < titles.GetLength(0); y++)
            {
                for (int x = 0; x < titles.GetLength(1); x++)
                {
                    titles[y, x] = "";
                }
            }

            //Convert to a 2d array
            int column = 0;
            ConvertXMLTreeTo2dArray(columnsXMLDocument.SelectSingleNode("Columns").ChildNodes, ref column, ref titles, maxLevels, visibleOnly);

            return titles;
        }

        /// <summary>
        /// Returns true if the specified column in the grid exists
        /// </summary>
        /// <param name="columnToFind">Column to check if hidden delimited by -> for example Order -> Id</param>
        /// <returns>True or False</returns>
        public bool ColumnExists(string columnToFind)
        {
            string[] delimiter = { " -> " };
            string[] columnHeader = columnToFind.Split(delimiter, StringSplitOptions.None);

            if (FindColumnInternal(columnHeader) == -1)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Returns the index of the specified column in the grid
        /// </summary>
        /// <param name="columnToFind">Column to check if hidden delimited by -> for example Order -> Id</param>
        /// <returns>The index of the column</returns>
        public int FindColumn(string columnToFind)
        {
            string[] delimiter = { " -> " };
            string[] columnHeader = columnToFind.Split(delimiter, StringSplitOptions.None);

            return FindColumn(columnHeader);
        }

        /// <summary>
        /// Returns the index of the specified column in the grid
        /// </summary>
        /// <param name="columnHeader">The column to check</param>
        /// <returns>The index of the column</returns>
        public int FindColumn(string[] columnHeader)
        {
            int column = -1;

            // Columns present may change so try twice
            try
            {
                column = FindColumnInternal(columnHeader);
            }
            catch
            {
                column = FindColumnInternal(columnHeader);
            }

            if (column == -1)
            {
                throw new Exception("Failed to find column");
            }

            return column;
        }

        private int FindColumnInternal(string[] columnHeader)
        {
            string[,] titles = GetColumnTitles(false);

            //Search for the column
            bool Found = false;
            int columns = titles.GetLength(1);
            int row = 0;
            int column = -1;
            for (column = 0; column < columns; column++)
            {
                if (!this.IsColumnHidden(column))
                {
                    if (titles[columnHeader.GetLength(0) - 1, column] == columnHeader[columnHeader.GetLength(0) - 1])
                    {
                        Found = true;
                        for (row = columnHeader.GetLength(0) - 2; row > -1; row--)
                        {
                            int tempColumn = column;

                            if (columnHeader[row] != "")
                            {
                                while (titles[row, tempColumn] == "")
                                {
                                    tempColumn--;
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

        private bool[] GetColumnsVisibilityState()
        {
            int columns = this.Columns();

            //Build a boolean array big enough to hold an entry for each column
            bool[] visibleColumns = new bool[columns];

            //Get an XML representation of the columns
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetColumnInfoXML", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string xmlColumns = GUI.m_APE.GetValueFromMessage();

            //Load the xml into an xml document
            XmlDocument columnsXMLDocument = new XmlDocument();
            columnsXMLDocument.LoadXml(xmlColumns);

            //Workout which columns are visible
            int column = 0;
            foreach (XmlNode node in columnsXMLDocument.SelectNodes(".//Column"))
            {
                if (node.Attributes.GetNamedItem("ColumnHeaderLevel").Value == "0")
                {
                    if (node.Attributes.GetNamedItem("IsVisible").Value == "True")
                    {
                        visibleColumns[column] = true;
                    }
                    column++;
                }
            }

            return visibleColumns;
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
        /// <param name="columnToFind">Column to check if hidden delimited by -> for example Order -> Id</param>
        /// <returns>True or False</returns>
        public bool IsColumnHidden(string columnToFind)
        {
            int column = FindColumn(columnToFind);
            return IsColumnHidden(column);
        }

        /// <summary>
        /// Returns true if the specified column is hidden in the grid
        /// </summary>
        /// <param name="column">Column index to check if hidden</param>
        /// <returns>True or False</returns>
        public bool IsColumnHidden(int column)
        {
            bool[] visibleColumns = GetColumnsVisibilityState();

            if (visibleColumns[column])
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
        /// <param name="row">Row to check if hidden</param>
        /// <returns>True or False</returns>
        public bool IsRowHidden(string row)
        {
            int RowNumber = FindRow(row);
            return IsRowHidden(RowNumber);
        }

        /// <summary>
        /// Returns true if the specified row is hidden in the grid
        /// </summary>
        /// <param name="row">Row index to check if hidden</param>
        /// <returns>True or False</returns>
        public bool IsRowHidden(int row)
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
        /// <param name="column1Text">The start column of the range delimited by -> for example Order -> Id</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Text">The end column of the range delimited by -> for example Order -> Id</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, string column1Text, int row2Index, string column2Text)
        {
            int column1Index = FindColumn(column1Text);
            int column2Index = FindColumn(column2Text);
            return GetCellRange(row1Index, column1Index, row2Index, column2Index);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Index">The start column of the range</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Text">The end column of the range delimited by -> for example Order -> Id</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, int column1Index, int row2Index, string column2Text)
        {
            int column2Index = FindColumn(column2Text);
            return GetCellRange(row1Index, column1Index, row2Index, column2Index);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Text">The start column of the range delimited by -> for example Order -> Id</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Index">The end column of the range</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, string column1Text, int row2Index, int column2Index)
        {
            int column1Index = FindColumn(column1Text);
            return GetCellRange(row1Index, column1Index, row2Index, column2Index);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Index">The start column of the range</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Index">The end column of the range</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, int column1Index, int row2Index, int column2Index)
        {
            string text = "";

            //Get how many column title rows we have
            int titleRows = TitleRows();

            //Get the title rows
            string[,] columnTitle = GetColumnTitles(false);

            //If the range we want includes the title rows
            if (row1Index < titleRows)
            {
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
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnText">Column of the cell delimited by -> for example Order -> Id</param>
        /// <returns></returns>
        public string GetCellValue(string rowText, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            return GetCellRange(rowIndex, columnIndex, rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowIndex">Row of the cell</param>
        /// <param name="columnText">Column of the cell delimited by -> for example Order -> Id</param>
        /// <returns>The cell value as a string</returns>
        /// <returns></returns>
        public string GetCellValue(int rowIndex, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            return GetCellRange(rowIndex, columnIndex, rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <returns></returns>
        public string GetCellValue(string rowText, int columnIndex)
        {
            int rowIndex = FindRow(rowText, columnIndex);
            return GetCellRange(rowIndex, columnIndex, rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <returns>The cell value as a string</returns>
        public string GetCellValue(int rowIndex, int columnIndex)
        {
            return GetCellRange(rowIndex, columnIndex, rowIndex, columnIndex);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the first visible column
        /// </summary>
        /// <param name="rowText">The value to look for in the first visible column or the treeview column if the grid is a tree</param>
        /// <returns>The index of the row or -1</returns>
        public int FindRow(string rowText)
        {
            int startAtRow = 0;
            return FindRow(rowText, -1, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnText">The column to look for the value in delimited by -> for example Order -> Id</param>
        /// <returns>The index of the row or -1</returns>
        public int FindRow(string rowText, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            int startAtRow = 0;
            return FindRow(rowText, columnIndex, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnText">The column to look for the value in delimited by -> for example Order -> Id</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row or -1</returns>
        public int FindRow(string rowText, string columnText, int startAtRow)
        {
            int columnIndex = FindColumn(columnText);
            return FindRow(rowText, columnIndex, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnIndex">The column to look for the value in</param>
        /// <returns>The index of the row or -1</returns>
        public int FindRow(string rowText, int columnIndex)
        {
            int startAtRow = 0;
            return FindRow(rowText, columnIndex, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnIndex">The column to look for the value in</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row or -1</returns>
        public int FindRow(string rowText, int columnIndex, int startAtRow)
        {
            int rowIndex;

            if (columnIndex == -1)
            {
                //TODO treeview
                //if (IsTreeView())
                //{
                //    rowIndex = FindNodeRow(rowText);
                //}
                //else
                //{
                columnIndex = FirstVisibleColumn();
                rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
                //}
            }
            else
            {
                rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
            }

            return rowIndex;
        }

        private int FindRowInternal(string rowText, int columnIndex, int startAtRow)
        {
            string columnText = this.GetCellRange(0, columnIndex, this.Rows() - 1, columnIndex);

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
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnText">Column text of the cell delimited by -> for example Order -> Id</param>
        public void Show(int rowIndex, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            Show(rowIndex, columnIndex);
        }

        /// <summary>
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        public void Show(int rowIndex, int columnIndex)
        {
            rowIndex -= TitleRows();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "EnsureCellVisible", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsTotalsRowShown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isVisible = GUI.m_APE.GetValueFromMessage();

            return isVisible;
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
                throw new Exception("Total row is not visible");
            }
        }

        /// <summary>
        /// Returns whether the filter row is visible or not
        /// </summary>
        /// <returns>True or False</returns>
        public bool FilterRowVisible()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsFiltersRowShown", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isVisible = GUI.m_APE.GetValueFromMessage();

            return isVisible;
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
                throw new Exception("Filter row is not visible");
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
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isVisible = GUI.m_APE.GetValueFromMessage();

            return isVisible;
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
                throw new Exception("Multi edit row is not visible");
            }
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void Select(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void Select(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void Select(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void Select(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void Select(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void Select(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void Select(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void Select(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        private void SelectInternal(int row, int column, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point Location = GetLocationInCell(row, column, locationInCell);
            base.MouseSingleClickInternal(Location.X, Location.Y, button, keyModifier);
        }

        /// <summary>
        /// Moves the mouse cursor to the specified location the specified cell
        /// </summary>
        /// <param name="Row">The row index of the cell to move the cursor to</param>
        /// <param name="Column">The column index of the cell to move the cursor to</param>
        /// <param name="LocationInCell">The location in the cell to move the cursor to</param>
        public void MouseMove(int Row, int Column, CellClickLocation LocationInCell)
        {
            Point Location = GetLocationInCell(Row, Column, LocationInCell);
            base.MouseMove(Location.X, Location.Y);
        }

        private Point GetLocationInCell(int rowIndex, int columnIndex, CellClickLocation locationInCell)
        {
            if (rowIndex < 0)
            {
                throw new Exception("Must supply a valid row index: " + rowIndex.ToString());
            }

            if (columnIndex < 0)
            {
                throw new Exception("Must supply a valid column index: " + columnIndex.ToString());
            }

            //Check to make sure the row isn't hidden
            if (IsRowHidden(rowIndex))
            {
                throw new Exception("Row is hidden");
            }

            //Check to make sure the column isn't hidden
            if (IsColumnHidden(columnIndex))
            {
                throw new Exception("Column is hidden");
            }

            //Scroll the cell into view
            Show(rowIndex, columnIndex);

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
                    throw new Exception("Implement for CellClickLocation: " + Location.ToString());
            }

            return Location;
        }

        private Rectangle GetCellRectangle(string row, string column)
        {
            int RowNumber = FindRow(row);
            int ColumnNumber = FindColumn(column);
            return GetCellRectangle(RowNumber, ColumnNumber);
        }

        private Rectangle GetCellRectangle(int row, string column)
        {
            int ColumnNumber = FindColumn(column);
            return GetCellRectangle(row, ColumnNumber);
        }

        private Rectangle GetCellRectangle(string row, int column)
        {
            int RowNumber = FindRow(row);
            return GetCellRectangle(RowNumber, column);
        }

        private Rectangle GetCellRectangle(int row, int column)
        {
            int titleRows = TitleRows();
            row -= titleRows;

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellDisplayBounds", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
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
            int X = GUI.m_APE.GetValueFromMessage();
            int Y = GUI.m_APE.GetValueFromMessage();
            int Width = GUI.m_APE.GetValueFromMessage();
            int Height = GUI.m_APE.GetValueFromMessage();

            Rectangle CellRectangle = new Rectangle(X, Y, Width, Height);
            return CellRectangle;
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

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, string columnText, string value)
        {
            return SetCellValue(rowText, columnText, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, string columnText, string value, string expectedValue)
        {
            return SetCellValue(rowText, columnText, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, string columnText, string value, string expectedValue, string submitKey)
        {
            int row = FindRow(rowText);
            int column = FindColumn(columnText);
            return SetCellValueInternal(rowText, columnText, row, column, value, expectedValue, submitKey);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int row, string columnText, string value)
        {
            return SetCellValue(row, columnText, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int row, string columnText, string value, string expectedValue)
        {
            return SetCellValue(row, columnText, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int row, string columnText, string value, string expectedValue, string submitKey)
        {
            int column = FindColumn(columnText);
            return SetCellValueInternal(row.ToString(), columnText, row, column, value, expectedValue, submitKey);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row of the cell to set the value of</param>
        /// <param name="column">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string row, int column, string value)
        {
            return SetCellValue(row, column, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row of the cell to set the value of</param>
        /// <param name="column">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string row, int column, string value, string expectedValue)
        {
            return SetCellValue(row, column, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="column">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, int column, string value, string expectedValue, string submitKey)
        {
            int row = FindRow(rowText);
            return SetCellValueInternal(rowText, column.ToString(), row, column, value, expectedValue, submitKey);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row index of the cell to set the value of</param>
        /// <param name="column">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int row, int column, string value)
        {
            return SetCellValue(row, column, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row index of the cell to set the value of</param>
        /// <param name="column">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int row, int column, string value, string expectedValue)
        {
            return SetCellValue(row, column, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="row">The row index of the cell to set the value of</param>
        /// <param name="column">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int row, int column, string value, string expectedValue, string submitKey)
        {
            return SetCellValueInternal(row.ToString(), column.ToString(), row, column, value, expectedValue, submitKey);
        }

        private bool SetCellValueInternal(string rowText, string columnText, int row, int column, string value, string expectedValue, string submitKey)
        {
            Stopwatch timer;

            if (expectedValue == null)
            {
                expectedValue = value;
            }

            if (submitKey == null)
            {
                submitKey = "{Enter}";
            }

            // Check if the cell is already set to the correct value
            string CurrentValue = this.GetCellValue(row, column);

            if (CurrentValue == expectedValue)
            {
                GUI.Log("Ensure " + Identity.Description + " row " + rowText + " column " + columnText + " is set to " + expectedValue, LogItemType.Action);
                return false;
            }

            // TODO get this fixed
            // Work around a bug in the elementstripgrid
            if (column > 0)
            {
                Show(row, column - 1);
            }
            if (column + 1 < this.Columns())
            {
                Show(row, column + 1);
            }

            // Get the editor type for the cell
            string editorType = GetEdititorType(row, column);
            IntPtr EditorHandle = IntPtr.Zero;

            // Select the cell we are trying to edit
            switch (editorType)
            {
                case "LatentZero.Utility.Controls.CheckStateControl.CheckStateRenderer":
                    // Do nothing special for this type of cell
                    break;
                default:
                    // Select the cell if its not selected
                    if (this.CursorCellRow() == row && this.CursorCellColumn() == column)
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowText + " column " + columnText + " is selected", LogItemType.Action);
                    }
                    else
                    {
                        GUI.Log("Single " + MouseButton.Left.ToString() + " click on " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
                        this.SelectInternal(row, column, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemType.Action);
                    base.SendKeysInternal("{F2}");
                    Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);

                    // Reget the editor type now the cell is in edit mode
                    editorType = GetEdititorType(row, column);
                    break;
            }

            //Set the cell value
            switch (editorType)
            {
                case "LatentZero.Utility.Controls.ComboBoxRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(row, column);
                    GUIComboBox comboBox = new GUIComboBox(ParentForm, Identity.Description + " combobox", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    comboBox.ItemSelect(value);
                    break;
                case "LatentZero.Utility.Controls.CheckBoxComboBoxRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(row, column);
                    GUIComboBox checkBoxComboBox = new GUIComboBox(ParentForm, Identity.Description + " combobox", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    checkBoxComboBox.ItemSelect(value);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                case "LatentZero.Utility.Controls.DatePickerRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(row, column);
                    GUIDateTimePicker datePicker = new GUIDateTimePicker(ParentForm, Identity.Description + " date picker", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    datePicker.SetText(value);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                case "LatentZero.Utility.Controls.CheckStateControl.CheckStateRenderer":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetRenderedControlHandle(row, column);
                    GUICheckStateControl checkStateControl = new GUICheckStateControl(ParentForm, Identity.Description + " check state control", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item (it is rendered so we need to click on the grid and it processes mouse down events)
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
                    Point Location = GetLocationInCell(row, column, CellClickLocation.CentreOfCell);
                    base.MouseDownInternal(Location.X, Location.Y, MouseButton.Left, MouseKeyModifier.None);
                    break;
                case "LatentZero.Capstone.Controls.Walker.WalkerControl":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetNonRenderedControlHandle(row, column);
                    GUIGenericWalker genericWalker = new GUIGenericWalker(ParentForm, Identity.Description + " generic walker", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    genericWalker.SetText(value);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                case "LatentZero.Capstone.Controls.TextValueWithTypeControl.TextValueWithTypeControl":
                    // Get the handle of the editor for this cell then locate the control
                    EditorHandle = GetNonRenderedControlHandle(row, column);
                    GUITextValueWithTypeControl textValueWithTypeControl = new GUITextValueWithTypeControl(ParentForm, Identity.Description + " text value with type control", new Identifier(Identifiers.Handle, EditorHandle));
                    // split the item up into type and text
                    char[] splitSeparator = { '|' };
                    string[] textType = value.Split(splitSeparator);
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
                    EditorHandle = GetNonRenderedControlHandle(row, column);
                    GUITextBox textBox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, EditorHandle));
                    // Select the item
                    textBox.SetText(value);
                    GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                    base.SendKeysInternal(submitKey);
                    break;
                default:
                    throw new Exception("Unsupported element strip grid editor: Type: " + editorType);
            }

            //Check the value was set
            timer = Stopwatch.StartNew();
            do
            {
                CurrentValue = this.GetCellValue(row, column);

                if (CurrentValue == expectedValue)
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

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to set the element strip grid cell value");
                }

                Thread.Sleep(15);
            }
            while (true);

            return true;
        }

        //TODO
        //public bool SetFilterCellValue()
        //{
        //    return false;
        //}

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
        /// Send the specified text to the control
        /// </summary>
        /// <param name="text">The text to send to the control</param>
        public void Type(string text)
        {
            base.SetFocus();
            base.SendKeys(text);
        }
    }
}