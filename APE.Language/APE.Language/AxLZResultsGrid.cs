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
using System.Diagnostics;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Capstone.ComSupport.ResultsGrid.AxLZResultsGrid
    /// AxDRILLDOWNLib.AxLZResultsGrid
    /// </summary>
    public sealed class GUIAxLZResultsGrid : GUIFocusableObject
    {
        public enum VSFlexgridCellPropertySettings
        {
            /// <summary>
            /// Cells text
            /// </summary>
            flexcpText = 0,
            /// <summary>
            /// Cells text style
            /// </summary>
            flexcpTextStyle = 1,
            /// <summary>
            /// Cells text alignment
            /// </summary>
            flexcpAlignment = 2,
            /// <summary>
            /// Cells picture
            /// </summary>
            flexcpPicture = 3,
            /// <summary>
            /// Cells picture alignment
            /// </summary>
            flexcpPictureAlignment = 4,
            /// <summary>
            /// Cells checkbox state
            /// </summary>
            flexcpChecked = 5,
            /// <summary>
            /// Cells back colour
            /// </summary>
            flexcpBackColor = 6,
            /// <summary>
            /// Cells fore colour
            /// </summary>
            flexcpForeColor = 7,
            /// <summary>
            /// Cells flood percentage
            /// </summary>
            flexcpFloodPercent = 8,
            /// <summary>
            /// cells flood colour
            /// </summary>
            flexcpFloodColor = 9,
            /// <summary>
            /// Cells font
            /// </summary>
            flexcpFont = 10,
            /// <summary>
            /// Cells font name
            /// </summary>
            flexcpFontName = 11,
            /// <summary>
            /// Cells font size
            /// </summary>
            flexcpFontSize = 12,
            /// <summary>
            /// Cells font boldness state
            /// </summary>
            flexcpFontBold = 13,
            /// <summary>
            /// Cells font italic state
            /// </summary>
            flexcpFontItalic = 14,
            /// <summary>
            /// Cells font underline state
            /// </summary>
            flexcpFontUnderline = 15,
            /// <summary>
            /// Cells font strikethru state
            /// </summary>
            flexcpFontStrikethru = 16,
            /// <summary>
            /// Cells font width
            /// </summary>
            flexcpFontWidth = 17,
            /// <summary>
            /// Cells value
            /// </summary>
            flexcpValue = 18,
            /// <summary>
            /// Cells formatted text
            /// </summary>
            flexcpTextDisplay = 19,
            /// <summary>
            /// Cells attached variant
            /// </summary>
            flexcpData = 20,
            /// <summary>
            /// Cells custom format (true or false)
            /// </summary>
            flexcpCustomFormat = 21,
            /// <summary>
            /// Cells left coordinate in twips
            /// </summary>
            flexcpLeft = 22,
            /// <summary>
            /// Cells top coordinate in twips
            /// </summary>
            flexcpTop = 23,
            /// <summary>
            /// Cells width in twips
            /// </summary>
            flexcpWidth = 24,
            /// <summary>
            /// Cells height in twips
            /// </summary>
            flexcpHeight = 25,
            /// <summary>
            /// Cells value as a double or string
            /// </summary>
            flexcpVariantValue = 26,
            /// <summary>
            /// Set to true to force a refresh
            /// </summary>
            flexcpRefresh = 27,
        }

        /// <summary>
        /// The data type of a flexgrid column
        /// </summary>
        public enum VSFlexgridColumnDataType
        {
            /// <summary>
            /// Empty
            /// </summary>
            flexDTEmpty = 0,
            /// <summary>
            /// Null
            /// </summary>
            flexDTNull = 1,
            /// <summary>
            /// Short
            /// </summary>
            flexDTShort = 2,
            /// <summary>
            /// Long
            /// </summary>
            flexDTLong = 3,
            /// <summary>
            /// Single precision float
            /// </summary>
            flexDTSingle = 4,
            /// <summary>
            /// Double precision float
            /// </summary>
            flexDTDouble = 5,
            /// <summary>
            /// Currency
            /// </summary>
            flexDTCurrency = 6,
            /// <summary>
            /// Date
            /// </summary>
            flexDTDate = 7,
            /// <summary>
            /// String
            /// </summary>
            flexDTString = 8,
            /// <summary>
            /// Dispatch
            /// </summary>
            flexDTDispatch = 9,
            /// <summary>
            /// Error
            /// </summary>
            flexDTError = 10,
            /// <summary>
            /// Boolean
            /// </summary>
            flexDTBoolean = 11,
            /// <summary>
            /// Variant
            /// </summary>
            flexDTVariant = 12,
            /// <summary>
            /// Unknown
            /// </summary>
            flexDTUnknown = 13,
            /// <summary>
            /// Decimal
            /// </summary>
            flexDTDecimal = 14,
            /// <summary>
            /// Long8
            /// </summary>
            flexDTLong8 = 20,
            /// <summary>
            /// StringC
            /// </summary>
            flexDTStringC = 30,
            /// <summary>
            /// StringW
            /// </summary>
            flexDTStringW = 31,
        }

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIAxLZResultsGrid(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }


        /// <summary>
        /// Copies the contents of the results grid to the clipboard
        /// </summary>
        public void CopyToClipboard()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "CopyToClipboard", MemberTypes.Method);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "RowHidden", MemberTypes.Property, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool hidden = GUI.m_APE.GetValueFromMessage();
            return hidden;
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
        public bool IsColumnHidden(int column)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ColHidden", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool hidden = GUI.m_APE.GetValueFromMessage();
            return hidden;
        }

        /// <summary>
        /// Determines if the specified cell is currently viewable without scrolling
        /// </summary>
        /// <param name="row">The row index of the cell to check</param>
        /// <param name="column">The column index of the cell to check</param>
        /// <returns>True if the cell is visible without scrolling otherwise false</returns>
        public bool IsCellVisible(int row, int column)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "LeftCol", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "RightCol", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store5, "TopRow", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store6, "BottomRow", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int leftColumn = GUI.m_APE.GetValueFromMessage();
            int rightColumn = GUI.m_APE.GetValueFromMessage();
            int topRow = GUI.m_APE.GetValueFromMessage();
            int bottomRow = GUI.m_APE.GetValueFromMessage();

            if (column <= leftColumn || column >= rightColumn || row <= topRow || row >= bottomRow)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Returns the first visible (non-hidden) row in the grid
        /// </summary>
        /// <returns>The first visible row</returns>
        public int FirstVisibleRow()
        {
            int RowCount = Rows();
            for (int Row = 0; Row < RowCount; Row++)
            {
                if (!IsRowHidden(Row))
                {
                    return Row;
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
            int ColumnCount = Columns();
            for (int Column = 0; Column < ColumnCount; Column++)
            {
                if (!IsColumnHidden(Column))
                {
                    return Column;
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "RowSel", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedRow = GUI.m_APE.GetValueFromMessage();
            return SelectedRow;
        }

        /// <summary>
        /// Returns the selected column
        /// </summary>
        /// <returns>The selected column index</returns>
        public int SelectedColumn()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ColSel", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedColumn = GUI.m_APE.GetValueFromMessage();
            return SelectedColumn;
        }

        /// <summary>
        /// Returns the number of fixed rows, that is a row which doesn't scroll, in the grid
        /// </summary>
        /// <returns>The number of fixed rows</returns>
        public int FixedRows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "FixedRows", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedRows = GUI.m_APE.GetValueFromMessage();
            return FixedRows;
        }

        /// <summary>
        /// Returns the number of fixed columns, that is a column which doesn't scroll, in the grid
        /// </summary>
        /// <returns>The number of fixed columns</returns>
        public int FixedColumns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "FixedCols", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedColumns = GUI.m_APE.GetValueFromMessage();
            return FixedColumns;
        }

        /// <summary>
        /// Returns the number of frozen columns, that is a column which doesn't scroll, in the grid
        /// </summary>
        /// <returns>The number of frozen columns</returns>
        public int FrozenColumns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "FrozenCols", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FrozenColumns = GUI.m_APE.GetValueFromMessage();
            return FrozenColumns;
        }

        /// <summary>
        /// Returns the number of rows in the grid, including those which are hidden
        /// </summary>
        /// <returns>The number of rows</returns>
        public int Rows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Rows", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int rows = GUI.m_APE.GetValueFromMessage();
            return rows;
        }

        /// <summary>
        /// Returns the number of columns, including those which are hidden
        /// </summary>
        /// <returns>The number of columns</returns>
        public int Columns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cols", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int columns = GUI.m_APE.GetValueFromMessage();
            return columns;
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
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int rowIndex, string columnText, string value)
        {
            return SetCellValue(rowIndex, columnText, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int rowIndex, string columnText, string value, string expectedValue)
        {
            return SetCellValue(rowIndex, columnText, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int rowIndex, string columnText, string value, string expectedValue, string submitKey)
        {
            int column = FindColumn(columnText);
            return SetCellValueInternal(rowIndex.ToString(), columnText, rowIndex, column, value, expectedValue, submitKey);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, int columnIndex, string value)
        {
            return SetCellValue(rowText, columnIndex, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, int columnIndex, string value, string expectedValue)
        {
            return SetCellValue(rowText, columnIndex, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(string rowText, int columnIndex, string value, string expectedValue, string submitKey)
        {
            int row = FindRow(rowText);
            return SetCellValueInternal(rowText, columnIndex.ToString(), row, columnIndex, value, expectedValue, submitKey);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int rowIndex, int columnIndex, string value)
        {
            return SetCellValue(rowIndex, columnIndex, value, null, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int rowIndex, int columnIndex, string value, string expectedValue)
        {
            return SetCellValue(rowIndex, columnIndex, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value you expect the cell to display after you have set it</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue(int rowIndex, int columnIndex, string value, string expectedValue, string submitKey)
        {
            return SetCellValueInternal(rowIndex.ToString(), columnIndex.ToString(), rowIndex, columnIndex, value, expectedValue, submitKey);
        }

        private bool SetCellValueInternal(string rowText, string columnText, int rowIndex, int columnIndex, string value, string expectedValue, string submitKey)
        {
            Stopwatch ook = Stopwatch.StartNew();

            if (expectedValue == null)
            {
                expectedValue = value;
            }

            if (submitKey == null)
            {
                submitKey = "{Enter}";
            }

            // Couple of sanity checks
            if (rowIndex < 0)
            {
                throw new Exception("RowIndex " + rowIndex.ToString() + " is not a valid row");
            }

            if (rowIndex < 0)
            {
                throw new Exception("ColumnIndex " + columnIndex.ToString() + " is not a valid column");
            }

            // Check if the cell is already set to the correct value
            string CurrentValue = this.GetCellValue(rowIndex, columnIndex, VSFlexgridCellPropertySettings.flexcpTextDisplay);
            if (CurrentValue == expectedValue)
            {
                GUI.Log("Ensure " + Identity.Description + " row " + rowText + " column " + columnText + " is set to " + expectedValue, LogItemType.Action);
                return false;
            }

            // Get the data type of the cell we want to set
            VSFlexgridColumnDataType cellDataType = GetColumnDataType(columnIndex);

            switch (cellDataType)
            {
                case VSFlexgridColumnDataType.flexDTEmpty:
                    // Click on the cell
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the cell in the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
                    this.SelectInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                case VSFlexgridColumnDataType.flexDTBoolean:
                    // Click on the checkbox
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
                    this.SelectInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                default:
                    throw new Exception("Not supported data type: " + cellDataType.ToString());
            }

            //Check the value was set
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                CurrentValue = this.GetCellValue(rowIndex, columnIndex, VSFlexgridCellPropertySettings.flexcpTextDisplay);

                if (CurrentValue == expectedValue)
                {
                    break;
                }

                // If the form it belongs to isn't enabled then there is likely a modal form displayed
                // so exit the loop so the code can continue (its up to the caller to validate the
                // value is set in these cases)
                if (!m_ParentForm.IsEnabled)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to set the flexgrid cell value");
                }

                Thread.Sleep(15);
            }
            while (true);

            return true;
        }

        /// <summary>
        /// Gets the indentation in pixels of a single tree level
        /// </summary>
        /// <returns>The indentation in pixels</returns>
        public int TreeViewIndent()
        {
            return 13;  //Seems to work
        }

        /// <summary>
        /// Gets the rows node outline level
        /// </summary>
        /// <param name="rowText">The row text</param>
        /// <returns>The node outline level</returns>
        public int NodeLevel(string rowText)
        {
            int RowNumber = FindRow(rowText);
            return NodeLevel(RowNumber);
        }

        /// <summary>
        /// Gets the rows node outline level
        /// </summary>
        /// <param name="rowIndex">The row index</param>
        /// <returns>The node outline level</returns>
        public int NodeLevel(int rowIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "RowOutlineLevel", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int NodeLevel = GUI.m_APE.GetValueFromMessage();

            return NodeLevel;
        }

        /// <summary>
        /// Gets the width of the border of the grid
        /// </summary>
        /// <returns>The width in pixels</returns>
        public int BorderWidth()
        {
            return 0;  //Seems to work
        }

        /// <summary>
        /// Gets the location given a row and column and position in a cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <param name="LocationInCell">The location in the cell</param>
        /// <returns>The points location</returns>
        private Point GetLocationInCell(int rowIndex, int columnIndex, CellClickLocation LocationInCell)
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
            switch (LocationInCell)
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
                case CellClickLocation.ExpandCollapseIconOfCell:
                    Location.X = CellRectangle.Left + 5 + 2 + BorderWidth() + (TreeViewIndent() * NodeLevel(rowIndex));
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                case CellClickLocation.LeftSideOfTreeItem:
                    Location.X = CellRectangle.Left + 5 + 2 + BorderWidth() + (TreeViewIndent() * (NodeLevel(rowIndex) + 1));
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                default:
                    throw new Exception("Implement for CellClickLocation: " + Location.ToString());
            }

            return Location;
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowText">The row text of the cell</param>
        /// <param name="columnText">The column text of the cell</param>
        /// <returns></returns>
        private Rectangle GetCellRectangle(string rowText, string columnText)
        {
            int RowNumber = FindRow(rowText);
            int ColumnNumber = FindColumn(columnText);
            return GetCellRectangle(RowNumber, ColumnNumber);
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnText">The column text of the cell</param>
        /// <returns></returns>
        private Rectangle GetCellRectangle(int rowIndex, string columnText)
        {
            int ColumnNumber = FindColumn(columnText);
            return GetCellRectangle(rowIndex, ColumnNumber);
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowText">The row text of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <returns></returns>
        private Rectangle GetCellRectangle(string rowText, int columnIndex)
        {
            int RowNumber = FindRow(rowText);
            return GetCellRectangle(RowNumber, columnIndex);
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <returns>The bound rectangle</returns>
        private Rectangle GetCellRectangle(int rowIndex, int columnIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ColPos", MemberTypes.Method, new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "RowPos", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store5, "ColWidth", MemberTypes.Method, new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store6, "RowHeight", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int twipsX = GUI.m_APE.GetValueFromMessage();
            int twipsY = GUI.m_APE.GetValueFromMessage();
            int twipsWidth = GUI.m_APE.GetValueFromMessage();
            int twipsHeight = GUI.m_APE.GetValueFromMessage();

            // Convert twips to pixels
            int x = TwipsToPixels(twipsX, Direction.Horizontal);
            int y = TwipsToPixels(twipsY, Direction.Vertical);
            int width = TwipsToPixels(twipsWidth, Direction.Horizontal);
            int height = TwipsToPixels(twipsHeight, Direction.Vertical);

            Rectangle CellRectangle = new Rectangle(x, y, width, height);
            return CellRectangle;
        }

        /// <summary>
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowText">Row index of the cell</param>
        /// <param name="columnText">Column text of the cell delimited by -> for example Order -> Id</param>
        public void Show(string rowText, string columnText)
        {
            int rowIndex = FindRow(rowText);
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
            if (!IsCellVisible(rowIndex, columnIndex))
            {
                //TODO move mouse over the grid if not in drag mode?
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ShowCell", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
            }
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void Select(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText);
            int columnIndex = FindColumn(columnText);

            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void Select(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText);
            int columnIndex = FindColumn(columnText);

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
        /// <param name="rowText">The row text of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void Select(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText);

            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void Select(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText);

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
            Point location = GetLocationInCell(row, column, locationInCell);

            // Find the actual native grid to click on
            GUIAxLZResultsGrid nativeGrid = new GUIAxLZResultsGrid(m_ParentForm, "native flexgrid", new Identifier(Identifiers.TypeName, "VSFlexGrid8N"), new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.ChildOf, this));

            nativeGrid.MouseSingleClickInternal(location.X, location.Y, button, keyModifier);
        }

        /// <summary>
        /// Move the mouse cursor over the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to move the mouse to</param>
        /// <param name="columnIndex">The column index of the cell to move the mouse to</param>
        /// <param name="locationInCell">The location in the cell to move the mouse to</param>
        public void MouseMove(int rowIndex, int columnIndex, CellClickLocation locationInCell)
        {
            Point location = GetLocationInCell(rowIndex, columnIndex, locationInCell);

            // Find the actual native grid to click on
            GUIAxLZResultsGrid nativeGrid = new GUIAxLZResultsGrid(m_ParentForm, "native flexgrid", new Identifier(Identifiers.TypeName, "VSFlexGrid8N"), new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.ChildOf, this));

            nativeGrid.MouseMove(location.X, location.Y);
        }






        public dynamic GetCellValue(string rowText, string columnText, VSFlexgridCellPropertySettings property)
        {
            int rowIndex = FindRow(rowText);
            int columnIndex = FindColumn(columnText);
            return GetCellValue(rowIndex, columnIndex, property);
        }

        public dynamic GetCellValue(int row, string columnText, VSFlexgridCellPropertySettings Property)
        {
            int columnIndex = FindColumn(columnText);
            return GetCellValue(row, columnIndex, Property);
        }

        public dynamic GetCellValue(string rowText, int column, VSFlexgridCellPropertySettings property)
        {
            int rowIndex = FindRow(rowText);
            return GetCellValue(rowIndex, column, property);
        }

        public dynamic GetCellValue(int rowIndex, int columnIndex, VSFlexgridCellPropertySettings property)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cell", MemberTypes.Method, new Parameter(GUI.m_APE, (int)property), new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex), new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic value = GUI.m_APE.GetValueFromMessage();
            return value;
        }

        public int ColumnWidth(int column)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ColWidth", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int widthTwips = GUI.m_APE.GetValueFromMessage();
            return widthTwips;
        }




        public bool IsTreeView()
        {
            return false;
        }

        private int FindNodeRow(string NodeText)
        {
            return -1;
        }




        public string GetCellRangeClip(int Row1, int Col1, int Row2, int Col2)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cell", MemberTypes.Method, new Parameter(GUI.m_APE, (int)VSFlexgridCellPropertySettings.flexcpText), new Parameter(GUI.m_APE, Row1), new Parameter(GUI.m_APE, Col1), new Parameter(GUI.m_APE, Row2), new Parameter(GUI.m_APE, Col2));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string CellRangeClip = GUI.m_APE.GetValueFromMessage();
            return CellRangeClip;
        }

        /// <summary>
        /// Returns the rows index of the specified value in the first visible column
        /// </summary>
        /// <param name="rowText">The value to look for in the first visible column</param>
        /// <returns>The index of the row</returns>
        public int FindRow(string rowText)
        {
            int columnIndex = FirstVisibleColumn();
            int startAtRow = 0; //FixedRows();
            return FindRow(rowText, columnIndex, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnText">The column to look for the value in delimited by -> for example Order -> Id</param>
        /// <returns>The index of the row</returns>
        public int FindRow(string rowText, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            int startAtRow = 0;//FixedRows();
            return FindRow(rowText, columnIndex, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnText">The column to look for the value in delimited by -> for example Order -> Id</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row</returns>
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
        /// <returns>The index of the row</returns>
        public int FindRow(string rowText, int columnIndex)
        {
            int startAtRow = 0;//FixedRows();
            return FindRow(rowText, columnIndex, startAtRow);
        }

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnIndex">The column to look for the value in</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row</returns>
        public int FindRow(string rowText, int columnIndex, int startAtRow)
        {
            int rowIndex;

            if (IsTreeView())
            {
                rowIndex = FindNodeRow(rowText);
            }
            else
            {
                rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
            }

            return rowIndex;
        }

        private int FindRowInternal(string Row, int Column, int StartAtRow)
        {
            int CurrentRow = StartAtRow - 1;

            do
            {
                CurrentRow++;
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "FindRow", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, CurrentRow), new Parameter(GUI.m_APE, Column), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, true));
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                CurrentRow = GUI.m_APE.GetValueFromMessage();

                if (CurrentRow == -1)
                {
                    return -1;
                }
            } while (IsRowHidden(CurrentRow));

            return CurrentRow;
        }

        public int FindColumn(string Column)
        {
            string[] Delimiter = { " -> " };
            string[] ColumnHeader = Column.Split(Delimiter, StringSplitOptions.None);

            return FindColumn(ColumnHeader);
        }

        public int FindColumn(string[] ColumnHeader)
        {
            int Column = -1;

            // Columns present may change so try twice
            try
            {
                Column = FindColumnInternal(ColumnHeader);
            }
            catch
            {
                Column = FindColumnInternal(ColumnHeader);
            }

            if (Column == -1)
            {
                throw new Exception("Failed to find column");
            }

            return Column;
        }

        private int FindColumnInternal(string[] ColumnHeader)
        {
            // Build a 2d array of the header
            int Columns = this.Columns();
            string[,] GridHeader = new string[ColumnHeader.Length, Columns];

            int VisibleRow = 0;
            int Row = 0;
            while (VisibleRow < ColumnHeader.Length)
            {
                if (!this.IsRowHidden(Row))
                {
                    string CurrentRow = GetCellRangeClip(Row, 0, Row, Columns - 1);

                    if (this.Columns() != Columns)
                    {
                        throw new Exception("Grid columns changed");
                    }

                    string[] Delimiter = { "\t" };
                    string[] CurrentRowArray = CurrentRow.Split(Delimiter, StringSplitOptions.None);

                    for (int Column = 0; Column < Columns; Column++)
                    {
                        GridHeader[VisibleRow, Column] = CurrentRowArray[Column];
                    }

                    VisibleRow++;
                }
                Row++;
            }

            //Search for the column
            bool Found = false;
            for (int Column = 0; Column < Columns; Column++)
            {
                if (!this.IsColumnHidden(Column))
                {
                    for (Row = 0; Row < ColumnHeader.Length; Row++)
                    {
                        if (GridHeader[Row, Column] == ColumnHeader[Row])
                        {
                            Found = true;
                        }
                        else
                        {
                            Found = false;
                            break;
                        }
                    }
                }
                if (Found)
                {
                    return Column;
                }
            }

            return -1;
        }

        private enum Direction
        {
            Vertical,
            Horizontal,
        }

        private int TwipsToPixels(int twips, Direction direction)
        {
            int pixelsPerInch;
            IntPtr deviceContext = NM.GetDC(this.Handle);
            
            switch (direction)
            {
                case Direction.Horizontal:
                    pixelsPerInch = NM.GetDeviceCaps(deviceContext, NM.DeviceCap.LOGPIXELSX);
                    break;
                case Direction.Vertical:
                    pixelsPerInch = NM.GetDeviceCaps(deviceContext, NM.DeviceCap.LOGPIXELSY);
                    break;
                default:
                    throw new Exception("Unknown direction: " + direction.ToString());
            }
            NM.ReleaseDC(this.Handle, deviceContext);
            int pixels = (int)Math.Round((float)twips / (float)1440 * (float)pixelsPerInch);
            return pixels;
        }

        private VSFlexgridColumnDataType GetColumnDataType(int column)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ColDataType", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int dataType = GUI.m_APE.GetValueFromMessage();
            return (VSFlexgridColumnDataType)dataType;
        }
    }
}
