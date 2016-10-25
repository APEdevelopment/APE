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
using System.Collections.Generic;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Capstone.ComSupport.ResultsGrid.AxLZResultsGrid
    /// AxDRILLDOWNLib.AxLZResultsGrid
    /// </summary>
    public sealed class GUIAxLZResultsGrid : GUIGridObject
    {
        private enum VSFlexgridCellPropertySettings
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "RowHidden", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
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
        /// <param name="columnText">Column to check if hidden delimited by -> for example Order -> Id</param>
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
        public bool IsColumnHidden(int columnIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ColHidden", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndex));
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
        /// <param name="rowIndex">The row index of the cell to check</param>
        /// <param name="columnIndex">The column index of the cell to check</param>
        /// <returns>True if the cell is visible without scrolling otherwise false</returns>
        public bool IsCellVisible(int rowIndex, int columnIndex)
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

            if (columnIndex <= leftColumn || columnIndex >= rightColumn || rowIndex <= topRow || rowIndex >= bottomRow)
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

            switch (compareMethod)
            {
                case ComparisonMethod.CompareUsingDefaultEqualityComparer:
                    // Check if the cell is already set to the correct value
                    string currentValue = this.GetCell(rowIndex, columnIndex, CellProperty.TextDisplay);
                    T currentValueT = (T)Convert.ChangeType(currentValue, typeof(T));
                    if (EqualityComparer<T>.Default.Equals(currentValueT, expectedValue))
                    {
                        GUI.Log("Ensure " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText + " is set to " + expectedValue, LogItemType.Action);
                        return false;
                    }
                    break;
                case ComparisonMethod.DoNotCompare:
                    break;
                default:
                    throw new Exception("Unsupported compare method: " + compareMethod.ToString());
            }

            // Get the data type of the cell we want to set
            VSFlexgridColumnDataType cellDataType = GetColumnDataType(columnIndex);

            switch (cellDataType)
            {
                case VSFlexgridColumnDataType.flexDTEmpty:
                    // Click on the cell
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the cell in the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                    this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                case VSFlexgridColumnDataType.flexDTBoolean:
                    // Click on the checkbox
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                    this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                default:
                    throw new Exception("Not supported data type: " + cellDataType.ToString());
            }

            switch (compareMethod)
            {
                case ComparisonMethod.CompareUsingDefaultEqualityComparer:
                    //Check the value was set
                    timer = Stopwatch.StartNew();
                    while (true)
                    {
                        string currentValue = GetCell(rowIndex, columnIndex, CellProperty.TextDisplay);

                        T currentValueT = (T)Convert.ChangeType(currentValue, typeof(T));
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
                            throw new Exception("Failed to set the flexgrid cell value");
                        }

                        Thread.Sleep(15);
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
            int rowIndex = FindRow(rowText);
            return NodeLevel(rowIndex);
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
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <returns>The bound rectangle</returns>
        internal override Rectangle GetCellRectangleInternal(int rowIndex, int columnIndex)
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
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        internal override void ShowCellInternal(int rowIndex, int columnIndex)
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
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        internal override void SingleClickCellInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            // Find the actual native grid to click on
            GUIAxLZResultsGrid nativeGrid = new GUIAxLZResultsGrid(ParentForm, "native flexgrid", new Identifier(Identifiers.TypeName, "VSFlexGrid8N"), new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.ChildOf, this));
            nativeGrid.SingleClickInternal(location.X, location.Y, button, keyModifier);
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
            // Find the actual native grid to click on
            GUIAxLZResultsGrid nativeGrid = new GUIAxLZResultsGrid(ParentForm, "native flexgrid", new Identifier(Identifiers.TypeName, "VSFlexGrid8N"), new Identifier(Identifiers.TechnologyType, "Windows Native"), new Identifier(Identifiers.ChildOf, this));
            nativeGrid.DoubleClickInternal(Location.X, Location.Y, button, keyModifier);
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
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        internal override dynamic GetCellInternal(int rowIndex, int columnIndex, CellProperty property)
        {
            int comProperty;

            switch (property)
            {
                case CellProperty.TextDisplay:
                    comProperty = (int)VSFlexgridCellPropertySettings.flexcpTextDisplay;
                    break;
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cell", MemberTypes.Method, new Parameter(GUI.m_APE, comProperty), new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex), new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
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



        // TODO treeview stuff
        public bool IsTreeView()
        {
            return false;
        }

        public int TreeViewColumn()
        {
            return -1;
        }

        private int FindNodeRow(string NodeText)
        {
            return -1;
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
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
                    GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Cell", MemberTypes.Method, new Parameter(GUI.m_APE, (int)VSFlexgridCellPropertySettings.flexcpText), new Parameter(GUI.m_APE, row1Index), new Parameter(GUI.m_APE, column1Index), new Parameter(GUI.m_APE, row2Index), new Parameter(GUI.m_APE, column2Index));
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeClip = GUI.m_APE.GetValueFromMessage();
                    return rangeClip;
 
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

        /// <summary>
        /// Returns the rows index of the specified value in the specified column
        /// </summary>
        /// <param name="rowText">The value to look for in the specified column</param>
        /// <param name="columnIndex">The column to look for the value in</param>
        /// <param name="startAtRow">The row to start the search at</param>
        /// <returns>The index of the row or -1</returns>
        internal override int FindRowInternal(string rowText, int columnIndex, int startAtRow)
        {
            int CurrentRow = startAtRow - 1;

            do
            {
                CurrentRow++;
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageSentinelGridsGetUnderlyingGrid(DataStores.Store1, DataStores.Store2);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "FindRow", MemberTypes.Method, new Parameter(GUI.m_APE, rowText), new Parameter(GUI.m_APE, CurrentRow), new Parameter(GUI.m_APE, columnIndex), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, true));
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

        /// <summary>
        /// Returns the index of the specified column in the grid
        /// </summary>
        /// <param name="columnHeader">The column to check</param>
        /// <returns>The index of the column</returns>
        internal override int FindColumnInternal(string[] columnHeader)
        {
            // Build a 2d array of the header
            int Columns = this.Columns();
            string[,] GridHeader = new string[columnHeader.Length, Columns];

            int VisibleRow = 0;
            int Row = 0;
            while (VisibleRow < columnHeader.Length)
            {
                if (!this.IsRowHidden(Row))
                {
                    string CurrentRow = GetCellRange(Row, 0, Row, Columns - 1);

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
                    for (Row = 0; Row < columnHeader.Length; Row++)
                    {
                        if (GridHeader[Row, Column] == columnHeader[Row])
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
