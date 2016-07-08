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
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using APE.Capture;
using APE.Communication;
using System.Threading;
using System.Drawing.Imaging;
using System.Security.Principal;
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
            flexcpText = 0,
            flexcpTextStyle = 1,
            flexcpAlignment = 2,
            flexcpPicture = 3,
            flexcpPictureAlignment = 4,
            flexcpChecked = 5,
            flexcpBackColor = 6,
            flexcpForeColor = 7,
            flexcpFloodPercent = 8,
            flexcpFloodColor = 9,
            flexcpFont = 10,
            flexcpFontName = 11,
            flexcpFontSize = 12,
            flexcpFontBold = 13,
            flexcpFontItalic = 14,
            flexcpFontUnderline = 15,
            flexcpFontStrikethru = 16,
            flexcpFontWidth = 17,
            flexcpValue = 18,
            flexcpTextDisplay = 19,
            flexcpData = 20,
            flexcpCustomFormat = 21,
            flexcpLeft = 22,
            flexcpTop = 23,
            flexcpWidth = 24,
            flexcpHeight = 25,
            flexcpVariantValue = 26,
            flexcpRefresh = 27,
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "CopyToClipboard", MemberTypes.Method);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "RowHidden", MemberTypes.Property, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ColHidden", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "LeftCol", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "RightCol", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store5, "TopRow", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store6, "BottomRow", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddMessageGetValue(DataStores.Store5);
            GUI.m_APE.AddMessageGetValue(DataStores.Store6);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "RowSel", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ColSel", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "FixedRows", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "FixedCols", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "FrozenCols", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Rows", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Cols", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int columns = GUI.m_APE.GetValueFromMessage();
            return columns;
        }











        public string GetCellValue(int row, int column, VSFlexgridCellPropertySettings property)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Cell", MemberTypes.Method, new Parameter(GUI.m_APE, (int)property), new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column), new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string value = GUI.m_APE.GetValueFromMessage();
            return value;
        }


        public int ColumnWidth(int column)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ColWidth", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
            GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Cell", MemberTypes.Method, new Parameter(GUI.m_APE, (int)VSFlexgridCellPropertySettings.flexcpText), new Parameter(GUI.m_APE, Row1), new Parameter(GUI.m_APE, Col1), new Parameter(GUI.m_APE, Row2), new Parameter(GUI.m_APE, Col2));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string CellRangeClip = GUI.m_APE.GetValueFromMessage();
            return CellRangeClip;
        }

        public int FindRow(string Row)
        {
            int ColumnNumber = FirstVisibleColumn();
            int StartAtRow = 0; //FixedRows();
            return FindRow(Row, ColumnNumber, StartAtRow);
        }

        public int FindRow(string Row, string Column)
        {
            int ColumnNumber = FindColumn(Column);
            int StartAtRow = 0; //FixedRows();
            return FindRow(Row, ColumnNumber, StartAtRow);
        }

        public int FindRow(string Row, string Column, int StartAtRow)
        {
            int ColumnNumber = FindColumn(Column);
            return FindRow(Row, ColumnNumber, StartAtRow);
        }

        public int FindRow(string Row, int Column, int StartAtRow)
        {
            int CurrentRow = StartAtRow;

            if (IsTreeView())
            {
                CurrentRow = FindNodeRow(Row);
            }
            else
            {
                CurrentRow = FindRowInternal(Row, Column, StartAtRow);
            }

            return CurrentRow;
        }

        private int FindRowInternal(string Row, int Column, int StartAtRow)
        {
            int CurrentRow = StartAtRow - 1;

            do
            {
                CurrentRow++;
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
                GUI.m_APE.AddMessageGetUnderlyingGridFromResultsGrid(DataStores.Store1, DataStores.Store2);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "FindRow", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, CurrentRow), new Parameter(GUI.m_APE, Column), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, true));
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
    }
}
