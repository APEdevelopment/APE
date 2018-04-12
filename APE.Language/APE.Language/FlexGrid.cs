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

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// C1.Win.C1FlexGrid.C1FlexGrid
    /// </summary>
    public sealed class GUIFlexgrid : GUIGridObject
    {
        /// <summary>
        /// Initialises a new instance of the GUIFlexgrid class
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIFlexgrid(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        private IntPtr EditWindow()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Editor", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic ReturnValue = GUI.m_APE.GetValueFromMessage();     //Could be a null so we use dynamic

            if (ReturnValue == null)
            {
                return IntPtr.Zero;
            }
            else
            {
                return ReturnValue;
            }
        }

        private bool IsInEditMode()
        {
            IntPtr EditWindowHandle = EditWindow();

            if (EditWindowHandle == IntPtr.Zero)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private string GetEditText()
        {
            IntPtr EditWindowHandle = EditWindow();

            if (EditWindowHandle == IntPtr.Zero)
            {
                throw new Exception(Description + " is not in edit mode");
            }

            return GUI.m_APE.GetWindowTextViaWindowMessage(EditWindowHandle);
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

        private bool IsNode(int rowIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "IsNode", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool RowIsNode = GUI.m_APE.GetValueFromMessage();

            return RowIsNode;
        }

        private bool IsNodeCollapsed(int rowIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Node", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Collapsed", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool collapsed = GUI.m_APE.GetValueFromMessage();

            return collapsed;
        }

        private bool HasNodeGotChildren(int rowIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Node", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Children", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int children = GUI.m_APE.GetValueFromMessage();

            if (children > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /// <summary>
        /// Expand all nodes in the flexgrid treeview column
        /// </summary>
        public void ExpandTreeView()
        {
            GUI.Log("Fully expand the treeview", LogItemType.Action);

            int treeColumnIndex = TreeViewColumn();

            if (treeColumnIndex == -1)
            {
                throw new Exception(Description + " is not a treeview");
            }

            for (int rowIndex = 0; rowIndex < Rows(); rowIndex++)
            {
                if (IsNode(rowIndex))
                {
                    if (HasNodeGotChildren(rowIndex))
                    {
                        if (IsNodeCollapsed(rowIndex))
                        {
                            SingleClickCellInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                            Stopwatch timer = Stopwatch.StartNew();
                            while (true)
                            {
                                if (!IsNodeCollapsed(rowIndex))
                                {
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw new Exception("Failed to expand the " + Description + " node");
                                }

                                Thread.Sleep(15);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Collapse all nodes in the flexgrid treeview column
        /// If you need a maximally collapsed grid then call ExpandTreeView before calling CollapseTreeView
        /// </summary>
        public void CollapseTreeView()
        {
            GUI.Log("Fully collapse the treeview", LogItemType.Action);

            int treeColumnIndex = TreeViewColumn();

            if (treeColumnIndex == -1)
            {
                throw new Exception(Description + " is not a treeview");
            }

            for (int rowIndex = Rows() - 1; rowIndex > -1; rowIndex--)
            {
                if (IsNode(rowIndex))
                {
                    if (HasNodeGotChildren(rowIndex))
                    {
                        if (!IsNodeCollapsed(rowIndex))
                        {
                            SingleClickCellInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                            Stopwatch timer = Stopwatch.StartNew();
                            while (true)
                            {
                                if (IsNodeCollapsed(rowIndex))
                                {
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw new Exception("Failed to collapse the " + Description + " node");
                                }

                                Thread.Sleep(15);
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Expands the specified nodes in the flexgrid treeview column
        /// </summary>
        /// <param name="nodePath">Node path to expand delimited by -> (or the user defined GridDelimiter property) for example RULE LIBRARY -> UCITS I -> France</param>
        public void ExpandNodes(string nodePath)
        {
            GUI.Log("Expand node " + nodePath, LogItemType.Action);

            string[] nodePathArray = nodePath.Split(GUI.GridDelimiterAsArray, StringSplitOptions.None);

            int treeColumnIndex = TreeViewColumn();
            string row = "";

            if (treeColumnIndex == -1)
            {
                throw new Exception(Description + " is not a treeview");
            }

            for (int i = 0; i < nodePathArray.GetLength(0); i++)
            {
                if (i == 0)
                {
                    row = nodePathArray[0];
                }
                else
                {
                    row += GUI.GridDelimiter + nodePathArray[i];
                }
                int rowIndex = FindNodeRow(row);

                if (!HasNodeGotChildren(rowIndex))
                {
                    throw new Exception("Can not expand the node of the " + Description + " is has no children");
                }

                if (IsNodeCollapsed(rowIndex))
                {
                    SingleClickCellInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (!IsNodeCollapsed(rowIndex))
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to expand the " + Description + " node");
                        }

                        Thread.Sleep(15);
                    }
                }
            }
        }

        /// <summary>
        /// Collapses the specified nodes in the flexgrid treeview column
        /// </summary>
        /// <param name="nodePath">Node path to collapse delimited by -> (or the user defined GridDelimiter property) for example RULE LIBRARY -> UCITS I -> France</param>
        public void CollapseNodes(string nodePath)
        {
            GUI.Log("Collapse node " + nodePath, LogItemType.Action);

            string[] nodePathArray = nodePath.Split(GUI.GridDelimiterAsArray, StringSplitOptions.None);

            int treeColumnIndex = TreeViewColumn();
            string row = "";

            if (treeColumnIndex == -1)
            {
                throw new Exception(Description + " is not a treeview");
            }

            for (int i = nodePathArray.GetLength(0) - 1; i > -1; i--)
            {
                for (int x = 0; x <= i; x++)
                {
                    if (x == 0)
                    {
                        row = nodePathArray[0];
                    }
                    else
                    {
                        row += GUI.GridDelimiter + nodePathArray[x];
                    }
                }

                int rowIndex = FindNodeRow(row);

                if (!HasNodeGotChildren(rowIndex))
                {
                    throw new Exception("Can not collapse the node of the " + Description + " is has no children");
                }

                if (!IsNodeCollapsed(rowIndex))
                {
                    SingleClickCellInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (IsNodeCollapsed(rowIndex))
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to collapse the " +  Description + " node");
                        }

                        Thread.Sleep(15);
                    }
                }
            }
        }

        private int GetNodeParentRow(int Row)
        {
            int RowNodeLevel = NodeLevel(Row);
            int CurrentRow = Row;
            int CurrentLevel = -1;

            if (RowNodeLevel > 0)
            {
                do
                {
                    CurrentRow--;
                    if (IsNode(CurrentRow))
                    {
                        CurrentLevel = NodeLevel(CurrentRow);

                        if (CurrentLevel == RowNodeLevel)
                        {
                            //Sibbling
                        }
                        else if (CurrentLevel == RowNodeLevel - 1)
                        {
                            //Parent
                            return CurrentRow;
                        }
                        else
                        {
                            //probably sibbling child
                        }
                    }
                } while (true);
            }
            else
            {
                return -1;
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
            // TODO include where height = 0 or 1?
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, row));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool Visible = GUI.m_APE.GetValueFromMessage();

            return !Visible;
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
        /// <param name="column">Column index to check if hidden</param>
        public bool IsColumnHidden(int column)
        {
            // TODO include where width = 0 or 1?
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, column));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool Visible = GUI.m_APE.GetValueFromMessage();

            return !Visible;
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "LeftCol", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "RightCol", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "TopRow", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store4, "BottomRow", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Row", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedRow = GUI.m_APE.GetValueFromMessage();

            return SelectedRow;
        }

        /// <summary>
        /// Polls for the specified row index to be the selected row
        /// </summary>
        /// <param name="rowIndex">The row index to wait to be selected</param>
        public void SelectedRowPollForIndex(int rowIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessagePollMember(DataStores.Store0, "Row", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, false));
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
        }

        /// <summary>
        /// Returns the selected column
        /// </summary>
        /// <returns>The selected column index</returns>
        public int SelectedColumn()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Col", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedColumn = GUI.m_APE.GetValueFromMessage();

            return SelectedColumn;
        }

        /// <summary>
        /// Returns the number of fixed rows, that is a row which makes up the header of the grid (rows may or may not be hidden)
        /// </summary>
        /// <returns>The number of fixed rows</returns>
        public int FixedRows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Fixed", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedRows = GUI.m_APE.GetValueFromMessage();

            return FixedRows;
        }

        /// <summary>
        /// Returns the number of frozen rows, that is a row which doesn't scroll, in the grid (rows may or may not be hidden)
        /// </summary>
        /// <returns>The number of frozen rows</returns>
        public int FrozenRows()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Frozen", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int frozenRows = GUI.m_APE.GetValueFromMessage();
            return frozenRows;
        }

        /// <summary>
        /// Returns the number of fixed columns, that is a column which doesn't scroll, in the grid (columns may or may not be hidden)
        /// </summary>
        /// <returns>The number of fixed columns</returns>
        public int FixedColumns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Fixed", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int fixedColumns = GUI.m_APE.GetValueFromMessage();
            return fixedColumns;
        }

        /// <summary>
        /// Returns the number of frozen columns, that is a column which doesn't scroll, in the grid (columns may or may not be hidden)
        /// </summary>
        /// <returns>The number of frozen columns</returns>
        public int FrozenColumns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Frozen", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FrozenColumns = GUI.m_APE.GetValueFromMessage();

            return FrozenColumns;
        }

        /// <summary>
        /// Returns the number of title (column header) rows in the grid (the rows may or may not be visible)
        /// </summary>
        /// <returns>The number of title rows</returns>
        public int TitleRows()
        {
            // bit of a hack as the flexgrid doesn't tell you
            int fixedRows = FixedRows();
            if (fixedRows == 0)
            {
                return 0;
            }

            int columnCount = Columns();
            if (columnCount == 0)
            {
                return 0;
            }

            // Get the first visible column which isn't selected (selected columns may change the colour)
            int visibleNonSelectedColumn = -1;
            
            int selectedColumn = SelectedColumn();
            for (int column = 0; column < columnCount; column++)
            {
                if (!IsColumnHidden(column))
                {
                    if (column != selectedColumn)
                    {
                        visibleNonSelectedColumn = column;
                        break;
                    }
                }
            }

            if (visibleNonSelectedColumn == -1)
            {
                visibleNonSelectedColumn = FirstVisibleColumn();
                if (visibleNonSelectedColumn == -1)
                {
                    visibleNonSelectedColumn = 0;
                }
            }

            // get the colour of each cell background of the first visible nonselected column for the fixed rows and return the number of rows which match the first row
            string backColourName = GetCellRange(0, visibleNonSelectedColumn, fixedRows - 1, visibleNonSelectedColumn, CellProperty.BackColourName);

            char[] splitSeparator = { '\r' };
            string[] backColourNameArray = backColourName.Split(splitSeparator);

            string headerBackColour = backColourNameArray[0];

            int titleRows = 1;
            for (int headerRow = 1; headerRow < backColourNameArray.GetLength(0); headerRow++)
            {
                if (backColourNameArray[headerRow] == headerBackColour)
                {
                    titleRows++;
                }
                else
                {
                    break;
                }
            }

            return titleRows;
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
            int Rows = GUI.m_APE.GetValueFromMessage();

            return Rows;
        }

        /// <summary>
        /// Returns the number of columns, including those which are hidden
        /// </summary>
        /// <returns>The number of columns</returns>
        public int Columns()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Columns = GUI.m_APE.GetValueFromMessage();

            return Columns;
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
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "AllowEditing", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                bool editable = GUI.m_APE.GetValueFromMessage();
                return editable;
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "AllowEditing", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool rowAllowEditing = GUI.m_APE.GetValueFromMessage();
            if (!rowAllowEditing)
            {
                return false;
            }

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, columnIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "AllowEditing", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool columnAllowEditing = GUI.m_APE.GetValueFromMessage();
            if (!columnAllowEditing)
            {
                return false;
            }

            return true;
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

            CellProperty propertyToCheck;

            // Get the data type of the cell we want to set
            string cellDataType = this.GetCell(rowIndex, columnIndex, CellProperty.DataType);
            string cellUserDataType = this.GetCell(rowIndex, columnIndex, CellProperty.UserDataType);

            //Check if its checkbox
            string cellCheckBox = this.GetCell(rowIndex, columnIndex, CellProperty.CheckBox);
            if (cellCheckBox == "None")
            {
                propertyToCheck = CellProperty.TextDisplay;
                if (cellUserDataType == "LzFGCheckBoxSupport.CellCheckedStatus")
                {
                    cellDataType = "System.Boolean";
                }
            }
            else
            {
                cellDataType = "System.Boolean";
                propertyToCheck = CellProperty.CheckBox;
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

            switch (cellDataType)
            {
                //case "System.DateTime":
                case "System.Boolean":  //checkbox
                    // Click on the checkbox
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + Identity.Description + " row " + rowFriendlyText + " column " + columnFriendlyText, LogItemType.Action);
                    int treeColumn = TreeViewColumn();
                    if (treeColumn == columnIndex)
                    {
                        this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.LeftSideOfTreeItem, MouseKeyModifier.None);
                    }
                    else
                    {
                        // May need to determine location to click at some point
                        this.SingleClickCellInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    }
                    break;
                default:
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

                    // Build a list of genericwalkes which currently exist
                    List<IntPtr> walkers = new List<IntPtr>(200);
                    GUIGenericWalker genericWalker = null;
                    for (int index = 1; index < 200; index++)
                    {
                        bool found = false;
                        if (GUI.Exists(ParentForm, new Identifier(Identifiers.Name, "lzGenericWalkerCtl"), new Identifier(Identifiers.Index, index)))
                        {
                            found = true;
                            genericWalker = new GUIGenericWalker(ParentForm, "walker", new Identifier(Identifiers.Name, "lzGenericWalkerCtl"), new Identifier(Identifiers.Index, index));
                            walkers.Add(genericWalker.Handle);
                        }

                        if (GUI.Exists(ParentForm, new Identifier(Identifiers.Name, "LzGenericWalkerCtl"), new Identifier(Identifiers.Index, index)))
                        {
                            found = true;
                            genericWalker = new GUIGenericWalker(ParentForm, "walker", new Identifier(Identifiers.Name, "LzGenericWalkerCtl"), new Identifier(Identifiers.Index, index));
                            walkers.Add(genericWalker.Handle);
                        }

                        if (!found)
                        {
                            genericWalker = null;
                            break;
                        }
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemType.Action);
                    base.SendKeysInternal("{F2}");

                    string APEDirectType;
                    string APEBaseType;
                    dynamic EditorHandle;

                    // Search for the editor
                    timer = Stopwatch.StartNew();
                    while (true)
                    {
                        GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Editor", MemberTypes.Property);
                        GUI.m_APE.AddMessageGetRecognisedType(DataStores.Store1, DataStores.Store2);
                        GUI.m_APE.AddMessageGetApeTypeFromObject(DataStores.Store1, DataStores.Store3); //Get this to help with debuging
                        GUI.m_APE.AddMessageGetApeTypeFromType(DataStores.Store2, DataStores.Store4);
                        GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store5, "Handle", MemberTypes.Property);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                        GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call;
                        APEDirectType = GUI.m_APE.GetValueFromMessage();
                        APEBaseType = GUI.m_APE.GetValueFromMessage();
                        EditorHandle = GUI.m_APE.GetValueFromMessage();

                        // If the editor isn't visible then its likely not being used, so search for the real editor
                        if (EditorHandle == null || !NM.IsWindowVisible(EditorHandle))
                        {
                            // Search for a generic walker which is new
                            for (int index = 1; index < 200; index++)
                            {
                                bool found = false;
                                if (GUI.Exists(ParentForm, new Identifier(Identifiers.Name, "lzGenericWalkerCtl"), new Identifier(Identifiers.Index, index)))
                                {
                                    found = true;
                                    genericWalker = new GUIGenericWalker(ParentForm, "walker", new Identifier(Identifiers.Name, "lzGenericWalkerCtl"), new Identifier(Identifiers.Index, index));
                                    if (!walkers.Contains(genericWalker.Handle))
                                    {
                                        EditorHandle = IntPtr.Zero;
                                        break;
                                    }
                                }

                                if (GUI.Exists(ParentForm, new Identifier(Identifiers.Name, "LzGenericWalkerCtl"), new Identifier(Identifiers.Index, index)))
                                {
                                    found = true;
                                    genericWalker = new GUIGenericWalker(ParentForm, "walker", new Identifier(Identifiers.Name, "LzGenericWalkerCtl"), new Identifier(Identifiers.Index, index));
                                    if (!walkers.Contains(genericWalker.Handle))
                                    {
                                        EditorHandle = IntPtr.Zero;
                                        break;
                                    }
                                }

                                if (!found)
                                {
                                    genericWalker = null;
                                    break;
                                }
                            }

                            if (genericWalker != null)
                            {
                                break;
                            }
                        }
                        else
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to find the " + Description + " cell editor");
                        }

                        Thread.Sleep(50);
                    }

                    // Add event handlers
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageAddFlexgridCellChangedHandler(DataStores.Store0);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);

                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageAddFlexgridAfterRowColChangeHandler(DataStores.Store0);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);

                    // Change the cell value
                    string valueText = value.ToString();
                    if (genericWalker != null)
                    {
                        genericWalker.SetText(valueText);
                    }
                    else
                    {
                        // Set the value
                        switch (APEBaseType)
                        {
                            case "GUIComboBox":
                                GUIComboBox flexgridComboBox = new GUIComboBox(ParentForm, Identity.Description + " combobox", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridComboBox.SingleClickItem(valueText);
                                break;
                            case "GUITextBox":
                                GUITextBox flexgridTextBox = new GUITextBox(ParentForm, Identity.Description + " textbox", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridTextBox.SetText(valueText);
                                GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                                base.SendKeysInternal(submitKey);
                                break;
                            case "GUIDateTimePicker":
                                GUIDateTimePicker flexgridDateTimePicker = new GUIDateTimePicker(ParentForm, Identity.Description + " datetime picker", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridDateTimePicker.SetText(valueText);
                                GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                                base.SendKeysInternal(submitKey);
                                break;
                            default:
                                throw new Exception("Unsupported flexgrid editor: Type: " + APEDirectType + " Base Type: " + APEBaseType);
                        }
                    }

                    //Wait for the event handlers then remove them
                    GUI.m_APE.AddFirstMessageWaitForAndRemoveFlexgridCellChangedHandler();
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);

                    if (this.SelectedColumn() == columnIndex && this.SelectedRow() == rowIndex)
                    {
                        GUI.m_APE.AddFirstMessageRemoveFlexgridAfterRowColChangeHandler();
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                    }
                    else
                    {
                        GUI.m_APE.AddFirstMessageWaitForAndRemoveFlexgridAfterRowColChangeHandler();
                        GUI.m_APE.SendMessages(EventSet.APE);
                        GUI.m_APE.WaitForMessages(EventSet.APE);
                    }

                    break;
            }

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
                        currentValueAsString = this.GetCell(rowIndex, columnIndex, propertyToCheck);
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

        //public string GetColumnType(int col)
        //{
        //    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property); ;
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, col));
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "DataType", MemberTypes.Property);
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Namespace", MemberTypes.Property);
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Name", MemberTypes.Property);
        //    GUI.m_APE.AddMessageGetValue(DataStores.Store4);
        //    GUI.m_APE.AddMessageGetValue(DataStores.Store5);
        //    GUI.m_APE.SendMessages(EventSet.APE);
        //    GUI.m_APE.WaitForMessages(EventSet.APE);
        //    //Get the value(s) returned MUST be done straight after the WaitForMessages call
        //    string columnTypeNamespace = GUI.m_APE.GetValueFromMessage();
        //    string columnTypeName = GUI.m_APE.GetValueFromMessage();
        //    return columnTypeNamespace + "." + columnTypeName;
        //}

        //public string GetColumnFormat(int col)
        //{
        //    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property); ;
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, col));
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Format", MemberTypes.Property);
        //    GUI.m_APE.AddMessageGetValue(DataStores.Store3);
        //    GUI.m_APE.SendMessages(EventSet.APE);
        //    GUI.m_APE.WaitForMessages(EventSet.APE);
        //    //Get the value(s) returned MUST be done straight after the WaitForMessages call
        //    string format = GUI.m_APE.GetValueFromMessage();
        //    return format;
        //}

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        internal override dynamic GetCellInternal(int rowIndex, int columnIndex, CellProperty property)
        {
            switch (property)
            {
                case CellProperty.TextDisplay:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "DataDisplay", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string CellDataDisplay = GUI.m_APE.GetValueFromMessage();
                    return CellDataDisplay;
                case CellProperty.BackColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "BackColor", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Name", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string backColourName = GUI.m_APE.GetValueFromMessage();
                    return backColourName;
                case CellProperty.ForeColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ForeColor", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Name", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string foreColourName = GUI.m_APE.GetValueFromMessage();
                    return foreColourName;
                case CellProperty.DataType:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "DataType", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Namespace", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Name", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string cellTypeNamespace = GUI.m_APE.GetValueFromMessage();
                    string cellTypeName = GUI.m_APE.GetValueFromMessage();
                    return cellTypeNamespace + "." + cellTypeName;
                case CellProperty.CheckBox:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Checkbox", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ToString", MemberTypes.Method);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string checkboxState = GUI.m_APE.GetValueFromMessage();
                    return checkboxState;
                case CellProperty.Image:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Image", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Image image = GUI.m_APE.GetValueFromMessage();
                    return image;
                case CellProperty.BackgroundImage:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "BackgroundImage", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Image backgroundImage = GUI.m_APE.GetValueFromMessage();
                    return backgroundImage;
                case CellProperty.FontStyle:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
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
                case CellProperty.UserDataType:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "UserData", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "GetType", MemberTypes.Method);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Namespace", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Name", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string cellUserDataTypeNamespace = GUI.m_APE.GetValueFromMessage();
                    string cellUserDataTypeName = GUI.m_APE.GetValueFromMessage();
                    if (cellUserDataTypeName == null || cellUserDataTypeName == null)
                    {
                        return null;
                    }
                    return cellUserDataTypeNamespace + "." + cellUserDataTypeName;
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
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
            switch (property)
            {
                case CellProperty.TextDisplay:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row1Index), new Parameter(GUI.m_APE, column1Index), new Parameter(GUI.m_APE, row2Index), new Parameter(GUI.m_APE, column2Index));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Clip", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeClip = GUI.m_APE.GetValueFromMessage();
                    return rangeClip;
                case CellProperty.BackColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1Index, column1Index, row2Index, column2Index, APEIPC.CellProperty.BackColourName);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeBackColourName = GUI.m_APE.GetValueFromMessage();
                    return rangeBackColourName;
                case CellProperty.ForeColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1Index, column1Index, row2Index, column2Index, APEIPC.CellProperty.ForeColourName);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeForeColourName = GUI.m_APE.GetValueFromMessage();
                    return rangeForeColourName;
                case CellProperty.DataType:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1Index, column1Index, row2Index, column2Index, APEIPC.CellProperty.DataType);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeDataType = GUI.m_APE.GetValueFromMessage();
                    return rangeDataType;
                case CellProperty.CheckBox:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1Index, column1Index, row2Index, column2Index, APEIPC.CellProperty.CheckBox);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeCheckBoxState = GUI.m_APE.GetValueFromMessage();
                    return rangeCheckBoxState;
                case CellProperty.Image:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1Index, column1Index, row2Index, column2Index, APEIPC.CellProperty.Image);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeHasImage = GUI.m_APE.GetValueFromMessage();
                    return rangeHasImage;
                case CellProperty.BackgroundImage:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1Index, column1Index, row2Index, column2Index, APEIPC.CellProperty.BackgroundImage);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeHasBackgroundImage = GUI.m_APE.GetValueFromMessage();
                    return rangeHasBackgroundImage;
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }
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
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ShowCell", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
                // TODO doing it twice is a workaround for a current bug which needs fixing
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ShowCell", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
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



            //bool removeAfterHandler = false;
            //bool removeBeforeHandler = false;
            //if (SelectedColumn() == columnIndex && SelectedRow() == rowIndex)
            //{
            //    //Cell is selected so dont add a handler
            //}
            //else
            //{
            //    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            //    GUI.m_APE.AddQueryMessageAddFlexgridBeforeSelChangeHandler(DataStores.Store0);
            //    GUI.m_APE.SendMessages(EventSet.APE);
            //    GUI.m_APE.WaitForMessages(EventSet.APE);
            //    removeBeforeHandler = true;

            //    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            //    GUI.m_APE.AddQueryMessageAddFlexgridAfterSelChangeHandler(DataStores.Store0);
            //    GUI.m_APE.SendMessages(EventSet.APE);
            //    GUI.m_APE.WaitForMessages(EventSet.APE);
            //    removeAfterHandler = true;
            //}

            SingleClickInternal(Location.X, Location.Y, button, keyModifier);

            //if (removeBeforeHandler)
            //{
            //    GUI.m_APE.AddFirstyMessageWaitForAndRemoveFlexgridBeforeSelChangeHandler();
            //    GUI.m_APE.SendMessages(EventSet.APE);
            //    GUI.m_APE.WaitForMessages(EventSet.APE);
            //}

            //if (removeAfterHandler)
            //{
            //    if (SelectedColumn() == columnIndex && SelectedRow() == rowIndex)
            //    {
            //        GUI.m_APE.AddFirstyMessageWaitForAndRemoveFlexgridAfterSelChangeHandler();
            //        GUI.m_APE.SendMessages(EventSet.APE);
            //        GUI.m_APE.WaitForMessages(EventSet.APE);
            //    }
            //    else
            //    {
            //        GUI.m_APE.AddFirstyMessageRemoveFlexgridAfterSelChangeHandler();
            //        GUI.m_APE.SendMessages(EventSet.APE);
            //        GUI.m_APE.WaitForMessages(EventSet.APE);
            //    }
            //}
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

            //Check to make sure the row isn't hidden
            if (IsRowHidden(rowIndex))
            {
                throw new Exception("Row is hidden in the " + Description);
            }

            //Check to make sure the column isn't hidden
            if (IsColumnHidden(columnIndex))
            {
                throw new Exception("Column is hidden in the " + Description);
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
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <returns>The bound rectangle</returns>
        internal override Rectangle GetCellRectangleInternal(int rowIndex, int columnIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRect", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
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
            int x = GUI.m_APE.GetValueFromMessage();
            int y = GUI.m_APE.GetValueFromMessage();
            int width = GUI.m_APE.GetValueFromMessage();
            int height = GUI.m_APE.GetValueFromMessage();

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
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "FindRow", MemberTypes.Method, new Parameter(GUI.m_APE, rowText), new Parameter(GUI.m_APE, CurrentRow), new Parameter(GUI.m_APE, columnIndex), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, false));
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
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
                    string CurrentRow = this.GetCellRange(Row, 0, Row, Columns - 1, CellProperty.TextDisplay);

                    if (this.Columns() != Columns)
                    {
                        throw new Exception("Number of grid columns changed in the " + Description);
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

        /// <summary>
        /// Gets the tree view column
        /// </summary>
        /// <returns>The tree view column index or -1</returns>
        public int TreeViewColumn()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Tree", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Column", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int TreeViewColumn = GUI.m_APE.GetValueFromMessage();

            return TreeViewColumn;
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

        /// <summary>
        /// Gets the indentation in pixels of a single tree level
        /// </summary>
        /// <returns>The indentation in pixels</returns>
        public int TreeViewIndent()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Tree", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Indent", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int TreeViewIndent = GUI.m_APE.GetValueFromMessage();

            return TreeViewIndent;
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Node", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Level", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "BorderStyle", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string BorderStyle = GUI.m_APE.GetValueFromMessage();

            switch (BorderStyle)
            {
                case "None":
                    return 0;
                case "FixedSingle":
                    return 1;
                case "Fixed3D":
                    return 2;
                case "Light3D":
                    return 1;
                case "XpThemes":
                    return 1;
            }

            throw new Exception("Unknown BorderStyle: " + BorderStyle);
        }

        //TODO
        //FlexGridFullyCollapseTreeView //do in a more generic way?
        //FlexGridFullyExpandTreeView

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
            string[] separatorComma = { "," };
            string[] separatorCr = { "\r" };
            string[] separatorTab = { "\t" };

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllColumnsHidden(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string columnsHiddenText = GUI.m_APE.GetValueFromMessage();
            string[] columnsHiddenTextArray = columnsHiddenText.Split(separatorComma, StringSplitOptions.None);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllRowsHidden(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string rowsHiddenText = GUI.m_APE.GetValueFromMessage();
            string[] rowsHiddenTextArray = rowsHiddenText.Split(separatorComma, StringSplitOptions.None);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllColumnsWidth(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string columnsWidthText = GUI.m_APE.GetValueFromMessage();
            string[] columnsWidthTextArray = columnsWidthText.Split(separatorComma, StringSplitOptions.None);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllRowsHeight(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string rowsHeightText = GUI.m_APE.GetValueFromMessage();
            string[] rowsHeightTextArray = rowsHeightText.Split(separatorComma, StringSplitOptions.None);

            //TODO exclude collapsed node rows

            int columns = this.Columns();
            int rows = this.Rows();

            bool[] columnsToExclude = new bool[columns];
            bool[] rowsToExclude = new bool[rows];

            for (int column = 0; column < columns; column++)
            {
                if (columnsHiddenTextArray[column] == "True")
                {
                    columnsToExclude[column] = true;
                }
                if (columnsWidthTextArray[column] == "0" || columnsWidthTextArray[column] == "1")
                {
                    columnsToExclude[column] = true;
                }
            }

            for (int row = 0; row < rows; row++)
            {
                if (rowsHiddenTextArray[row] == "True")
                {
                    rowsToExclude[row] = true;
                }
                if (rowsHeightTextArray[row] == "0" || rowsHeightTextArray[row] == "1")
                {
                    rowsToExclude[row] = true;
                }
            }

            // Copy the whole grid
            if (rows > 0 && columns > 0)
            {
                string fullGrid = GetCellRange(0, 0, rows - 1, columns - 1, property);

                StringBuilder grid = new StringBuilder(10240);
                bool doneColumn;

                string[] fullGridRows = fullGrid.Split(separatorCr, StringSplitOptions.None);
                for (int fullGridRow = 0; fullGridRow < rows; fullGridRow++)
                {
                    if (!rowsToExclude[fullGridRow])
                    {
                        string[] fullGridColumns = fullGridRows[fullGridRow].Split(separatorTab, StringSplitOptions.None);
                        doneColumn = false;
                        for (int fullGridColumn = 0; fullGridColumn < columns; fullGridColumn++)
                        {
                            if (!columnsToExclude[fullGridColumn])
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
        /// THIS WILL BE REMOVED SOMETIME IN 2017 DO NOT USE IT
        /// Returns a range of cell values column separated by \t and row separated by \r where
        /// each cell has a width and height greater than 1 pixel and the row or column is not
        /// hidden.  Collapsed nodes are also excluded.
        /// </summary>
        /// <param name="property">The property of the cell to get</param>
        /// <param name="excludeWidthHeightSmallerThan2">Exclude columns and rows with widths or heights less than 2 pixels</param>
        /// <returns>A string containing the range of values</returns>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public string GetAllVisibleCells(CellProperty property, bool excludeWidthHeightSmallerThan2)
        {
            string[] separatorComma = { "," };
            string[] separatorCr = { "\r" };
            string[] separatorTab = { "\t" };

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllColumnsHidden(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string columnsHiddenText = GUI.m_APE.GetValueFromMessage();
            string[] columnsHiddenTextArray = columnsHiddenText.Split(separatorComma, StringSplitOptions.None);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllRowsHidden(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string rowsHiddenText = GUI.m_APE.GetValueFromMessage();
            string[] rowsHiddenTextArray = rowsHiddenText.Split(separatorComma, StringSplitOptions.None);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllColumnsWidth(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string columnsWidthText = GUI.m_APE.GetValueFromMessage();
            string[] columnsWidthTextArray = columnsWidthText.Split(separatorComma, StringSplitOptions.None);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageFlexgridGetAllRowsHeight(DataStores.Store0, DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string rowsHeightText = GUI.m_APE.GetValueFromMessage();
            string[] rowsHeightTextArray = rowsHeightText.Split(separatorComma, StringSplitOptions.None);

            //TODO exclude collapsed node rows

            int columns = this.Columns();
            int rows = this.Rows();

            bool[] columnsToExclude = new bool[columns];
            bool[] rowsToExclude = new bool[rows];

            for (int column = 0; column < columns; column++)
            {
                if (columnsHiddenTextArray[column] == "True")
                {
                    columnsToExclude[column] = true;
                }
                if (excludeWidthHeightSmallerThan2)
                {
                    if (columnsWidthTextArray[column] == "0" || columnsWidthTextArray[column] == "1")
                    {
                        columnsToExclude[column] = true;
                    }
                }
            }

            for (int row = 0; row < rows; row++)
            {
                if (rowsHiddenTextArray[row] == "True")
                {
                    rowsToExclude[row] = true;
                }
                if (excludeWidthHeightSmallerThan2)
                {
                    if (rowsHeightTextArray[row] == "0" || rowsHeightTextArray[row] == "1")
                    {
                        rowsToExclude[row] = true;
                    }
                }
            }

            // Copy the whole grid
            if (rows > 0 && columns > 0)
            {
                string fullGrid = GetCellRange(0, 0, rows - 1, columns - 1, property);

                StringBuilder grid = new StringBuilder(10240);
                bool doneColumn;

                string[] fullGridRows = fullGrid.Split(separatorCr, StringSplitOptions.None);
                for (int fullGridRow = 0; fullGridRow < rows; fullGridRow++)
                {
                    if (!rowsToExclude[fullGridRow])
                    {
                        string[] fullGridColumns = fullGridRows[fullGridRow].Split(separatorTab, StringSplitOptions.None);
                        doneColumn = false;
                        for (int fullGridColumn = 0; fullGridColumn < columns; fullGridColumn++)
                        {
                            if (!columnsToExclude[fullGridColumn])
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

        //private bool CellHasComboList()
        //{
        //    //TODO
        //}

        //public void ExpandNodes()
        //{
        //    //TODO
        //}

        //public string ColumnFormat()
        //{
        //    //TODO
        //}

        //public string ColumnComboList()
        //{
        //    //TODO
        //}

        //public bool IsParentNode()
        //{
        //    //TODO
        //}

        //public int SelectedRowsContains()
        //{
        //    //TODO
        //}

        //Rows.Selected.Count
        //Rows.Selected.Item(x).Index

        //public int SelectedRows()
        //{
        //    //TODO
        //}

        //public void GetSelectedRows()
        //{
        //    //TODO
        //}

        //public void HeaderCheckedStatus()
        //{
        //    //TODO
        //}
    }
}