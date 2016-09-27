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

namespace APE.Language
{
    /// <summary>
    /// Specifies what cell(s) property to get
    /// </summary>
    public enum CellProperty
    {
        /// <summary>
        /// Cell(s) value formatted as a string
        /// </summary>
        TextDisplay,
        /// <summary>
        /// Cells(s) back colour name as a string
        /// </summary>
        BackColourName,
        /// <summary>
        /// Cells(s) fore colour name as a string
        /// </summary>
        ForeColourName,
        /// <summary>
        /// The data type of the cell as a string
        /// Supported in GetCellValue only
        /// </summary>
        DataType,
        /// <summary>
        /// Cells(s) checkbox state as a string
        /// </summary>
        CheckBox,
        /// <summary>
        /// For GetCellValue this returns the actual image
        /// For GetCellRange this returns whether the cells contain a image
        /// </summary>
        Image,
        /// <summary>
        /// For GetCellValue this returns the actual background image
        /// For GetCellRange this returns whether the cells contain a background image
        /// </summary>
        BackgroundImage,
        /// <summary>
        /// The font style of the cell
        /// Supported in GetCellValue only
        /// </summary>
        FontStyle,
    }

    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// C1.Win.C1FlexGrid.C1FlexGrid
    /// </summary>
    public sealed class GUIFlexgrid : GUIFocusableObject
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
                throw new Exception("Flexgrid is not in edit mode");
            }

            return GUI.m_APE.GetWindowTextViaWindowMessage(EditWindowHandle);
        }

        private int FindNodeRow(string NodeText)
        {
            //TODO this is bit inefficent, it would be better to find from parent to child rather than child to parent
            string[] Separator = new string[] { " -> " };
            string[] SplitNodeText = NodeText.Split(Separator, StringSplitOptions.None);

            int Column = TreeViewColumn();
            int CurrentRow = FixedRows();
            string ChildNodeText = SplitNodeText[SplitNodeText.GetUpperBound(0)];
            string CurrentNodeText;

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                //Find a row which matches the child node we are after
                CurrentRow = FindRowInternal(ChildNodeText, Column, CurrentRow);

                if (CurrentRow == -1)
                {
                    break;
                }

                //Check if its parents match the node we are after
                CurrentNodeText = GetNodePath(CurrentRow, Column);

                if (CurrentNodeText == NodeText)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to find node row");
                }
            }

            return CurrentRow;
        }

        private string GetNodePath(int Row, int Column)
        {
            string NodePath = GetCellValue(Row, Column, CellProperty.TextDisplay);
            int CurrentRow = GetNodeParentRow(Row);

            while (CurrentRow > -1)
            {
                NodePath = GetCellValue(CurrentRow, Column, CellProperty.TextDisplay) + " -> " + NodePath;
                CurrentRow = GetNodeParentRow(CurrentRow);
            }

            return NodePath;
        }

        private bool IsNode(int rowIndex)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
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
                throw new Exception("Flexgrid is not a treeview");
            }

            for (int rowIndex = 0; rowIndex < Rows(); rowIndex++)
            {
                if (IsNode(rowIndex))
                {
                    if (HasNodeGotChildren(rowIndex))
                    {
                        if (IsNodeCollapsed(rowIndex))
                        {
                            SelectInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                            Stopwatch timer = Stopwatch.StartNew();
                            while (true)
                            {
                                if (!IsNodeCollapsed(rowIndex))
                                {
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw new Exception("Failed to expand the flexgrid node");
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
                throw new Exception("Flexgrid is not a treeview");
            }

            for (int rowIndex = Rows() - 1; rowIndex > -1; rowIndex--)
            {
                if (IsNode(rowIndex))
                {
                    if (HasNodeGotChildren(rowIndex))
                    {
                        if (!IsNodeCollapsed(rowIndex))
                        {
                            SelectInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                            Stopwatch timer = Stopwatch.StartNew();
                            while (true)
                            {
                                if (IsNodeCollapsed(rowIndex))
                                {
                                    break;
                                }

                                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                                {
                                    throw new Exception("Failed to collapse the flexgrid node");
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
        /// <param name="nodePath">Node path to expand delimited by -> for example RULE LIBRARY -> UCITS I -> France</param>
        public void ExpandNodes(string nodePath)
        {
            GUI.Log("Expand node " + nodePath, LogItemType.Action);

            string[] delimiter = { " -> " };
            string[] nodePathArray = nodePath.Split(delimiter, StringSplitOptions.None);

            int treeColumnIndex = TreeViewColumn();
            string row = "";

            if (treeColumnIndex == -1)
            {
                throw new Exception("Flexgrid is not a treeview");
            }

            for (int i = 0; i < nodePathArray.GetLength(0); i++)
            {
                if (i == 0)
                {
                    row = nodePathArray[0];
                }
                else
                {
                    row += " -> " + nodePathArray[i];
                }
                int rowIndex = FindNodeRow(row);

                if (!HasNodeGotChildren(rowIndex))
                {
                    throw new Exception("Can not expand node is has no children");
                }

                if (IsNodeCollapsed(rowIndex))
                {
                    SelectInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (!IsNodeCollapsed(rowIndex))
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to expand the flexgrid node");
                        }

                        Thread.Sleep(15);
                    }
                }
            }
        }

        /// <summary>
        /// Collapses the specified nodes in the flexgrid treeview column
        /// </summary>
        /// <param name="nodePath">Node path to collapse delimited by -> for example RULE LIBRARY -> UCITS I -> France</param>
        public void CollapseNodes(string nodePath)
        {
            GUI.Log("Collapse node " + nodePath, LogItemType.Action);

            string[] delimiter = { " -> " };
            string[] nodePathArray = nodePath.Split(delimiter, StringSplitOptions.None);

            int treeColumnIndex = TreeViewColumn();
            string row = "";

            if (treeColumnIndex == -1)
            {
                throw new Exception("Flexgrid is not a treeview");
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
                        row += " -> " + nodePathArray[x];
                    }
                }

                int rowIndex = FindNodeRow(row);

                if (!HasNodeGotChildren(rowIndex))
                {
                    throw new Exception("Can not collapse node is has no children");
                }

                if (!IsNodeCollapsed(rowIndex))
                {
                    SelectInternal(rowIndex, treeColumnIndex, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell, MouseKeyModifier.None);

                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (IsNodeCollapsed(rowIndex))
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to collapse the flexgrid node");
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, row));
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, column));
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "RowSel", MemberTypes.Property);
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
            GUI.m_APE.AddMessagePollMember(DataStores.Store0, "RowSel", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, false));
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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ColSel", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedColumn = GUI.m_APE.GetValueFromMessage();

            return SelectedColumn;
        }

        /// <summary>
        /// Returns the number of fixed rows, that is a row which makes up the header of the grid
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
        /// Returns the number of frozen rows, that is a row which doesn't scroll, in the grid
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
        /// Returns the number of fixed columns, that is a column which doesn't scroll, in the grid
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
            string CurrentValue = this.GetCellValue(rowIndex, columnIndex, CellProperty.TextDisplay);
            if (CurrentValue == expectedValue)
            {
                GUI.Log("Ensure " + m_DescriptionOfControl + " row " + rowText + " column " + columnText + " is set to " + expectedValue, LogItemType.Action);
                return false;
            }

            // Get the data type of the cell we want to set
            string cellDataType = this.GetCellValue(rowIndex, columnIndex, CellProperty.DataType);

            switch (cellDataType)
            {
                //case "System.DateTime":
                case "System.Boolean":  //checkbox
                    // Click on the checkbox
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemType.Action);
                    this.SelectInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                default:
                    // Select the cell if its not selected
                    if (this.SelectedRow() == rowIndex && this.SelectedColumn() == columnIndex)
                    {
                        GUI.Log("Ensure " + m_DescriptionOfControl + " row " + rowText + " column " + columnText + " is selected", LogItemType.Action);
                    }
                    else
                    {
                        GUI.Log("Single " + MouseButton.Left.ToString() + " click on " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemType.Action);
                        this.SelectInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemType.Action);
                    base.SendKeysInternal("{F2}");

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
                    string APEDirectType = GUI.m_APE.GetValueFromMessage();
                    string APEBaseType = GUI.m_APE.GetValueFromMessage();
                    dynamic EditorHandle = GUI.m_APE.GetValueFromMessage();

                    if (EditorHandle == null)
                    {
                        EditorHandle = IntPtr.Zero;
                        //throw new Exception("Could not find the flexgrid cell editor");
                    }

                    // If the editor isn't visible then its likely not being used, so search for the real editor
                    if (!NM.IsWindowVisible(EditorHandle))
                    {
                        bool suceeded = false;
                        int timeout = GUI.GetTimeOut();

                        // See if a generic walker exists
                        GUIGenericWalker genericWalker = null;
                        try
                        {
                            GUI.SetTimeOut(0);
                            genericWalker = new GUIGenericWalker(m_ParentForm, m_DescriptionOfControl + " generic walker", new Identifier(Identifiers.Name, "lzGenericWalkerCtl"), new Identifier(Identifiers.SiblingOf, this));
                        }
                        finally
                        {
                            GUI.SetTimeOut(timeout);
                        }
                        if (genericWalker != null && genericWalker.Handle != IntPtr.Zero)
                        {
                            genericWalker.SetText(value);
                            suceeded = true;
                        }

                        if (!suceeded)
                        {
                            throw new Exception("Could not find a visible flexgrid cell editor");
                        }
                    }
                    else
                    {
                        //Set the value
                        switch (APEBaseType)
                        {
                            case "GUIComboBox":
                                GUIComboBox flexgridComboBox = new GUIComboBox(m_ParentForm, m_DescriptionOfControl + " combobox", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridComboBox.ItemSelect(value);
                                break;
                            case "GUITextBox":
                                GUITextBox flexgridTextBox = new GUITextBox(m_ParentForm, m_DescriptionOfControl + " textbox", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridTextBox.SetText(value);
                                GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                                base.SendKeysInternal(submitKey);
                                break;
                            case "GUIDateTimePicker":
                                GUIDateTimePicker flexgridDateTimePicker = new GUIDateTimePicker(m_ParentForm, m_DescriptionOfControl + " datetime picker", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridDateTimePicker.SetText(value);
                                GUI.Log("Press " + submitKey + " to set the value", LogItemType.Action);
                                base.SendKeysInternal(submitKey);
                                break;
                            default:
                                throw new Exception("Unsupported flexgrid editor: Type: " + APEDirectType + " Base Type: " + APEBaseType);
                        }
                    }
                    break;
            }

            //Check the value was set
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                CurrentValue = this.GetCellValue(rowIndex, columnIndex, CellProperty.TextDisplay);

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

        //public string GetColumnType(int col)
        //{
        //    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property); ;
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, col));
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
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, col));
        //    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Format", MemberTypes.Property);
        //    GUI.m_APE.AddMessageGetValue(DataStores.Store3);
        //    GUI.m_APE.SendMessages(EventSet.APE);
        //    GUI.m_APE.WaitForMessages(EventSet.APE);
        //    //Get the value(s) returned MUST be done straight after the WaitForMessages call
        //    string format = GUI.m_APE.GetValueFromMessage();
        //    return format;
        //}

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnText">Column of the cell delimited by -> for example Order -> Id</param>
        /// <returns></returns>
        public string GetCellValue(string rowText, string columnText)
        {
            return GetCellValue(rowText, columnText, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnText">Column of the cell delimited by -> for example Order -> Id</param>
        /// <returns>The cell value as a string</returns>
        /// <returns></returns>
        public string GetCellValue(int rowIndex, string columnText)
        {
            return GetCellValue(rowIndex, columnText, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <returns>The cell value as a string</returns>
        public string GetCellValue(string rowText, int columnIndex)
        {
            return GetCellValue(rowText, columnIndex, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <returns>The cell value as a string</returns>
        public string GetCellValue(int rowIndex, int columnIndex)
        {
            return GetCellValue(rowIndex, columnIndex, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnText">Column of the cell delimited by -> for example Order -> Id</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCellValue(string rowText, string columnText, CellProperty property)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            return GetCellValue(rowIndex, columnIndex, property);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnText">Column of the cell delimited by -> for example Order -> Id</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCellValue(int rowIndex, string columnText, CellProperty property)
        {
            int column = FindColumn(columnText);
            return GetCellValue(rowIndex, column, property);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCellValue(string rowText, int columnIndex, CellProperty property)
        {
            int rowIndex = FindRow(rowText, columnIndex);
            return GetCellValue(rowIndex, columnIndex, property);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCellValue(int rowIndex, int columnIndex, CellProperty property)
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
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }
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
            return GetCellRange(row1Index, column1Text, row2Index, column2Text, CellProperty.TextDisplay);
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
            return GetCellRange(row1Index, column1Index, row2Index, column2Text, CellProperty.TextDisplay);
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
            return GetCellRange(row1Index, column1Text, row2Index, column2Index, CellProperty.TextDisplay);
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
            return GetCellRange(row1Index, column1Index, row2Index, column2Index, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Text">The start column of the range delimited by -> for example Order -> Id</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Text">The end column of the range delimited by -> for example Order -> Id</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, string column1Text, int row2Index, string column2Text, CellProperty property)
        {
            int column1 = FindColumn(column1Text);
            int column2 = FindColumn(column2Text);
            return GetCellRange(row1Index, column1, row2Index, column2, property);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Index">The start column of the range</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Text">The end column of the range delimited by -> for example Order -> Id</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, int column1Index, int row2Index, string column2Text, CellProperty property)
        {
            int column2 = FindColumn(column2Text);
            return GetCellRange(row1Index, column1Index, row2Index, column2, property);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Text">The start column of the range delimited by -> for example Order -> Id</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Index">The end column of the range</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, string column1Text, int row2Index, int column2Index, CellProperty property)
        {
            int column1 = FindColumn(column1Text);
            return GetCellRange(row1Index, column1, row2Index, column2Index, property);
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
        public string GetCellRange(int row1Index, int column1Index, int row2Index, int column2Index, CellProperty property)
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
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }
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
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ShowCell", MemberTypes.Method, new Parameter(GUI.m_APE, rowIndex), new Parameter(GUI.m_APE, columnIndex));
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

            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemType.Action);
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

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemType.Action);
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

            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
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

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
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

            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
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

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
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
            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
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
            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        internal void SelectInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            Point location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            base.MouseSingleClickInternal(location.X, location.Y, button, keyModifier);
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
            base.MouseMove(location.X, location.Y);
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
                rowIndex = FindRowInternal(rowText, columnIndex, startAtRow);
            }

            return rowIndex;
        }

        private int FindRowInternal(string rowText, int columnIndex, int startAtRow)
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
                    string CurrentRow = this.GetCellRange(Row, 0, Row, Columns - 1, CellProperty.TextDisplay);

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
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, rowIndex));
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