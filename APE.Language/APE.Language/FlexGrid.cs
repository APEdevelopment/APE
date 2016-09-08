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
            GUI.Log("Fully expand the treeview", LogItemTypeEnum.Action);

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
            GUI.Log("Fully collapse the treeview", LogItemTypeEnum.Action);

            int treeColumnIndex = TreeViewColumn();

            if (treeColumnIndex == -1)
            {
                throw new Exception("Flexgrid is not a treeview");
            }

            for (int rowIndex = Rows() -1; rowIndex > -1; rowIndex--)
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
            GUI.Log("Expand node " + nodePath, LogItemTypeEnum.Action);

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
            GUI.Log("Collapse node " + nodePath, LogItemTypeEnum.Action);

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
                GUI.Log("Ensure " + m_DescriptionOfControl + " row " + rowText + " column " + columnText + " is set to " + expectedValue, LogItemTypeEnum.Action);
                return false;
            }

            // Get the data type of the cell we want to set
            string cellDataType = this.GetCellValue(rowIndex, columnIndex, CellProperty.DataType);

            switch (cellDataType)
            {
                //case "System.DateTime":
                case "System.Boolean":  //checkbox
                    // Click on the checkbox
                    GUI.Log("Single " + MouseButton.Left.ToString() + " click on the checkbox in the " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemTypeEnum.Action);
                    this.SelectInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    break;
                default:
                    // Select the cell if its not selected
                    if (this.SelectedRow() == rowIndex && this.SelectedColumn() == columnIndex)
                    {
                        GUI.Log("Ensure " + m_DescriptionOfControl + " row " + rowText + " column " + columnText + " is selected", LogItemTypeEnum.Action);
                    }
                    else
                    {
                        GUI.Log("Single " + MouseButton.Left.ToString() + " click on " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemTypeEnum.Action);
                        this.SelectInternal(rowIndex, columnIndex, MouseButton.Left, CellClickLocation.CentreOfCell, MouseKeyModifier.None);
                    }

                    // Put the cell into edit mode
                    GUI.Log("Press F2 to enter edit mode", LogItemTypeEnum.Action);
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
                        throw new Exception("Could not find the flexgrid cell editor");
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
                                GUI.Log("Press " + submitKey + " to set the value", LogItemTypeEnum.Action);
                                base.SendKeysInternal(submitKey);
                                break;
                            case "GUIDateTimePicker":
                                GUIDateTimePicker flexgridDateTimePicker = new GUIDateTimePicker(m_ParentForm, m_DescriptionOfControl + " datetime picker", new Identifier(Identifiers.Handle, EditorHandle));
                                flexgridDateTimePicker.SetText(value);
                                GUI.Log("Press " + submitKey + " to set the value", LogItemTypeEnum.Action);
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

        public dynamic GetCellValue(string rowText, string columnText, CellProperty property)
        {
            int row = FindRow(rowText);
            int column = FindColumn(columnText);
            return GetCellValue(row, column, property);
        }

        public dynamic GetCellValue(int row, string columnText, CellProperty Property)
        {
            int column = FindColumn(columnText);
            return GetCellValue(row, column, Property);
        }

        public dynamic GetCellValue(string rowText, int column, CellProperty property)
        {
            int row = FindRow(rowText);
            return GetCellValue(row, column, property);
        }

        public dynamic GetCellValue(int row, int column, CellProperty property)
        {
            switch (property)
            {
                case CellProperty.TextDisplay:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "DataDisplay", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string CellDataDisplay = GUI.m_APE.GetValueFromMessage();
                    return CellDataDisplay;
                case CellProperty.BackColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
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
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
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
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
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
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
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
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Image", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Image image = GUI.m_APE.GetValueFromMessage();
                    return image;
                case CellProperty.BackgroundImage:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row), new Parameter(GUI.m_APE, column));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "BackgroundImage", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Image backgroundImage = GUI.m_APE.GetValueFromMessage();
                    return backgroundImage;
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }
        }

        public string GetCellRange(int row1, string columnText1, int row2, string columnText2, CellProperty property)
        {
            int column1 = FindColumn(columnText1);
            int column2 = FindColumn(columnText2);
            return GetCellRange(row1, column1, row2, column2, property);
        }

        public string GetCellRange(int row1, int column1, int row2, string columnText2, CellProperty property)
        {
            int column2 = FindColumn(columnText2);
            return GetCellRange(row1, column1, row2, column2, property);
        }

        public string GetCellRange(int row1, string columnText1, int row2, int column2, CellProperty property)
        {
            int column1 = FindColumn(columnText1);
            return GetCellRange(row1, column1, row2, column2, property);
        }

        public string GetCellRange(int row1, int column1, int row2, int column2, CellProperty property)
        {
            switch (property)
            {
                case CellProperty.TextDisplay:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, row1), new Parameter(GUI.m_APE, column1), new Parameter(GUI.m_APE, row2), new Parameter(GUI.m_APE, column2));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Clip", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeClip = GUI.m_APE.GetValueFromMessage();
                    return rangeClip;
                case CellProperty.BackColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1, column1, row2, column2, APEIPC.CellProperty.BackColourName);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeBackColourName = GUI.m_APE.GetValueFromMessage();
                    return rangeBackColourName;
                case CellProperty.ForeColourName:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1, column1, row2, column2, APEIPC.CellProperty.ForeColourName);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string rangeForeColourName = GUI.m_APE.GetValueFromMessage();
                    return rangeForeColourName;
                case CellProperty.DataType:
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageFlexgridGetCellRange(DataStores.Store0, DataStores.Store1, row1, column1, row2, column2, APEIPC.CellProperty.DataType);
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

            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemTypeEnum.Action);
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

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowText + " column " + columnText, LogItemTypeEnum.Action);
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

            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnText, LogItemTypeEnum.Action);
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

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnText, LogItemTypeEnum.Action);
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

            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowText + " column " + columnIndex.ToString(), LogItemTypeEnum.Action);
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

            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowText + " column " + columnIndex.ToString(), LogItemTypeEnum.Action);
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
            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemTypeEnum.Action);
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
            GUI.Log("Single " + button.ToString() + " click while pressinig key " + keyModifier.ToString() + " on " + m_DescriptionOfControl + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemTypeEnum.Action);
            SelectInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        private void SelectInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
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

        public int FindRow(string Row)
        {
            int ColumnNumber = FirstVisibleColumn();
            int StartAtRow = 0; //FixedRows();
            return FindRow(Row, ColumnNumber, StartAtRow);
        }

        public int FindRow(string Row, string Column)
        {
            int ColumnNumber = FindColumn(Column);
            int StartAtRow = 0;//FixedRows();
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
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "FindRow", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, CurrentRow), new Parameter(GUI.m_APE, Column), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, false));
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

        ////FlexGridFullyCollapseTreeView //do in a more generic way?
        ////FlexGridFullyExpandTreeView

        //public string GetAllVisibleCells()
        //{
        //    //TODO
        //}

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
