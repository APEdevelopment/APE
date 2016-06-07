using System;
using System.Drawing;
using System.Text;
using System.Reflection;
using APE.Capture;
using APE.Communication;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// C1.Win.C1FlexGrid.C1FlexGrid
    /// </summary>
    public sealed class GUIFlexgrid : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Editor", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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

            return GUI.m_APE.GetWindowText(EditWindowHandle);
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

            do
            {
                //Find a row which matches the child node we are after
                CurrentRow = FindRowInternal(ChildNodeText, Column, CurrentRow);

                if (CurrentRow == -1)
                {
                    throw new Exception("Failed to find child node");
                }

                //Check if its parents match the node we are after
                CurrentNodeText = GetNodePath(CurrentRow, Column);
            } while (CurrentNodeText != NodeText);

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

        private bool IsNode(int Row)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Row));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "IsNode", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool RowIsNode = GUI.m_APE.GetValueFromMessage();

            return RowIsNode;
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

        public bool IsRowHidden(string Row)
        {
            int RowNumber = FindRow(Row);
            return IsRowHidden(RowNumber);
        }

        public bool IsRowHidden(int Row)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Row));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool Visible = GUI.m_APE.GetValueFromMessage();

            return !Visible;
        }

        public bool IsColumnHidden(string Column)
        {
            int ColumnNumber = FindColumn(Column);
            return IsColumnHidden(ColumnNumber);
        }

        public bool IsColumnHidden(int Column)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Column));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Visible", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool Visible = GUI.m_APE.GetValueFromMessage();

            return !Visible;
        }

        public bool IsCellVisible(int Row, int Column)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "LeftCol", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store2, "RightCol", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store3, "TopRow", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store4, "BottomRow", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int LeftColumn = GUI.m_APE.GetValueFromMessage();
            int RightColumn = GUI.m_APE.GetValueFromMessage();
            int TopRow = GUI.m_APE.GetValueFromMessage();
            int BottomRow = GUI.m_APE.GetValueFromMessage();

            if (Column <= LeftColumn || Column >= RightColumn || Row <= TopRow || Row >= BottomRow)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

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

        public int SelectedRow()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "RowSel", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedRow = GUI.m_APE.GetValueFromMessage();

            return SelectedRow;
        }

        public int SelectedColumn()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "ColSel", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int SelectedColumn = GUI.m_APE.GetValueFromMessage();

            return SelectedColumn;
        }

        public int FixedRows()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Fixed", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedRows = GUI.m_APE.GetValueFromMessage();

            return FixedRows;
        }

        public int FixedColumns()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Fixed", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedColumns = GUI.m_APE.GetValueFromMessage();

            return FixedColumns;
        }

        public int FrozenColumns()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Frozen", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FrozenColumns = GUI.m_APE.GetValueFromMessage();

            return FrozenColumns;
        }

        public int Rows()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Rows = GUI.m_APE.GetValueFromMessage();

            return Rows;
        }

        public int Columns()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Cols", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Columns = GUI.m_APE.GetValueFromMessage();

            return Columns;
        }

        public void SetCellValue(string Row, string Column, string Value)
        {
            SetCellValue(Row, Column, Value, null, null);
        }

        public void SetCellValue(string Row, string Column, string Value, string ExpectedValue)
        {
            SetCellValue(Row, Column, Value, ExpectedValue, null);
        }

        public void SetCellValue(string Row, string Column, string Value, string ExpectedValue, string submitKey)
        {
            this.Select(Row, Column, MouseButton.Left, CellClickLocation.CentreOfCell);
            int RowNumber = FindRow(Row);
            int ColumnNumber = FindColumn(Column);
            SetCellValueInternal(RowNumber, ColumnNumber, Value, ExpectedValue, submitKey);
        }

        public void SetCellValue(int Row, string Column, string Value)
        {
            SetCellValue(Row, Column, Value, null, null);
        }

        public void SetCellValue(int Row, string Column, string Value, string ExpectedValue)
        {
            SetCellValue(Row, Column, Value, ExpectedValue, null);
        }

        public void SetCellValue(int Row, string Column, string Value, string ExpectedValue, string submitKey)
        {
            this.Select(Row, Column, MouseButton.Left, CellClickLocation.CentreOfCell);
            int ColumnNumber = FindColumn(Column);
            SetCellValueInternal(Row, ColumnNumber, Value, ExpectedValue, submitKey);
        }

        public void SetCellValue(string Row, int Column, string Value)
        {
            SetCellValue(Row, Column, Value, null, null);
        }

        public void SetCellValue(string Row, int Column, string Value, string ExpectedValue)
        {
            SetCellValue(Row, Column, Value, ExpectedValue, null);
        }

        public void SetCellValue(string Row, int Column, string Value, string ExpectedValue, string submitKey)
        {
            this.Select(Row, Column, MouseButton.Left, CellClickLocation.CentreOfCell);
            int RowNumber = FindRow(Row);
            SetCellValueInternal(RowNumber, Column, Value, ExpectedValue, submitKey);
        }

        public void SetCellValue(int Row, int Column, string Value)
        {
            SetCellValue(Row, Column, Value, null, null);
        }

        public void SetCellValue(int Row, int Column, string Value, string ExpectedValue)
        {
            SetCellValue(Row, Column, Value, ExpectedValue, null);
        }

        public void SetCellValue(int Row, int Column, string Value, string ExpectedValue, string submitKey)
        {
            this.Select(Row, Column, MouseButton.Left, CellClickLocation.CentreOfCell);
            SetCellValueInternal(Row, Column, Value, ExpectedValue, submitKey);
        }

        private void SetCellValueInternal(int row, int column, string value, string expectedValue, string submitKey)
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

            // Check if the cell is already set to the correct value
            string CurrentValue = this.GetCellValue(row, column, CellProperty.TextDisplay);

            if (CurrentValue == expectedValue)
            {
                GUI.Log("Ensure the cell is set to " + expectedValue, LogItemTypeEnum.Action);
                return;
            }

            // Put the cell into edit mode
            GUI.Log("Press F2 to enter edit mode", LogItemTypeEnum.Action);
            base.SendKeys("{F2}");

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Editor", MemberTypes.Property);
            GUI.m_APE.AddMessageGetRecognisedType(DataStores.Store1, DataStores.Store2);
            GUI.m_APE.AddMessageGetApeTypeFromObject(DataStores.Store1, DataStores.Store3); //Get this to help with debuging
            GUI.m_APE.AddMessageGetApeTypeFromType(DataStores.Store2, DataStores.Store4);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store5, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddMessageGetValue(DataStores.Store5);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call;
            string APEDirectType = GUI.m_APE.GetValueFromMessage();
            string APEBaseType = GUI.m_APE.GetValueFromMessage();
            dynamic EditorHandle = GUI.m_APE.GetValueFromMessage();

            if (EditorHandle == null)
            {
                throw new Exception("Could not find the flexgrid cell editor");
            }

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
                    GUI.Log("Press Enter to set the value", LogItemTypeEnum.Action);
                    base.SendKeys(submitKey);
                    break;
                default:
                    throw new Exception("Unsupported flexgrid editor: Type: " + APEDirectType + " Base Type: " + APEBaseType);
            }

            //Check the value was set
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                CurrentValue = this.GetCellValue(row, column, CellProperty.TextDisplay);

                if (CurrentValue == expectedValue)
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
        }

        public string GetCellRangeClip(string Row1, string Column1, string Row2, string Column2)
        {
            int RowNumber1 = FindRow(Row1);
            int ColumnNumber1 = FindColumn(Column1);
            int RowNumber2 = FindRow(Row2);
            int ColumnNumber2 = FindColumn(Column2);

            return GetCellRangeClip(RowNumber1, ColumnNumber1, RowNumber2, ColumnNumber2);
        }

        public string GetCellRangeClip(int Row1, int Col1, int Row2, int Col2)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, Row1), new Parameter(GUI.m_APE, Col1), new Parameter(GUI.m_APE, Row2), new Parameter(GUI.m_APE, Col2));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Clip", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string CellRangeClip = GUI.m_APE.GetValueFromMessage();

            return CellRangeClip;
        }

        public enum CellProperty
        {
            TextDisplay,
            BackColor,
            ForeColor,
        }

        public dynamic GetCellValue(string Row, string Column, CellProperty Property)
        {
            int RowNumber = FindRow(Row);
            int ColumnNumber = FindColumn(Column);
            return GetCellValue(RowNumber, ColumnNumber, Property);
        }

        public dynamic GetCellValue(int Row, string Column, CellProperty Property)
        {
            int ColumnNumber = FindColumn(Column);
            return GetCellValue(Row, ColumnNumber, Property);
        }

        public dynamic GetCellValue(string Row, int Column, CellProperty Property)
        {
            int RowNumber = FindRow(Row);
            return GetCellValue(RowNumber, Column, Property);
        }

        public dynamic GetCellValue(int Row, int Column, CellProperty Property)
        {
            switch (Property)
            {
                case CellProperty.TextDisplay:
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, Column));
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "DataDisplay", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string CellDataDisplay = GUI.m_APE.GetValueFromMessage();
                    return CellDataDisplay;
                case CellProperty.BackColor:
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, Column));
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "BackColor", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Color CellBackColor = GUI.m_APE.GetValueFromMessage();
                    return CellBackColor;
                case CellProperty.ForeColor:
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetCellRange", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, Column));
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "StyleDisplay", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ForeColor", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Color CellForeColor = GUI.m_APE.GetValueFromMessage();
                    return CellForeColor;
                default:
                    throw new Exception("Implement support for getting cell property " + Property.ToString());
            }
        }

        public string GetCellCheck(string Row, string Column)
        {
            int RowNumber = FindRow(Row);
            int ColumnNumber = FindColumn(Column);

            return GetCellCheck(RowNumber, ColumnNumber);
        }

        public string GetCellCheck(int Row, string Column)
        {
            int ColumnNumber = FindColumn(Column);

            return GetCellCheck(Row, ColumnNumber);
        }

        public string GetCellCheck(string Row, int Column)
        {
            int RowNumber = FindRow(Row);

            return GetCellCheck(RowNumber, Column);
        }

        public string GetCellCheck(int Row, int Column)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetCellCheck", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, Column));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string Checked = GUI.m_APE.GetValueFromMessage();

            return Checked;
        }

        public void Show(string Row, string Column)
        {
            int RowNumber = FindRow(Row);
            int ColumnNumber = FindColumn(Column);
            Show(RowNumber, ColumnNumber);
        }

        public void Show(int Row, int Column)
        {
            if (!IsCellVisible(Row, Column))
            {
                //TODO move mouse over the grid if not in drag mode?
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "ShowCell", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, Column));
                GUI.m_APE.AddMessageGetValue(DataStores.Store1);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                dynamic Ignore = GUI.m_APE.GetValueFromMessage();
            }
        }

        public void Select(string Row, string Column, MouseButton Button, CellClickLocation LocationInCell)
        {
            int RowNumber = FindRow(Row);
            int ColumnNumber = FindColumn(Column);

            GUI.Log("Single " + Button.ToString() + " click on " + m_DescriptionOfControl + " row " + Row + " column " + Column, LogItemTypeEnum.Action);
            SelectInternal(RowNumber, ColumnNumber, Button, LocationInCell);
        }

        public void Select(int Row, string Column, MouseButton Button, CellClickLocation LocationInCell)
        {
            int ColumnNumber = FindColumn(Column);

            GUI.Log("Single " + Button.ToString() + " click on " + m_DescriptionOfControl + " row " + Row.ToString() + " column " + Column, LogItemTypeEnum.Action);
            SelectInternal(Row, ColumnNumber, Button, LocationInCell);
        }

        public void Select(string Row, int Column, MouseButton Button, CellClickLocation LocationInCell)
        {
            int RowNumber = FindRow(Row);

            GUI.Log("Single " + Button.ToString() + " click on " + m_DescriptionOfControl + " row " + Row + " column " + Column.ToString(), LogItemTypeEnum.Action);
            SelectInternal(RowNumber, Column, Button, LocationInCell);
        }

        public void Select(int Row, int Column, MouseButton Button, CellClickLocation LocationInCell)
        {
            GUI.Log("Single " + Button.ToString() + " click on " + m_DescriptionOfControl + " row " + Row.ToString() + " column " + Column.ToString(), LogItemTypeEnum.Action);
            SelectInternal(Row, Column, Button, LocationInCell);
        }

        private void SelectInternal(int Row, int Column, MouseButton Button, CellClickLocation LocationInCell)
        {
            Point Location = GetLocationInCell(Row, Column, LocationInCell);
            base.MouseSingleClickInternal(Location.X, Location.Y, Button, MouseKeyModifier.None);
        }

        public void MouseMove(int Row, int Column, CellClickLocation LocationInCell)
        {
            Point Location = GetLocationInCell(Row, Column, LocationInCell);
            base.MouseMove(Location.X, Location.Y);
        }

        private Point GetLocationInCell(int Row, int Column, CellClickLocation LocationInCell)
        {
            //Check to make sure the row isn't hidden
            if (IsRowHidden(Row))
            {
                throw new Exception("Row is hidden");
            }

            //Check to make sure the column isn't hidden
            if (IsColumnHidden(Column))
            {
                throw new Exception("Column is hidden");
            }
    
            //Scroll the cell into view
            Show(Row, Column);

            Rectangle CellRectangle = GetCellRectangle(Row, Column);
            
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
                    Location.X = CellRectangle.Left + 5 + 2 + BorderWidth() + (TreeViewIndent() * NodeLevel(Row));
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                case CellClickLocation.LeftSideOfTreeItem:
                    Location.X = CellRectangle.Left + 5 + 2 + BorderWidth() + (TreeViewIndent() * (NodeLevel(Row) + 1));
                    Location.Y = CellRectangle.Top + (CellRectangle.Height / 2);
                    break;
                default:
                    throw new Exception("Implement for CellClickLocation: " + Location.ToString());
            }

            return Location;
        }

        public Rectangle GetCellRectangle(string Row, string Column)
        {
            int RowNumber = FindRow(Row);
            int ColumnNumber = FindColumn(Column);
            return GetCellRectangle(RowNumber, ColumnNumber);
        }

        public Rectangle GetCellRectangle(int Row, string Column)
        {
            int ColumnNumber = FindColumn(Column);
            return GetCellRectangle(Row, ColumnNumber);
        }

        public Rectangle GetCellRectangle(string Row, int Column)
        {
            int RowNumber = FindRow(Row);
            return GetCellRectangle(RowNumber, Column);
        }

        public Rectangle GetCellRectangle(int Row, int Column)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetCellRect", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, Column));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "X", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store3, "Y", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store4, "Width", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store5, "Height", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.AddMessageGetValue(DataStores.Store5);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int X = GUI.m_APE.GetValueFromMessage();
            int Y = GUI.m_APE.GetValueFromMessage();
            int Width = GUI.m_APE.GetValueFromMessage();
            int Height = GUI.m_APE.GetValueFromMessage();

            Rectangle CellRectangle = new Rectangle(X, Y, Width, Height);
            return CellRectangle;
        }

        //TODO if there is a treeview column, use that, otherwise use first visible
        public int FindRow(string Row)
        {
            int ColumnNumber = FirstVisibleColumn();
            int StartAtRow = 0;//FixedRows();
            return FindRow(Row, ColumnNumber, StartAtRow);
        }

        public int FindRow(int Row)
        {
            return Row;
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

        public int FindRow(int Row, string Column)
        {
            return Row;
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
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "FindRow", MemberTypes.Method, new Parameter(GUI.m_APE, Row), new Parameter(GUI.m_APE, CurrentRow), new Parameter(GUI.m_APE, Column), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, true), new Parameter(GUI.m_APE, false));
                GUI.m_APE.AddMessageGetValue(DataStores.Store1);
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

        public int FindRow(int Row, int Column)
        {
            return Row;
        }

        public int FindColumn(int Column)
        {
            return Column;
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
                    string CurrentRow = this.GetCellRangeClip(Row, 0, Row, Columns - 1);

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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Tree", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Column", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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

        public int TreeViewIndent()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Tree", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Indent", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int TreeViewIndent = GUI.m_APE.GetValueFromMessage();

            return TreeViewIndent;
        }

        public int NodeLevel(string Row)
        {
            int RowNumber = FindRow(Row);
            return NodeLevel(RowNumber);
        }

        public int NodeLevel(int Row)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Rows", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Row));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Node", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Level", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int NodeLevel = GUI.m_APE.GetValueFromMessage();

            return NodeLevel;
        }

        public int BorderWidth()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "BorderStyle", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
