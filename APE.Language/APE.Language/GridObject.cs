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
using System.Diagnostics;
using System.Drawing;
using System.Threading;

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
        /// <summary>
        /// The type of the user data associated with a cell (if any) as a string
        /// Supported in GetCellValue only
        /// </summary>
        UserDataType,
    }

    /// <summary>
    /// The location inside a grid cell where to click
    /// </summary>
    public enum CellClickLocation
    {
        /// <summary>
        /// Slightly in from the left side of the cell, centred vertically
        /// </summary>
        LeftSideOfCell = 0,
        /// <summary>
        /// Centre of the cell both horizontally and vertically
        /// </summary>
        CentreOfCell = 1,
        /// <summary>
        /// Slightly in from the right side of the cell, centred vertically
        /// </summary>
        RightSideOfCell = 2,
        /// <summary>
        /// The centre of the expand / collapse icon on a tree view cell
        /// </summary>
        ExpandCollapseIconOfCell = 3,
        /// <summary>
        /// Slightly to the right of the expand / collapse icon of a tree view cell, centred vertically
        /// </summary>
        LeftSideOfTreeItem = 4,
        /// <summary>
        /// The child column picker icon the elementstripgrid
        /// </summary>
        ChildColumnPicker = 5,
        /// <summary>
        /// Slightly in from the top right corner of the cell
        /// </summary>
        TopRightCorner = 6,
    }

    /// <summary>
    /// What checking to perform
    /// </summary>
    public enum ComparisonMethod
    {
        /// <summary>
        /// Compare the value before setting (to check if the value needs to be set) and after setting (to check it has been sucessfully set)
        /// </summary>
        CompareUsingDefaultEqualityComparer,
        /// <summary>
        /// Do not do any checking before or after setting a value
        /// </summary>
        DoNotCompare,
    }

    /// <summary>
    /// Automation class used to automate grid controls
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class GUIGridObject : GUIFocusableObject
    {
        /// <summary>
        /// Initialises a new instance of the GUIElementStripGrid class
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIGridObject(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }       

        /// <summary>
        /// Returns true if the specified column in the grid exists
        /// </summary>
        /// <param name="columnToFind">Column to check if hidden delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns>True or False</returns>
        public virtual bool ColumnExists(string columnToFind)
        {
            string[] columnHeader = columnToFind.Split(GUI.GridDelimiterAsArray, StringSplitOptions.None);

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
        /// <param name="columnToFind">Column to check if hidden delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns>The index of the column</returns>
        public virtual int FindColumn(string columnToFind)
        {
            string[] columnHeader = columnToFind.Split(GUI.GridDelimiterAsArray, StringSplitOptions.None);

            return FindColumn(columnHeader);
        }

        /// <summary>
        /// Returns the index of the specified column in the grid
        /// </summary>
        /// <param name="columnHeader">The column to check</param>
        /// <returns>The index of the column</returns>
        public virtual int FindColumn(string[] columnHeader)
        {
            int column = -1;
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                try
                {
                    column = FindColumnInternal(columnHeader);
                    if (column != -1)
                    {
                        break;
                    }
                }
                catch
                {
                }

                if (timer.ElapsedMilliseconds > 2000)
                {
                    throw GUI.ApeException("Failed to find column " + string.Join(GUI.GridDelimiter, columnHeader) + " in the " + Description);
                }

                Thread.Sleep(50);
            }
            return column;
        }

        internal abstract int FindColumnInternal(string[] columnHeader);

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
        /// <param name="columnText">The column to look for the value in delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
        /// <param name="columnText">The column to look for the value in delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
            return FindRowTemp(rowText, columnIndex, startAtRow);
        }

        // TODO remove this
        internal abstract int FindRowTemp(string rowText, int columnIndex, int startAtRow);
        internal abstract int FindRowInternal(string rowText, int columnIndex, int startAtRow);

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, string columnText, T value)
        {
            return SetCellValue(rowText, columnText, value, value, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, string columnText, T value, ComparisonMethod compareMethod)
        {
            return SetCellValue(rowText, columnText, value, compareMethod, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, string columnText, T value, ComparisonMethod compareMethod, string submitKey)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);
            return SetCellValueInternal(rowText, columnText, rowIndex, columnIndex, value, value, submitKey, compareMethod);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, string columnText, T value, T expectedValue)
        {
            return SetCellValue(rowText, columnText, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, string columnText, T value, T expectedValue, string submitKey)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);
            return SetCellValueInternal(rowText, columnText, rowIndex, columnIndex, value, expectedValue, submitKey, ComparisonMethod.CompareUsingDefaultEqualityComparer);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, string columnText, T value)
        {
            return SetCellValue(rowIndex, columnText, value, value, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, string columnText, T value, ComparisonMethod compareMethod)
        {
            return SetCellValue(rowIndex, columnText, value, compareMethod, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, string columnText, T value, ComparisonMethod compareMethod, string submitKey)
        {
            int columnIndex = FindColumn(columnText);
            return SetCellValueInternal(null, columnText, rowIndex, columnIndex, value, value, submitKey, compareMethod);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, string columnText, T value, T expectedValue)
        {
            return SetCellValue(rowIndex, columnText, value, expectedValue, null);
        }

        /// <summary>
        /// /// Sets the specified cell to the specified value comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnText">The column of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, string columnText, T value, T expectedValue, string submitKey)
        {
            int columnIndex = FindColumn(columnText);
            return SetCellValueInternal(null, columnText, rowIndex, columnIndex, value, expectedValue, submitKey, ComparisonMethod.CompareUsingDefaultEqualityComparer);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, int columnIndex, T value)
        {
            return SetCellValue(rowText, columnIndex, value, value, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, int columnIndex, T value, ComparisonMethod compareMethod)
        {
            return SetCellValue(rowText, columnIndex, value, compareMethod, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, int columnIndex, T value, ComparisonMethod compareMethod, string submitKey)
        {
            int rowIndex = FindRow(rowText, columnIndex);
            return SetCellValueInternal(rowText, null, rowIndex, columnIndex, value, value, submitKey, compareMethod);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, int columnIndex, T value, T expectedValue)
        {
            return SetCellValue(rowText, columnIndex, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowText">The row of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(string rowText, int columnIndex, T value, T expectedValue, string submitKey)
        {
            int rowIndex = FindRow(rowText, columnIndex);
            return SetCellValueInternal(rowText, null, rowIndex, columnIndex, value, expectedValue, submitKey, ComparisonMethod.CompareUsingDefaultEqualityComparer);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, int columnIndex, T value)
        {
            return SetCellValue(rowIndex, columnIndex, value, value, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <returns></returns>
        public bool SetCellValue<T>(int rowIndex, int columnIndex, T value, ComparisonMethod compareMethod)
        {
            return SetCellValue(rowIndex, columnIndex, value, compareMethod, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="compareMethod">Whether or how to compare the value before and after it is set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns></returns>
        public bool SetCellValue<T>(int rowIndex, int columnIndex, T value, ComparisonMethod compareMethod, string submitKey)
        {
            return SetCellValueInternal(null, null, rowIndex, columnIndex, value, value, submitKey, compareMethod);
        }

        /// <summary>
        /// Sets the specified cell to the specified value using {Enter} as the submit key and comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, int columnIndex, T value, T expectedValue)
        {
            return SetCellValue(rowIndex, columnIndex, value, expectedValue, null);
        }

        /// <summary>
        /// Sets the specified cell to the specified value comparing the value before to see if it actually needs setting and after to confirm it has been set
        /// </summary>
        /// <typeparam name="T">The type of the value to set</typeparam>
        /// <param name="rowIndex">The row index of the cell to set the value of</param>
        /// <param name="columnIndex">The column index of the cell to set the value of</param>
        /// <param name="value">The value you wish to set the cell to.  If the cell uses a TextValueWithTypeControl editor then separate the value from the type with a | character.  For instance 1000|Qty</param>
        /// <param name="expectedValue">The value the cell displays after it has been set</param>
        /// <param name="submitKey">The key to press to submit the value you are trying to set</param>
        /// <returns>True if the cell was set or False if it was already set to the value</returns>
        public bool SetCellValue<T>(int rowIndex, int columnIndex, T value, T expectedValue, string submitKey)
        {
            return SetCellValueInternal(null, null, rowIndex, columnIndex, value, expectedValue, submitKey, ComparisonMethod.CompareUsingDefaultEqualityComparer);
        }

        internal abstract bool SetCellValueInternal<T>(string rowText, string columnText, int rowIndex, int columnIndex, T value, T expectedValue, string submitKey, ComparisonMethod compareMethod);

        /// <summary>
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnText">Column text of the cell delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        public void ShowCell(int rowIndex, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            ShowCell(rowIndex, columnIndex);
        }

        /// <summary>
        /// Scrolls the specified cell into view
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        public void ShowCell(int rowIndex, int columnIndex)
        {
            ShowCellInternal(rowIndex, columnIndex);
        }

        internal abstract void ShowCellInternal(int rowIndex, int columnIndex);

        /// <summary>
        /// Moves the mouse cursor over the specified cell, scrolling it into view if need be
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnText">The column text of the cell</param>
        /// <param name="locationInCell">The location in the cell to move the mouse to</param>
        public void MoveToCell(string rowText, string columnText, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Move the mouse over the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            MoveToCellInternal(Location.X, Location.Y);
        }

        /// <summary>
        /// Moves the mouse cursor over the specified cell, scrolling it into view if need be
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnText">The column text of the cell</param>
        /// <param name="locationInCell">The location in the cell to move the mouse to</param>
        public void MoveToCell(int rowIndex, string columnText, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Move the mouse over the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            MoveToCellInternal(Location.X, Location.Y);
        }

        /// <summary>
        /// Moves the mouse cursor over the specified cell, scrolling it into view if need be
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <param name="locationInCell">The location in the cell to move the mouse to</param>
        public void MoveToCell(string rowText, int columnIndex, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Move the mouse over the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            MoveToCellInternal(Location.X, Location.Y);
        }

        /// <summary>
        /// Moves the mouse cursor over the specified cell, scrolling it into view if need be
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <param name="locationInCell">The location in the cell to move the mouse to</param>
        public void MoveToCell(int rowIndex, int columnIndex, CellClickLocation locationInCell)
        {
            GUI.Log("Move the mouse over the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            Point Location = GetLocationInCell(rowIndex, columnIndex, locationInCell);
            MoveToCellInternal(Location.X, Location.Y);
        }

        internal abstract void MoveToCellInternal(int X, int Y);

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void SingleClickCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void SingleClickCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void SingleClickCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Single " + button.ToString() + " click on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void SingleClickCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Single " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void SingleClickCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void SingleClickCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Single " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void SingleClickCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            GUI.Log("Single " + button.ToString() + " click on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Selects the specified cell by scrolling it into view and clicking on it
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void SingleClickCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            GUI.Log("Single " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            SingleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        internal abstract void SingleClickCellInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier);
        
        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void DoubleClickCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Double " + button.ToString() + " click on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void DoubleClickCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Double " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void DoubleClickCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Double " + button.ToString() + " click on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void DoubleClickCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Double " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void DoubleClickCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Double " + button.ToString() + " click on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void DoubleClickCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Double " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void DoubleClickCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            GUI.Log("Double " + button.ToString() + " click on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Double clicks the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void DoubleClickCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            GUI.Log("Double " + button.ToString() + " click while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            DoubleClickCellInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        internal abstract void DoubleClickCellInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier);

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseDownCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Press " + button.ToString() + " button on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseDownCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Press " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseDownCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Press " + button.ToString() + " button on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseDownCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Press " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseDownCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Press " + button.ToString() + " button on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseDownCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Press " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseDownCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            GUI.Log("Press " + button.ToString() + " button on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Presses the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseDownCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            GUI.Log("Press " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseDownInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        internal abstract void MouseDownInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier);

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseUpCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Release " + button.ToString() + " button on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseUpCell(string rowText, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Release " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnText, LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseUpCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Release " + button.ToString() + " button on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnText">The column text of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseUpCell(int rowIndex, string columnText, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int columnIndex = FindColumn(columnText);

            GUI.Log("Release " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnText, LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseUpCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Release " + button.ToString() + " button on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowText">The row text of the cell to select in the specified column</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseUpCell(string rowText, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            int rowIndex = FindRow(rowText, columnIndex);

            GUI.Log("Release " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowText + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        public void MouseUpCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell)
        {
            GUI.Log("Release " + button.ToString() + " button on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, MouseKeyModifier.None);
        }

        /// <summary>
        /// Releases the mouse button on the specified cell after scrolling it into view
        /// </summary>
        /// <param name="rowIndex">The row index of the cell to select</param>
        /// <param name="columnIndex">The column index of the cell to select</param>
        /// <param name="button">The button with which to click</param>
        /// <param name="locationInCell">The location in the cell to click</param>
        /// <param name="keyModifier">The key to press while clicking</param>
        public void MouseUpCell(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier)
        {
            GUI.Log("Release " + button.ToString() + " button while pressing key " + keyModifier.ToString() + " on the " + Identity.Description + " row " + rowIndex.ToString() + " column " + columnIndex.ToString(), LogItemType.Action);
            MouseUpInternal(rowIndex, columnIndex, button, locationInCell, keyModifier);
        }

        internal abstract void MouseUpInternal(int rowIndex, int columnIndex, MouseButton button, CellClickLocation locationInCell, MouseKeyModifier keyModifier);

        internal abstract Point GetLocationInCell(int rowIndex, int columnIndex, CellClickLocation locationInCell);

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowText">The row text of the cell</param>
        /// <param name="columnText">The column text of the cell</param>
        /// <returns></returns>
        internal Rectangle GetCellRectangle(string rowText, string columnText)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);
            return GetCellRectangle(rowIndex, columnIndex);
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnText">The column text of the cell</param>
        /// <returns></returns>
        internal Rectangle GetCellRectangle(int rowIndex, string columnText)
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
        internal Rectangle GetCellRectangle(string rowText, int columnIndex)
        {
            int RowNumber = FindRow(rowText, columnIndex);
            return GetCellRectangle(RowNumber, columnIndex);
        }

        /// <summary>
        /// Gets the bounding rectangle of the specified cell
        /// </summary>
        /// <param name="rowIndex">The row index of the cell</param>
        /// <param name="columnIndex">The column index of the cell</param>
        /// <returns>The bound rectangle</returns>
        internal Rectangle GetCellRectangle(int rowIndex, int columnIndex)
        {
            return GetCellRectangleInternal(rowIndex, columnIndex);
        }

        internal abstract Rectangle GetCellRectangleInternal(int rowIndex, int columnIndex);

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Text">The start column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Text">The end column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
        /// <param name="column2Text">The end column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns>A string containing the range of values</returns>
        public string GetCellRange(int row1Index, int column1Index, int row2Index, string column2Text)
        {
            return GetCellRange(row1Index, column1Index, row2Index, column2Text, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns a range of cell values column separated by \t and row separated by \r
        /// </summary>
        /// <param name="row1Index">The start row of the range</param>
        /// <param name="column1Text">The start column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
        /// <param name="column1Text">The start column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <param name="row2Index">The end row of the range</param>
        /// <param name="column2Text">The end column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
        /// <param name="column2Text">The end column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
        /// <param name="column1Text">The start column of the range delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
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
            return GetCellRangeInternal(row1Index, column1Index, row2Index, column2Index, property);
        }

        internal abstract string GetCellRangeInternal(int row1Index, int column1Index, int row2Index, int column2Index, CellProperty property);

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnText">Column of the cell delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns></returns>
        public string GetCell(string rowText, string columnText)
        {
            return GetCell(rowText, columnText, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnText">Column of the cell delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <returns>The cell value as a string</returns>
        /// <returns></returns>
        public string GetCell(int rowIndex, string columnText)
        {
            return GetCell(rowIndex, columnText, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <returns>The cell value as a string</returns>
        public string GetCell(string rowText, int columnIndex)
        {
            return GetCell(rowText, columnIndex, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell as a string
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <returns>The cell value as a string</returns>
        public string GetCell(int rowIndex, int columnIndex)
        {
            return GetCell(rowIndex, columnIndex, CellProperty.TextDisplay);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnText">Column of the cell delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCell(string rowText, string columnText, CellProperty property)
        {
            int columnIndex = FindColumn(columnText);
            int rowIndex = FindRow(rowText, columnIndex);

            return GetCell(rowIndex, columnIndex, property);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnText">Column of the cell delimited by -> (or the user defined GridDelimiter property) for example Order -> Id</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCell(int rowIndex, string columnText, CellProperty property)
        {
            int column = FindColumn(columnText);
            return GetCell(rowIndex, column, property);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowText">The row text of the cell in the specified column</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCell(string rowText, int columnIndex, CellProperty property)
        {
            int rowIndex = FindRow(rowText, columnIndex);
            return GetCell(rowIndex, columnIndex, property);
        }

        /// <summary>
        /// Returns the value of the specified cell property
        /// </summary>
        /// <param name="rowIndex">Row index of the cell</param>
        /// <param name="columnIndex">Column index of the cell</param>
        /// <param name="property">The property of the cell to get</param>
        /// <returns>The cell property</returns>
        public dynamic GetCell(int rowIndex, int columnIndex, CellProperty property)
        {
            return GetCellInternal(rowIndex, columnIndex, property);
        }

        internal abstract dynamic GetCellInternal(int rowIndex, int columnIndex, CellProperty property);

        /// <summary>
        /// Send the specified text to the control
        /// </summary>
        /// <param name="text">The text to send to the control</param>
        public void Type(string text)
        {
            SetFocus();
            SendKeys(text);
        }

        internal string[,] ConvertStringToString2dArray(string text)
        {
            string[,] array = null;
            string[] rows = text.Split('\r');
            for (int row = 0; row < rows.GetLength(0); row++)
            {
                string[] columns = rows[row].Split('\t');
                if (row == 0)
                {
                    array = new string[rows.GetLength(0), columns.GetLength(0)];
                }
                for (int column = 0; column < columns.GetLength(0); column++)
                {
                    array[row, column] = columns[column];
                }
            }

            return array;
        }
    }
}
