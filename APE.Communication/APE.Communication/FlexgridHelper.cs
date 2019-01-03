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
using System.Diagnostics;
using System.Reflection;
using System.Text;
using System.Threading;
using WF = System.Windows.Forms;
using System.Drawing;
using Fasterflect;
using System.Linq;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to improve performance of accessing flexgrid controls
    /// </summary>
    public partial class APEIPC
    {
        public enum FlexgridGridType
        {
            ActiveX,
            Extended,
            Normal,
        }

        public enum CellProperty
        {
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
            /// Only for COM flexgrid to get displayed text of a range
            /// </summary>
            TextDisplay,
            /// <summary>
            /// The style of the font
            /// </summary>
            FontStyle,
        }

        string m_ColumnSeparator = "";
        string m_RowSeparator = "";

        /// <summary>
        /// Initialise the flexgrid helper delegates for use
        /// </summary>
        private void SetupFlexgridHelperDelegates()
        {
            m_FlexgridGetCellRangeBackColourNameDelegater = new GetCellRangeBackColourNameDelegate(FlexgridGetCellRangeBackColourNameInternal);
            m_FlexgridGetCellRangeForeColourNameDelegater = new GetCellRangeForeColourNameDelegate(FlexgridGetCellRangeForeColourNameInternal);
            m_FlexgridGetCellRangeDataTypeDelegater = new GetCellRangeDataTypeDelegate(FlexgridGetCellRangeDataTypeInternal);
            m_FlexgridGetCellRangeCheckBoxDelegater = new GetCellRangeCheckBoxDelegate(FlexgridGetCellRangeCheckBoxInternal);
            m_FlexgridGetCellRangeImageDelegater = new GetCellRangeImageDelegate(FlexgridGetCellRangeImageInternal);
            m_FlexgridGetCellRangeBackgroundImageDelegater = new GetCellRangeBackgroundImageDelegate(FlexgridGetCellRangeBackgroundImageInternal);
            m_FlexgridGetAllColumnsHiddenDelegater = new GetAllColumnsHiddenDelegate(FlexgridGetAllColumnsHiddenInternal);
            m_FlexgridGetAllColumnsWidthDelegater = new GetAllColumnsWidthDelegate(FlexgridGetAllColumnsWidthInternal);
            m_FlexgridGetAllRowsHiddenDelegater = new GetAllRowsHiddenDelegate(FlexgridGetAllRowsHiddenInternal);
            m_FlexgridGetAllRowsHeightDelegater = new GetAllRowsHeightDelegate(FlexgridGetAllRowsHeightInternal);
            m_FlexgridGetNodeCollapsedStateDelegater = new GetNodeCollapsedStateDelegate(FlexgridGetNodeCollapsedStateInternal);
            m_FlexgridGetCellRangeTextDisplayDelegater = new GetCellRangeTestDisplayDelegate(FlexgridGetCellRangeTextDisplayInternal);
            m_FlexgridGetCellRangeFontStyleDelegater = new GetCellRangeFontStyleDelegate(FlexgridGetCellRangeFontStyleInternal);
            m_FlexgridGetCellBackgroundImageDelegater = new GetCellBackgroundImageDelegate(FlexgridGetCellBackgroundImageInternal);
        }

        /// <summary>
        /// Determine what the grid currently uses to indicate column and row breaks
        /// </summary>
        /// <param name="clipSeparators">The current grid separators setting</param>
        private void DetermineClipSeparators(string clipSeparators)
        {
            if (string.IsNullOrEmpty(clipSeparators))
            {
                m_ColumnSeparator = "\t";
                m_RowSeparator = "\r";
            }
            else
            {
                m_ColumnSeparator = clipSeparators.Substring(0, 1);
                m_RowSeparator = clipSeparators.Substring(1, 1);
            }
        }

        //
        //  FlexgridGetCellRange
        //

        /// <summary>
        /// Calls into the AUT to iterate over every cell in the grid returning a \t \r separated string of the
        /// required property of the cells
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        unsafe public void AddQueryMessageFlexgridGetCellRange(DataStores sourceStore, DataStores destinationStore, int row1, int column1, int row2, int column2, CellProperty property, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            switch (property)
            {
                case CellProperty.BackColourName:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeBackColourName;
                    break;
                case CellProperty.ForeColourName:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeForeColourName;
                    break;
                case CellProperty.DataType:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeDataType;
                    break;
                case CellProperty.CheckBox:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeCheckBox;
                    break;
                case CellProperty.Image:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeImage;
                    break;
                case CellProperty.BackgroundImage:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeBackgroundImage;
                    break;
                case CellProperty.TextDisplay:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeTextDisplay;
                    break;
                case CellProperty.FontStyle:
                    ptrMessage->Action = MessageAction.FlexgridGetCellRangeFontStyle;
                    break;
                default:
                    throw new Exception("Implement support for getting cell property " + property.ToString());
            }

            Parameter row1Parameter = new Parameter(this, row1);
            Parameter column1Parameter = new Parameter(this, column1);
            Parameter row2Parameter = new Parameter(this, row2);
            Parameter column2Parameter = new Parameter(this, column2);
            Parameter gridTypeParameter = new Parameter(this, (int)gridType);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls the appropriate FlexgridGetCellRange method 
        /// on the correct thread storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetCellRange(Message* ptrMessage, CellProperty property)
        {
            int row1 = GetParameterInt32(ptrMessage, 0);
            int column1 = GetParameterInt32(ptrMessage, 1);
            int row2 = GetParameterInt32(ptrMessage, 2);
            int column2 = GetParameterInt32(ptrMessage, 3);
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 4);
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, row1, column1, row2, column2, gridType };
                switch (property)
                {
                    case CellProperty.BackColourName:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeBackColourNameDelegater, theParameters);
                        break;
                    case CellProperty.ForeColourName:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeForeColourNameDelegater, theParameters);
                        break;
                    case CellProperty.DataType:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeDataTypeDelegater, theParameters);
                        break;
                    case CellProperty.CheckBox:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeCheckBoxDelegater, theParameters);
                        break;
                    case CellProperty.Image:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeImageDelegater, theParameters);
                        break;
                    case CellProperty.BackgroundImage:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeBackgroundImageDelegater, theParameters);
                        break;
                    case CellProperty.TextDisplay:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeTextDisplayDelegater, theParameters);
                        break;
                    case CellProperty.FontStyle:
                        destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeFontStyleDelegater, theParameters);
                        break;
                    default:
                        throw new Exception("Implement support for getting cell property " + property.ToString());
                }
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        //
        //  FlexgridGetCellRangeBackColourName
        //

        private delegate string GetCellRangeBackColourNameDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeBackColourNameDelegate m_FlexgridGetCellRangeBackColourNameDelegater;
        private const int flexcpBackColor = 6;
        private Fasterflect.MethodInvoker m_MethodInvokerBackColourName = null;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of the back colours of the cells
        /// any cells which have the default back colour of the grid return an empty string for that cell
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the back colours</returns>
        private string FlexgridGetCellRangeBackColourNameInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            string gridBackColourName;
            if (gridType == FlexgridGridType.ActiveX)
            {
                gridBackColourName = null;
            }
            else
            {
                gridBackColourName = grid.BackColor.Name;
            }

            string currentCellBackColourName;
            StringBuilder rangeBackColourName = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    switch (gridType)
                    {
                        case FlexgridGridType.ActiveX:
                            int colourOle = grid.Cell(flexcpBackColor, row, column, row, column);
                            currentCellBackColourName = ColorTranslator.FromOle(colourOle).Name;
                            break;
                        case FlexgridGridType.Extended:
                            if (m_MethodInvokerBackColourName == null)
                            {
                                Type typeGrid = grid.GetType();
                                Type typeExtension = typeGrid.Assembly.GetTypes().SingleOrDefault(type => type.Name == "VSFlexGridExtensions");
                                MethodInfo miGetCellBackColor = typeExtension.GetMethod("getCellBackColor", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                                ParameterInfo[] pi = miGetCellBackColor.GetParameters();
                                Type[] typeParameters = new Type[pi.Length];
                                for (int parameter = 0; parameter < typeParameters.Length; parameter++)
                                {
                                    typeParameters[parameter] = pi[parameter].ParameterType;
                                }
                                m_MethodInvokerBackColourName = typeExtension.DelegateForCallMethod("getCellBackColor", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, typeParameters);
                            }
                            object[] parameters = new object[3];
                            parameters[0] = grid;
                            parameters[1] = row;
                            parameters[2] = column;
                            currentCellBackColourName = m_MethodInvokerBackColourName(grid, parameters).Name;
                            break;
                        default:
                            currentCellBackColourName = grid.GetCellRange(row, column).StyleDisplay.BackColor.Name;
                            break;
                    }

                    if (gridBackColourName == currentCellBackColourName)
                    {
                        //nothing as its the default grid colour
                    }
                    else
                    {
                        rangeBackColourName.Append(currentCellBackColourName);
                    }

                    if (column < column2)
                    {
                        rangeBackColourName.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeBackColourName.Append(m_RowSeparator);
                }
            }

            return rangeBackColourName.ToString();
        }

        //
        //  FlexgridGetCellRangeForeColourName
        //

        private delegate string GetCellRangeForeColourNameDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeForeColourNameDelegate m_FlexgridGetCellRangeForeColourNameDelegater;
        private const int flexcpForeColor = 7;
        private Fasterflect.MethodInvoker m_MethodInvokerForeColourName = null;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of the fore colours of the cells
        /// any cells which have the default fore colour of the grid return an empty string for that cell
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the fore colours</returns>
        private string FlexgridGetCellRangeForeColourNameInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);
            string gridForeColourName;
            if (gridType == FlexgridGridType.ActiveX)
            {
                gridForeColourName = null;
            }
            else
            {
                gridForeColourName = grid.ForeColor.Name;
            }

            string currentCellForeColourName;
            StringBuilder rangeForeColourName = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    switch (gridType)
                    {
                        case FlexgridGridType.ActiveX:
                            int colourOle = grid.Cell(flexcpForeColor, row, column, row, column);
                            currentCellForeColourName = ColorTranslator.FromOle(colourOle).Name;
                            break;
                        case FlexgridGridType.Extended:
                            if (m_MethodInvokerForeColourName == null)
                            {
                                Type typeGrid = grid.GetType();
                                Type typeExtension = typeGrid.Assembly.GetTypes().SingleOrDefault(type => type.Name == "VSFlexGridExtensions");
                                MethodInfo miGetCellBackColor = typeExtension.GetMethod("getCellForeColor", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                                ParameterInfo[] pi = miGetCellBackColor.GetParameters();
                                Type[] typeParameters = new Type[pi.Length];
                                for (int parameter = 0; parameter < typeParameters.Length; parameter++)
                                {
                                    typeParameters[parameter] = pi[parameter].ParameterType;
                                }
                                m_MethodInvokerForeColourName = typeExtension.DelegateForCallMethod("getCellForeColor", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, typeParameters);
                            }
                            object[] parameters = new object[3];
                            parameters[0] = grid;
                            parameters[1] = row;
                            parameters[2] = column;
                            currentCellForeColourName = m_MethodInvokerForeColourName(grid, parameters).Name;
                            break;
                        default:
                            currentCellForeColourName = grid.GetCellRange(row, column).StyleDisplay.ForeColor.Name;
                            break;
                    }

                    if (gridForeColourName == currentCellForeColourName)
                    {
                        //nothing as its the default grid colour
                    }
                    else
                    {
                        rangeForeColourName.Append(currentCellForeColourName);
                    }

                    if (column < column2)
                    {
                        rangeForeColourName.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeForeColourName.Append(m_RowSeparator);
                }
            }

            return rangeForeColourName.ToString();
        }

        //
        //  FlexgridGetCellRangeDataType
        //

        private delegate string GetCellRangeDataTypeDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeDataTypeDelegate m_FlexgridGetCellRangeDataTypeDelegater;
        private Fasterflect.MethodInvoker m_MethodInvokerDataType = null;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of the data type of the cells
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the data type</returns>
        private string FlexgridGetCellRangeDataTypeInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            Type currentCellDataType = null;
            StringBuilder rangeDataType = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    switch (gridType)
                    {
                        case FlexgridGridType.ActiveX:
                            throw new Exception("Not supported");
                        case FlexgridGridType.Extended:
                            if (m_MethodInvokerDataType == null)
                            {
                                Type typeGrid = grid.GetType();
                                Type typeExtension = typeGrid.Assembly.GetTypes().SingleOrDefault(type => type.Name == "VSFlexGridExtensions");
                                MethodInfo miGetCellBackColor = typeExtension.GetMethod("getCellDataType", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                                ParameterInfo[] pi = miGetCellBackColor.GetParameters();
                                Type[] typeParameters = new Type[pi.Length];
                                for (int parameter = 0; parameter < typeParameters.Length; parameter++)
                                {
                                    typeParameters[parameter] = pi[parameter].ParameterType;
                                }
                                m_MethodInvokerDataType = typeExtension.DelegateForCallMethod("getCellDataType", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, typeParameters);
                            }
                            object[] parameters = new object[3];
                            parameters[0] = grid;
                            parameters[1] = row;
                            parameters[2] = column;
                            currentCellDataType = m_MethodInvokerDataType(grid, parameters);
                            break;
                        default:
                            currentCellDataType = grid.GetCellRange(row, column).StyleDisplay.DataType;
                            break;
                    }

                    rangeDataType.Append(currentCellDataType.Namespace);
                    rangeDataType.Append(".");
                    rangeDataType.Append(currentCellDataType.Name);

                    if (column < column2)
                    {
                        rangeDataType.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeDataType.Append(m_RowSeparator);
                }
            }

            return rangeDataType.ToString();
        }

        //
        //  FlexgridGetCellRangeCheckBox
        //

        private delegate string GetCellRangeCheckBoxDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeCheckBoxDelegate m_FlexgridGetCellRangeCheckBoxDelegater;
        private const int flexcpChecked = 5;
        private const int flexcpData = 20;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of the checkbox state of the cells
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the data type</returns>
        private string FlexgridGetCellRangeCheckBoxInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            string currentCellCheckBox;
            StringBuilder rangeDataType = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    if (gridType == FlexgridGridType.ActiveX)
                    {
                        currentCellCheckBox = "None";
                        int gridCheckbox = grid.Cell(flexcpChecked, row, column, row, column);

                        if (gridCheckbox == 0)
                        {
                            dynamic cellCheckedStatus = grid.Cell(flexcpData, row, column, row, column);
                            string interfaceName = GetObjectInterfaceName(cellCheckedStatus);
                            if (interfaceName == "_CellCheckedStatus")
                            {
                                bool state = cellCheckedStatus.bCurrent;
                                if (state)
                                {
                                    currentCellCheckBox = "Checked";
                                }
                                else
                                {
                                    currentCellCheckBox = "Unchecked";
                                }
                            }
                        }
                        else
                        {
                            if (gridCheckbox == 1)
                            {
                                currentCellCheckBox = "Checked";
                            }
                            else if (gridCheckbox == 2)
                            {
                                currentCellCheckBox = "Unchecked";
                            }
                        }
                    }
                    else
                    {
                        currentCellCheckBox = grid.GetCellRange(row, column).Checkbox.ToString();
                    }

                    rangeDataType.Append(currentCellCheckBox);

                    if (column < column2)
                    {
                        rangeDataType.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeDataType.Append(m_RowSeparator);
                }
            }

            return rangeDataType.ToString();
        }

        //
        //  FlexgridGetCellRangeImage
        //

        private delegate string GetCellRangeImageDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeImageDelegate m_FlexgridGetCellRangeImageDelegater;
        private const int flexcpPicture = 3;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of whether the cell contains an image
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the data type</returns>
        private string FlexgridGetCellRangeImageInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            StringBuilder rangeDataType = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    object image;
                    if (gridType == FlexgridGridType.ActiveX)
                    {
                        image = grid.Cell(flexcpPicture, row, column, row, column);
                    }
                    else
                    {
                        image = grid.GetCellRange(row, column).Image;
                    }
                    if (image != null)
                    {
                        rangeDataType.Append("Image");
                    }

                    if (column < column2)
                    {
                        rangeDataType.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeDataType.Append(m_RowSeparator);
                }
            }

            return rangeDataType.ToString();
        }

        //
        //  FlexgridGetCellBackgroundImage
        //

        private delegate Image GetCellBackgroundImageDelegate(dynamic grid, int row, int column, FlexgridGridType gridType);
        private GetCellBackgroundImageDelegate m_FlexgridGetCellBackgroundImageDelegater;

        /// <summary>
        /// Calls into the AUT to iterate over every cell in the grid returning a \t \r separated string of the
        /// required property of the cells
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="row">The row index to get</param>
        /// <param name="column">The column index to get</param>
        /// <param name="gridType">The type of flexgrid grid</param>
        unsafe public void AddQueryMessageFlexgridGetCellBackgroundImage(DataStores sourceStore, DataStores destinationStore, int row, int column, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            ptrMessage->Action = MessageAction.FlexgridGetCellBackgroundImage;

            Parameter row1Parameter = new Parameter(this, row);
            Parameter column1Parameter = new Parameter(this, column);
            Parameter gridTypeParameter = new Parameter(this, (int)gridType);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls FlexgridGetCellRangeBackgroundImageInternal method 
        /// on the correct thread storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetCellBackgroundImage(Message* ptrMessage)
        {
            int row = GetParameterInt32(ptrMessage, 0);
            int column = GetParameterInt32(ptrMessage, 1);
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 2);
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, row, column, gridType };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellBackgroundImageDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Get the background image of the specified cell
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row">The start row of the range</param>
        /// <param name="column">The start column of the range</param>
        /// <returns>The image</returns>
        private Image FlexgridGetCellBackgroundImageInternal(dynamic grid, int row, int column, FlexgridGridType gridType)
        {            
            Image image = null;
            switch (gridType)
            {
                case FlexgridGridType.ActiveX:
                    image = grid.Cell(flexcpPicture, row, column, row, column);
                    break;
                case FlexgridGridType.Extended:
                    if (m_MethodInvokerBackgroundImage == null)
                    {
                        Type typeGrid = grid.GetType();
                        Type typeExtension = typeGrid.Assembly.GetTypes().SingleOrDefault(type => type.Name == "VSFlexGridExtensions");
                        MethodInfo miGetCellBackColor = typeExtension.GetMethod("getCellBackgroundImage", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                        ParameterInfo[] pi = miGetCellBackColor.GetParameters();
                        Type[] typeParameters = new Type[pi.Length];
                        for (int parameter = 0; parameter < typeParameters.Length; parameter++)
                        {
                            typeParameters[parameter] = pi[parameter].ParameterType;
                        }
                        m_MethodInvokerBackgroundImage = typeExtension.DelegateForCallMethod("getCellBackgroundImage", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, typeParameters);
                    }
                    object[] parameters = new object[3];
                    parameters[0] = grid;
                    parameters[1] = row;
                    parameters[2] = column;
                    image = m_MethodInvokerBackgroundImage(grid, parameters);
                    break;
                default:
                    image = grid.GetCellRange(row, column).StyleDisplay.BackgroundImage;
                    break;
            }
            return image;
        }

        //
        //  FlexgridGetCellRangeBackgroundImage
        //

        private delegate string GetCellRangeBackgroundImageDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeBackgroundImageDelegate m_FlexgridGetCellRangeBackgroundImageDelegater;
        private Fasterflect.MethodInvoker m_MethodInvokerBackgroundImage = null;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of whether the cell contains an 
        /// background image
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of BackgroundImage or empty string</returns>
        private string FlexgridGetCellRangeBackgroundImageInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            StringBuilder rangeDataType = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    object image;

                    switch (gridType)
                    {
                        case FlexgridGridType.ActiveX:
                            image = grid.Cell(flexcpPicture, row, column, row, column);
                            break;
                        case FlexgridGridType.Extended:
                            if (m_MethodInvokerBackgroundImage == null)
                            {
                                Type typeGrid = grid.GetType();
                                Type typeExtension = typeGrid.Assembly.GetTypes().SingleOrDefault(type => type.Name == "VSFlexGridExtensions");
                                MethodInfo miGetCellBackColor = typeExtension.GetMethod("getCellBackgroundImage", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                                ParameterInfo[] pi = miGetCellBackColor.GetParameters();
                                Type[] typeParameters = new Type[pi.Length];
                                for (int parameter = 0; parameter < typeParameters.Length; parameter++)
                                {
                                    typeParameters[parameter] = pi[parameter].ParameterType;
                                }
                                m_MethodInvokerBackgroundImage = typeExtension.DelegateForCallMethod("getCellBackgroundImage", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, typeParameters);
                            }
                            object[] parameters = new object[3];
                            parameters[0] = grid;
                            parameters[1] = row;
                            parameters[2] = column;
                            image = m_MethodInvokerBackgroundImage(grid, parameters);
                            break;
                        default:
                            image = grid.GetCellRange(row, column).StyleDisplay.BackgroundImage;
                            break;
                    }

                    if (image != null)
                    {
                        rangeDataType.Append("BackgroundImage");
                    }

                    if (column < column2)
                    {
                        rangeDataType.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeDataType.Append(m_RowSeparator);
                }
            }

            return rangeDataType.ToString();
        }

        //
        //  FlexgridGetCellRangeFontStyle
        //

        private delegate string GetCellRangeFontStyleDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeFontStyleDelegate m_FlexgridGetCellRangeFontStyleDelegater;
        private const int flexcpFontBold = 13;
        private const int flexcpFontItalic = 14;
        private const int flexcpFontUnderline = 15;
        private const int flexcpFontStrikethru = 16;
        private Fasterflect.MethodInvoker m_MethodInvokerFont = null;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of font style
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the font style</returns>
        private string FlexgridGetCellRangeFontStyleInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            StringBuilder rangeTextDisplay = new StringBuilder(10240);
            StringBuilder fontStyleStringBuilder = new StringBuilder();

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    fontStyleStringBuilder.Clear();
                    switch (gridType)
                    {
                        case FlexgridGridType.ActiveX:
                            bool bold = grid.Cell(flexcpFontBold, row, column, row, column);
                            bool italic = grid.Cell(flexcpFontItalic, row, column, row, column);
                            bool underline = grid.Cell(flexcpFontUnderline, row, column, row, column);
                            bool strikeout = grid.Cell(flexcpFontStrikethru, row, column, row, column);

                            if (bold)
                            {
                                fontStyleStringBuilder.Append("Bold,");
                            }
                            if (italic)
                            {
                                fontStyleStringBuilder.Append("Italic,");
                            }
                            if (underline)
                            {
                                fontStyleStringBuilder.Append("Underline,");
                            }
                            if (strikeout)
                            {
                                fontStyleStringBuilder.Append("Strikeout,");
                            }
                            if (fontStyleStringBuilder.Length == 0)
                            {
                                fontStyleStringBuilder.Append("Regular,");
                            }
                            fontStyleStringBuilder.Length--;
                            break;
                        case FlexgridGridType.Extended:
                            if (m_MethodInvokerFont == null)
                            {
                                Type typeGrid = grid.GetType();
                                Type typeExtension = typeGrid.Assembly.GetTypes().SingleOrDefault(type => type.Name == "VSFlexGridExtensions");
                                MethodInfo miGetCellBackColor = typeExtension.GetMethod("getCellFont", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic);
                                ParameterInfo[] pi = miGetCellBackColor.GetParameters();
                                Type[] typeParameters = new Type[pi.Length];
                                for (int parameter = 0; parameter < typeParameters.Length; parameter++)
                                {
                                    typeParameters[parameter] = pi[parameter].ParameterType;
                                }
                                m_MethodInvokerFont = typeExtension.DelegateForCallMethod("getCellFont", BindingFlags.Static | BindingFlags.Public | BindingFlags.NonPublic, typeParameters);
                            }
                            object[] parameters = new object[3];
                            parameters[0] = grid;
                            parameters[1] = row;
                            parameters[2] = column;
                            fontStyleStringBuilder.Append(m_MethodInvokerFont(grid, parameters).Style.ToString());
                            break;
                        default:
                            fontStyleStringBuilder.Append(grid.GetCellRange(row, column).StyleDisplay.Font.Style.ToString());
                            break;
                    }

                    rangeTextDisplay.Append(fontStyleStringBuilder.ToString());

                    if (column < column2)
                    {
                        rangeTextDisplay.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeTextDisplay.Append(m_RowSeparator);
                }
            }

            return rangeTextDisplay.ToString();
        }

        //
        //  FlexgridGetCellRangeTextDisplay
        //

        private delegate string GetCellRangeTestDisplayDelegate(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType);
        private GetCellRangeTestDisplayDelegate m_FlexgridGetCellRangeTextDisplayDelegater;
        private const int flexcpTextDisplay = 19;

        /// <summary>
        /// Iterates over every cell in the grid returning a \t \r separated string of displayed text
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A \t \r delimited string of the displayed text</returns>
        private string FlexgridGetCellRangeTextDisplayInternal(dynamic grid, int row1, int column1, int row2, int column2, FlexgridGridType gridType)
        {
            DetermineClipSeparators(grid.ClipSeparators);

            StringBuilder rangeTextDisplay = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    string textDisplay;
                    switch (gridType)
                    {
                        case FlexgridGridType.ActiveX:
                            textDisplay = grid.Cell(flexcpTextDisplay, row, column, row, column);
                            break;
                        default:
                            throw new Exception("Not supported");
                    }

                    if (string.IsNullOrEmpty(textDisplay))
                    {
                        rangeTextDisplay.Append("");
                    }
                    else
                    {
                        rangeTextDisplay.Append(textDisplay);
                    }

                    if (column < column2)
                    {
                        rangeTextDisplay.Append(m_ColumnSeparator);
                    }
                }

                if (row < row2)
                {
                    rangeTextDisplay.Append(m_RowSeparator);
                }
            }

            return rangeTextDisplay.ToString();
        }

        //
        //  FlexgridGetAllColumnsHidden
        //

        private delegate string GetAllColumnsHiddenDelegate(dynamic grid, FlexgridGridType gridType);
        private GetAllColumnsHiddenDelegate m_FlexgridGetAllColumnsHiddenDelegater;

        /// <summary>
        /// Calls into the AUT to iterate over every column in the grid returning a comma separated string of whether
        /// the column is hidden (True) or visible (False)
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="gridType">The type of flexgrid</param>
        unsafe public void AddQueryMessageFlexgridGetAllColumnsHidden(DataStores sourceStore, DataStores destinationStore, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            Parameter gridTypeParam = new Parameter(this, (int)gridType);

            ptrMessage->Action = MessageAction.FlexgridGetAllColumnsHidden;
            
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Calls the FlexgridGetAllColumnsHiddenInternal method on the correct thread storing the results
        /// in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetAllColumnsHidden(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            // p1  = bool
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 0);

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, gridType };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetAllColumnsHiddenDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Iterates over every column in the grid returning a comma separated string of whether the column is
        /// hidden (True) or visible (False)
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="gridType">The type of flexgrid</param>
        /// <returns>A comma delimited string of whether the column is hidden</returns>
        private string FlexgridGetAllColumnsHiddenInternal(dynamic grid, FlexgridGridType gridType)
        {
            StringBuilder columnsHiddenState = new StringBuilder(10240);
            if (gridType == FlexgridGridType.ActiveX)
            {
                int columns = grid.Cols;
                for (int column = 0; column < columns; column++)
                {
                    bool hidden = grid.ColHidden[column];
                    if (column > 0)
                    {
                        columnsHiddenState.Append(",");
                    }
                    if (hidden)
                    {
                        columnsHiddenState.Append("True");
                    }
                    else
                    {
                        columnsHiddenState.Append("False");
                    }
                }
            }
            else
            {
                int columns = grid.Cols.Count;
                for (int column = 0; column < columns; column++)
                {
                    bool visible = grid.Cols[column].Visible;
                    if (column > 0)
                    {
                        columnsHiddenState.Append(",");
                    }
                    if (visible)
                    {
                        columnsHiddenState.Append("False");
                    }
                    else
                    {
                        columnsHiddenState.Append("True");
                    }
                }
            }
            return columnsHiddenState.ToString();
        }

        //
        //  FlexgridGetAllColumnsWidth
        //

        private delegate string GetAllColumnsWidthDelegate(dynamic grid, FlexgridGridType gridType);
        private GetAllColumnsWidthDelegate m_FlexgridGetAllColumnsWidthDelegater;

        /// <summary>
        /// Calls into the AUT to iterate over every column in the grid returning a comma separated string of
        /// the columns width
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="gridType">The type of flexgrid</param>
        unsafe public void AddQueryMessageFlexgridGetAllColumnsWidth(DataStores sourceStore, DataStores destinationStore, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            Parameter gridTypeParam = new Parameter(this, (int)gridType);

            ptrMessage->Action = MessageAction.FlexgridGetAllColumnsWidth;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Calls the FlexgridGetAllColumnsWidthInternal method on the correct thread storing the results
        /// in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetAllColumnsWidth(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            // p1  = bool
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 0);

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, gridType };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetAllColumnsWidthDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Iterates over every column in the grid returning a comma separated string of the column width
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="gridType">The type of flexgrid</param>
        /// <returns>A comma delimited string of the column width</returns>
        private string FlexgridGetAllColumnsWidthInternal(dynamic grid, FlexgridGridType gridType)
        {
            StringBuilder columnsWidth = new StringBuilder(10240);
            if (gridType == FlexgridGridType.ActiveX)
            {
                int columns = grid.Cols;
                for (int column = 0; column < columns; column++)
                {
                    int width = grid.ColWidth[column];
                    if (column > 0)
                    {
                        columnsWidth.Append(",");
                    }
                    columnsWidth.Append(width.ToString());
                }
            }
            else
            {
                int columns = grid.Cols.Count;
                for (int column = 0; column < columns; column++)
                {
                    int width = grid.Cols[column].WidthDisplay;
                    if (column > 0)
                    {
                        columnsWidth.Append(",");
                    }
                    columnsWidth.Append(width.ToString());
                }
            }
            return columnsWidth.ToString();
        }

        //
        //  FlexgridGetAllRowsHidden
        //

        private delegate string GetAllRowsHiddenDelegate(dynamic grid, FlexgridGridType gridType);
        private GetAllRowsHiddenDelegate m_FlexgridGetAllRowsHiddenDelegater;

        /// <summary>
        /// Calls into the AUT to iterate over every row in the grid returning a comma separated string of whether
        /// the row is hidden (True) or visible (False)
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="gridType">The type of flexgrid</param>
        unsafe public void AddQueryMessageFlexgridGetAllRowsHidden(DataStores sourceStore, DataStores destinationStore, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            Parameter gridTypeParam = new Parameter(this, (int)gridType);

            ptrMessage->Action = MessageAction.FlexgridGetAllRowsHidden;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Calls the FlexgridGetAllRowsHiddenInternal method on the correct thread storing the results
        /// in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetAllRowsHidden(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            // p1  = bool
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 0);

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, gridType };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetAllRowsHiddenDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Iterates over every row in the grid returning a comma separated string of whether the row is
        /// hidden (True) or visible (False)
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="gridType">The type of flexgrid</param>
        /// <returns>A comma delimited string of whether the row is hidden</returns>
        private string FlexgridGetAllRowsHiddenInternal(dynamic grid, FlexgridGridType gridType)
        {
            StringBuilder rowsHiddenState = new StringBuilder(10240);
            if (gridType == FlexgridGridType.ActiveX)
            {
                int rows = grid.Rows;
                for (int row = 0; row < rows; row++)
                {
                    bool hidden = grid.RowHidden[row];
                    if (row > 0)
                    {
                        rowsHiddenState.Append(",");
                    }
                    if (hidden)
                    {
                        rowsHiddenState.Append("True");
                    }
                    else
                    {
                        rowsHiddenState.Append("False");
                    }
                }
            }
            else
            {
                int rows = grid.Rows.Count;
                for (int row = 0; row < rows; row++)
                {
                    bool visible = grid.Rows[row].Visible;
                    if (row > 0)
                    {
                        rowsHiddenState.Append(",");
                    }
                    if (visible)
                    {
                        rowsHiddenState.Append("False");
                    }
                    else
                    {
                        rowsHiddenState.Append("True");
                    }
                }
            }
            return rowsHiddenState.ToString();
        }

        //
        //  FlexgridGetAllRowsHeight
        //

        private delegate string GetAllRowsHeightDelegate(dynamic grid, FlexgridGridType gridType);
        private GetAllRowsHeightDelegate m_FlexgridGetAllRowsHeightDelegater;

        /// <summary>
        /// Calls into the AUT to iterate over every row in the grid returning a comma separated string of
        /// the rows height
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="gridType">The type of flexgrid</param>
        unsafe public void AddQueryMessageFlexgridGetAllRowsHeight(DataStores sourceStore, DataStores destinationStore, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            Parameter gridTypeParam = new Parameter(this, (int)gridType);

            ptrMessage->Action = MessageAction.FlexgridGetAllRowsHeight;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Calls the FlexgridGetAllRowsHeightInternal method on the correct thread storing the results
        /// in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetAllRowsHeight(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            // p1  = bool
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 0);

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, gridType };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetAllRowsHeightDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Iterates over every row in the grid returning a comma separated string of the row height
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="gridType">The type of flexgrid</param>
        /// <returns>A comma delimited string of the row height</returns>
        private string FlexgridGetAllRowsHeightInternal(dynamic grid, FlexgridGridType gridType)
        {
            StringBuilder rowHeight = new StringBuilder(10240);
            if (gridType == FlexgridGridType.ActiveX)
            {
                int rows = grid.Rows;
                for (int row = 0; row < rows; row++)
                {
                    int height = grid.RowHeight[row];
                    if (row > 0)
                    {
                        rowHeight.Append(",");
                    }
                    rowHeight.Append(height.ToString());
                }
            }
            else
            {
                int rows = grid.Rows.Count;
                for (int row = 0; row < rows; row++)
                {
                    int height = grid.Rows[row].HeightDisplay;
                    if (row > 0)
                    {
                        rowHeight.Append(",");
                    }
                    rowHeight.Append(height.ToString());
                }
            }
            return rowHeight.ToString();
        }

        //
        //  AddFlexgridCellChangedHandler
        //

        private bool m_FlexgridCellChanged = false;
        private Delegate m_FlexgridCellChangedHandler = null;
        private EventInfo m_FlexgridCellChangedEventInfo = null;
        private dynamic m_FlexgridControl = null;

        /// <summary>
        /// Calls into the AUT to add a cell changed handler to the specific control
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the control object</param>
        unsafe public void AddQueryMessageAddFlexgridCellChangedHandler(DataStores sourceStore)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->Action = MessageAction.AddFlexgridCellChangedHandler;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Adds a cell changed handler to the control
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void AddFlexgridCellChangedHandler(Message* ptrMessage)
        {
            m_FlexgridControl = GetObjectFromDatastore(ptrMessage->SourceStore);

            m_FlexgridCellChanged = false;

            if (m_FlexgridControl != null)
            {
                if (m_FlexgridCellChangedHandler == null)
                {
                    MethodInfo flexgridCellChangedMethodInfo = this.GetType().GetMethod("FlexgridCellChangedHandler", BindingFlags.NonPublic | BindingFlags.Instance);
                    m_FlexgridCellChangedEventInfo = m_FlexgridControl.GetType().GetEvent("CellChanged");
                    m_FlexgridCellChangedHandler = Delegate.CreateDelegate(m_FlexgridCellChangedEventInfo.EventHandlerType, this, flexgridCellChangedMethodInfo);
                }

                m_FlexgridCellChangedEventInfo.AddEventHandler(m_FlexgridControl, m_FlexgridCellChangedHandler);
            }

            CleanUpMessage(ptrMessage);
        }

        private void FlexgridCellChangedHandler(object sender, EventArgs e)
        {
            m_FlexgridCellChanged = true;
        }

        //
        //  WaitForAndRemoveCellChangedHandler
        //

        /// <summary>
        /// Calls into the AUT to wait for then remove the cell changed handler
        /// </summary>
        unsafe public void AddFirstMessageWaitForAndRemoveFlexgridCellChangedHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.WaitForAndRemoveFlexgridCellChangedHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Waits for the handler to set the m_FlexgridCellChanged variable to true then removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void WaitForAndRemoveFlexgridCellChangedHandler(Message* ptrMessage)
        {
            if (m_FlexgridControl != null)
            {
                try
                {
                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (m_FlexgridCellChanged || m_FlexgridControl.Disposing || m_FlexgridControl.IsDisposed)
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > m_TimeOut)
                        {
                            throw new Exception("Failed to find flexgrid cell changed event");
                        }

                        Thread.Sleep(15);
                    }
                }
                finally
                {
                    m_FlexgridCellChangedEventInfo.RemoveEventHandler(m_FlexgridControl, m_FlexgridCellChangedHandler);
                    m_FlexgridControl = null;
                }
            }

            CleanUpMessage(ptrMessage);
        }

        //
        //  AddFlexgridAfterRowColChangedHandler
        //

        private bool m_FlexgridAfterRowColChange = false;
        private Delegate m_FlexgridAfterRowColChangeHandler = null;
        private EventInfo m_FlexgridAfterRowColChangeEventInfo = null;

        /// <summary>
        /// Calls into the AUT to add a after row col change handler to the specific control
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the control object</param>
        unsafe public void AddQueryMessageAddFlexgridAfterRowColChangeHandler(DataStores sourceStore)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->Action = MessageAction.AddFlexgridAfterRowColChangeHandler;
            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Adds a after row col change handler to the control
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void AddFlexgridAfterRowColChangeHandler(Message* ptrMessage)
        {
            m_FlexgridControl = GetObjectFromDatastore(ptrMessage->SourceStore);

            m_FlexgridAfterRowColChange = false;

            if (m_FlexgridControl != null)
            {
                if (m_FlexgridAfterRowColChangeHandler == null)
                {
                    MethodInfo flexgridAfterRowColChangeMethodInfo = this.GetType().GetMethod("FlexgridAfterRowColChangeHandler", BindingFlags.NonPublic | BindingFlags.Instance);
                    m_FlexgridAfterRowColChangeEventInfo = m_FlexgridControl.GetType().GetEvent("AfterRowColChange");
                    m_FlexgridAfterRowColChangeHandler = Delegate.CreateDelegate(m_FlexgridAfterRowColChangeEventInfo.EventHandlerType, this, flexgridAfterRowColChangeMethodInfo);
                }

                m_FlexgridAfterRowColChangeEventInfo.AddEventHandler(m_FlexgridControl, m_FlexgridAfterRowColChangeHandler);
            }

            CleanUpMessage(ptrMessage);
        }

        private void FlexgridAfterRowColChangeHandler(object sender, EventArgs e)
        {
            m_FlexgridAfterRowColChange = true;
        }

        //
        //  WaitForAndRemoveAfterRowColChangeHandler
        //

        /// <summary>
        /// Calls into the AUT to wait for then remove the after row col change handler
        /// </summary>
        unsafe public void AddFirstMessageWaitForAndRemoveFlexgridAfterRowColChangeHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.WaitForAndRemoveFlexgridAfterRowColChangeHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Waits for the handler to set the m_FlexgridAfterRowColChange variable to true then removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void WaitForAndRemoveFlexgridAfterRowColChangeHandler(Message* ptrMessage)
        {
            if (m_FlexgridControl != null)
            {
                try
                {
                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (m_FlexgridAfterRowColChange || m_FlexgridControl.Disposing || m_FlexgridControl.IsDisposed)
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > m_TimeOut)
                        {
                            throw new Exception("Failed to find flexgrid after row col change event");
                        }

                        Thread.Sleep(15);
                    }
                }
                finally
                {
                    m_FlexgridAfterRowColChangeEventInfo.RemoveEventHandler(m_FlexgridControl, m_FlexgridAfterRowColChangeHandler);
                    m_FlexgridControl = null;
                }
            }

            CleanUpMessage(ptrMessage);
        }

        //
        //  RemoveAfterRowColChangeHandler
        //

        /// <summary>
        /// Calls into the AUT to remove the after row col change handler
        /// </summary>
        unsafe public void AddFirstMessageRemoveFlexgridAfterRowColChangeHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.RemoveFlexgridAfterRowColChangeHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void RemoveFlexgridAfterRowColChangeHandler(Message* ptrMessage)
        {
            if (m_FlexgridControl != null)
            {
                m_FlexgridAfterRowColChangeEventInfo.RemoveEventHandler(m_FlexgridControl, m_FlexgridAfterRowColChangeHandler);
                m_FlexgridControl = null;
            }

            CleanUpMessage(ptrMessage);
        }

        //
        //  FlexgridGetNodeCollapsedState
        //

        private delegate string GetNodeCollapsedStateDelegate(dynamic grid, FlexgridGridType gridType);
        private GetNodeCollapsedStateDelegate m_FlexgridGetNodeCollapsedStateDelegater;

        /// <summary>
        /// Calls into the AUT to iterates over every row in the grid returning a comma separated string of 
        /// 'whether the node is a row' | 'is collapsed / visible' | 'has children'
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="gridType">The type of flexgrid</param>
        unsafe public void AddQueryMessageFlexgridGetNodeCollapsedState(DataStores sourceStore, DataStores destinationStore, FlexgridGridType gridType)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            Parameter gridTypeParam = new Parameter(this, (int)gridType);

            ptrMessage->Action = MessageAction.FlexgridGetNodeCollapsedState;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Calls the FlexgridGetNodeCollapsedStateInternal method on the correct thread storing the results
        /// in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetNodeCollapsedState(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            // p1  = bool
            FlexgridGridType gridType = (FlexgridGridType)GetParameterInt32(ptrMessage, 0);

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, gridType };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetNodeCollapsedStateDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

        /// <summary>
        /// Iterates over every row in the grid returning a comma separated string of 'whether the node is a row' | 'is collapsed / visible' | 'has children'
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="gridType">The type of flexgrid</param>
        /// <returns>A comma delimited string</returns>
        private string FlexgridGetNodeCollapsedStateInternal(dynamic grid, FlexgridGridType gridType)
        {
            StringBuilder nodeCollapsedState = new StringBuilder(10240);
            if (gridType == FlexgridGridType.ActiveX)
            {
                int rows = grid.Rows;
                for (int row = 0; row < rows; row++)
                {
                    if (row > 0)
                    {
                        nodeCollapsedState.Append(",");
                    }
                    try
                    {
                        dynamic node = grid.GetNode(row);
                        int nodeRow = node.Row;
                        if (nodeRow == row)
                        {
                            int children = node.Children;
                            bool collapsed = !node.Expanded;
                            if (collapsed)
                            {
                                nodeCollapsedState.Append("True|True");
                            }
                            else
                            {
                                nodeCollapsedState.Append("True|False");
                            }
                            if (children == 0)
                            {
                                nodeCollapsedState.Append("|False");
                            }
                            else
                            {
                                nodeCollapsedState.Append("|True");
                            }
                        }
                        else
                        {
                            bool rowHidden = grid.RowHidden(row);
                            if (rowHidden)
                            {
                                nodeCollapsedState.Append("False|True|False");
                            }
                            else
                            {
                                nodeCollapsedState.Append("False|False|False");
                            }
                        }
                    }
                    catch
                    {
                        nodeCollapsedState.Append("NA|NA|NA");
                    }
                }
            }
            else
            {
                throw new Exception("Not implemented");
            }
            return nodeCollapsedState.ToString();
        }
    }
}
