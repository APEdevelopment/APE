using System;
using System.Text;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to improve performance of accessing flexgrid controls
    /// </summary>
    public partial class APEIPC
    {
        private delegate string GetCellRangeBackColorDelegate(dynamic grid, int row1, int column1, int row2, int column2);
        private GetCellRangeBackColorDelegate m_FlexgridGetCellRangeBackColourDelegater;

        /// <summary>
        /// Initialise the flexgrid helper delegates for use
        /// </summary>
        private void SetupFlexgridHelperDelegates()
        {
            m_FlexgridGetCellRangeBackColourDelegater = new GetCellRangeBackColorDelegate(FlexgridGetCellRangeBackColourInternal);
        }

        /// <summary>
        /// Iterates over every cell in the grid returning a /t /r separated string of the back colours of the cells
        /// any cells which have the default back colour of the grid return an empty string for that cell
        /// </summary>
        /// <param name="grid">The grid object</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        /// <returns>A /t /r delimited string of the back colours</returns>
        private string FlexgridGetCellRangeBackColourInternal(dynamic grid, int row1, int column1, int row2, int column2)
        {
            string gridBackColour = grid.BackColor.Name;
            string currentCellBackColour;
            StringBuilder rangeBackColour = new StringBuilder(10240);

            for (int row = row1; row <= row2; row++)
            {
                for (int column = column1; column <= column2; column++)
                {
                    currentCellBackColour = grid.GetCellRange(row, column).StyleDisplay.BackColor.Name;

                    if (gridBackColour == currentCellBackColour)
                    {
                        //nothing as its the default grid colour
                    }
                    else
                    {
                        rangeBackColour.Append(currentCellBackColour);
                    }

                    if (column < column2)
                    {
                        rangeBackColour.Append("\t");
                    }
                }

                if (row < row2)
                {
                    rangeBackColour.Append("\r");
                }
            }

            return rangeBackColour.ToString();
        }

        /// <summary>
        /// Calls into the AUT to iterate over every cell in the grid returning a /t /r separated string of the back colours
        /// of the cells any cells which have the default back colour of the grid return an empty string for that cell
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        /// <param name="row1">The start row of the range</param>
        /// <param name="column1">The start column of the range</param>
        /// <param name="row2">The end row of the range</param>
        /// <param name="column2">The end column of the range</param>
        unsafe public void AddQueryMessageFlexgridGetCellRangeBackColour(DataStores sourceStore, DataStores destinationStore, int row1, int column1, int row2, int column2)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the flexgrid before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;
            ptrMessage->Action = MessageAction.FlexgridGetCellRangeBackColour;

            Parameter row1Parameter = new Parameter(this, row1);
            Parameter column1Parameter = new Parameter(this, column1);
            Parameter row2Parameter = new Parameter(this, row2);
            Parameter column2Parameter = new Parameter(this, column2);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Gets the parameters from the message then calls FlexgridGetCellRangeBackColour storing the results in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void FlexgridGetCellRangeBackColour(Message* ptrMessage)
        {
            int row1 = GetParameterInt32(ptrMessage, 0);
            int column1 = GetParameterInt32(ptrMessage, 1);
            int row2 = GetParameterInt32(ptrMessage, 2);
            int column2 = GetParameterInt32(ptrMessage, 3);
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                object[] theParameters = { sourceObject, row1, column1, row2, column2 };
                destinationObject = ((WF.Control)tempStore0).Invoke(m_FlexgridGetCellRangeBackColourDelegater, theParameters);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }
    }
}
