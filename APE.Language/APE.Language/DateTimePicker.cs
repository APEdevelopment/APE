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
using System.Threading;
using APE.Communication;
using System.Reflection;
using System.Drawing;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.DateTimePicker
    /// </summary>
    public sealed class GUIDateTimePicker : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIDateTimePicker(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Gets if the date time picker currently has a checkbox shown
        /// </summary>
        /// <returns>True if it has a checkbox shown otherwise false</returns>
        public bool HasCheckbox()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "ShowCheckBox", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool hasCheckbox = GUI.m_APE.GetValueFromMessage();
            return hasCheckbox;
        }

        /// <summary>
        /// Gets the checked state of the date time picker checkbox
        /// </summary>
        /// <returns>>True the checkbox is checked otherwise false</returns>
        public bool Checked()
        {
            if (HasCheckbox())
            {
                return GetState();
            }
            else
            {
                throw new Exception(Description + " does not have a checkbox shown");
            }
        }

        /// <summary>
        /// Checks the checkbox
        /// </summary>
        public void Check()
        {
            if (GetState() == true)
            {
                GUI.Log("Ensure " + Description + " checkbox is checked", LogItemType.Action);
            }
            else
            {
                GUI.Log("Check " + Description + " checkbox", LogItemType.Action);
                Rectangle location = GetCheckboxRectangle();
                SingleClickInternal(location.X + (location.Width / 2), location.Y + (location.Height / 2), MouseButton.Left, MouseKeyModifier.None);
                PollForState(true);
            }
        }

        /// <summary>
        /// Unchecks the checkbox
        /// </summary>
        public void Uncheck()
        {
            if (GetState() == false)
            {
                GUI.Log("Ensure " + Description + " checkbox is unchecked", LogItemType.Action);
            }
            else
            {
                GUI.Log("Uncheck " + Description + " checkbox", LogItemType.Action);
                Rectangle location = GetCheckboxRectangle();
                SingleClickInternal(location.X + (location.Width / 2), location.Y + (location.Height / 2), MouseButton.Left, MouseKeyModifier.None);
                PollForState(false);
            }
        }

        private Rectangle GetCheckboxRectangle()
        {
            GUI.m_APE.AddFirstMessageGetDateTimePickerCheckboxRectangle(Identity.Handle);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            int top = GUI.m_APE.GetValueFromMessage();
            int left = GUI.m_APE.GetValueFromMessage();
            int bottom = GUI.m_APE.GetValueFromMessage();
            int right = GUI.m_APE.GetValueFromMessage();

            return new Rectangle(left, top, (right - left), (bottom - top));
        }

        private Rectangle GetButtonRectangle()
        {
            GUI.m_APE.AddFirstMessageGetDateTimePickerButtonRectangle(Identity.Handle);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            int top = GUI.m_APE.GetValueFromMessage();
            int left = GUI.m_APE.GetValueFromMessage();
            int bottom = GUI.m_APE.GetValueFromMessage();
            int right = GUI.m_APE.GetValueFromMessage();

            return new Rectangle(left, top, (right - left), (bottom - top));
        }

        private bool GetState()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Checked", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isChecked = GUI.m_APE.GetValueFromMessage();
            return isChecked;
        }

        private void PollForState(bool state)
        {
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (GetState() == state)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    string stateText;
                    if (state)
                    {
                        stateText = "check";
                    }
                    else
                    {
                        stateText = "uncheck";
                    }
                    throw GUI.ApeException("Failed to " + stateText + " the " + Description + " checkbox");
                }

                Thread.Sleep(15);
            }
        }

        /// <summary>
        /// Sets the text portion of the datetimepicker to the specified text by sending keystrokes
        /// </summary>
        /// <param name="dateText">The text to set the text portion of the datetimepicker to</param>
        public void SetText(string dateText)
        {
            Stopwatch timer;
            string currentDateText;
            string customFormat = "";
            string formatedDateText = "";

            //Parse the text into a DateTime
            DateTime dateValue = DateTime.Parse(dateText);

            //Get the format of the date time picker
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Format", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string datePickerFormat = GUI.m_APE.GetValueFromMessage();

            switch (datePickerFormat)
            {
                case "Long":
                    formatedDateText = dateValue.ToLongDateString();
                    break;
                case "Short":
                    formatedDateText = dateValue.ToShortDateString();
                    break;
                case "Time":
                    formatedDateText = dateValue.ToLongTimeString();
                    break;
                case "Custom":
                    //Get the format of the date time picker
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CustomFormat", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    customFormat = GUI.m_APE.GetValueFromMessage();
                    formatedDateText = dateValue.ToString(customFormat);
                    break;
                default:
                    throw GUI.ApeException("Implement support for date time picker format " + datePickerFormat);
            }

            currentDateText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);
            if (currentDateText == formatedDateText)
            {
                GUI.Log("Ensure " + Identity.Description + " is set to " + dateText, LogItemType.Action);
                return;
            }

            Input.Block();
            try
            {
                //click on the left side of the control
                int x = 5;
                int y = 5;
                if (HasCheckbox())
                {
                    Rectangle location = GetCheckboxRectangle();
                    x += location.Right;
                }
                base.SingleClick(x, y, MouseButton.Left);

                char[] splitSeparator = { '/', ':', ' ' };
                string[] dateParts;
                switch (datePickerFormat)
                {
                    case "Long":
                    case "Short":
                        dateParts = dateValue.ToShortDateString().Split(splitSeparator);
                        break;
                    case "Time":
                        dateParts = dateValue.ToLongTimeString().Split(splitSeparator);
                        break;
                    case "Custom":
                        dateParts = dateValue.ToString(customFormat).Split(splitSeparator);
                        break;
                    default:
                        throw GUI.ApeException("Implement support for date time picker format " + datePickerFormat);
                }

                foreach (string part in dateParts)
                {
                    base.SendKeys(part);
                    base.SendKeys(":");
                }

                //wait for .Text to == text
                timer = Stopwatch.StartNew();
                do
                {
                    currentDateText = GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw GUI.ApeException("Failed to set the text of the TextBox");
                    }

                    Thread.Sleep(15);
                }
                while (currentDateText != formatedDateText);
                timer.Stop();
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
