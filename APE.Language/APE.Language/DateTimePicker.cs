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
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Threading;
using NM = APE.Native.NativeMethods;
using APE.Communication;
using System.Reflection;

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
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Format", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
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
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "CustomFormat", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    customFormat = GUI.m_APE.GetValueFromMessage();
                    formatedDateText = dateValue.ToString(customFormat);
                    break;
                default:
                    throw new Exception("Implement support for date time picker format " + datePickerFormat);
            }

            currentDateText = GUI.m_APE.GetWindowText(Identity.Handle);
            if (currentDateText == formatedDateText)
            {
                GUI.Log("Ensure " + m_DescriptionOfControl + " is set to " + dateText, LogItemTypeEnum.Action);
                return;
            }

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                //click on the left side of the control
                base.MouseSingleClick(5, 5, MouseButton.Left);

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
                        throw new Exception("Implement support for date time picker format " + datePickerFormat);
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
                    currentDateText = GUI.m_APE.GetWindowText(Identity.Handle);

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to set the text of the TextBox");
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
