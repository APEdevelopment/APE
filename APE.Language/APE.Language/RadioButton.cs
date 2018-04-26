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
    /// System.Windows.Forms.RadioButton
    /// </summary>
    public sealed class GUIRadioButton : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIRadioButton(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        private void pollForState(bool State)
        {
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException("RadioButton failed to be set to state [" + State.ToString() + "]");
                }

                Thread.Sleep(15);
            }
            while (GetState() != State);
            timer.Stop();
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="Button">The button to click</param>
        /// <param name="Keys">The key to hold while clicking</param>
        public override void SingleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            base.SingleClick(X, Y, Button, Keys);
            if (Button == MouseButton.Left)
            {
                pollForState(true);
            }
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="Button">The button to double click</param>
        public override void DoubleClick(int X, int Y, MouseButton Button)
        {
            base.DoubleClick(X, Y, Button);
            if (Button == MouseButton.Left)
            {
                pollForState(true);
            }
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="Button">The button to triple click</param>
        public override void TripleClick(int X, int Y, MouseButton Button)
        {
            base.TripleClick(X, Y, Button);
            if (Button == MouseButton.Left)
            {
                pollForState(true);
            }
        }

        /// <summary>
        /// Whether the radio button is selected or not
        /// </summary>
        public bool Selected
        {
            get
            {
                return GetState();
            }
        }

        /// <summary>
        /// Selects the radio button
        /// </summary>
        public void Select()
        {
            if (GetState() == true)
            {
                GUI.Log("Ensure " + Identity.Description + " is selected", LogItemType.Action);
            }
            else
            {
                this.SingleClick(MouseButton.Left);
            }
        }

        private bool GetState()
        {
            //Get the state
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Checked", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            return GUI.m_APE.GetValueFromMessage();
        }
    }
}
