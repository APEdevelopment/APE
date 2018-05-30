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
    /// System.Windows.Forms.CheckBox
    /// </summary>
    public sealed class GUICheckStateControl : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUICheckStateControl(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
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
            int InitialState = GetState();
            base.SingleClick(X, Y, Button, Keys);
            if (Button == MouseButton.Left)
            {
                pollForState(InitialState, false);
            }
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="Button">The button to double click</param>
        /// <param name="Keys">The key to hold while double clicking</param>
        public override void DoubleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            int InitialState = GetState();
            base.DoubleClick(X, Y, Button);
            if (Button == MouseButton.Left)
            {
                pollForState(InitialState, false);
            }
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="Button">The button to triple click</param>
        /// <param name="Keys">The key to hold while triple clicking</param>
        public override void TripleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            int InitialState = GetState();
            base.TripleClick(X, Y, Button);
            if (Button == MouseButton.Left)
            {
                pollForState(InitialState, false);
            }
        }

        /// <summary>
        /// Gets the state of the control
        /// </summary>
        public int State
        {
            get
            {
                return GetState();
            }
        }

        private int GetState()
        {
            //Get the number of items
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CurrentStateId", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int state = GUI.m_APE.GetValueFromMessage();
            return state;
        }

        private void pollForState(int state, bool waitToBeThisState)
        {
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException(Identity.Description + " failed to be set to state [" + state.ToString() + "]");
                }

                if (waitToBeThisState)
                {
                    if (GetState() == state)
                    {
                        break;
                    }
                }
                else
                {
                    if (GetState() != state)
                    {
                        break;
                    }
                }

                Thread.Sleep(15);
            }
            timer.Stop();
        }
    }
}
