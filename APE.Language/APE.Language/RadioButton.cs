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
                    throw new Exception("RadioButton failed to be set to state [" + State.ToString() + "]");
                }

                Thread.Sleep(15);
            }
            while (GetState() != State);
            timer.Stop();
        }

        public override void MouseSingleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            base.MouseSingleClick(X, Y, Button, Keys);
            if (Button == MouseButton.Left)
            {
                pollForState(true);
            }
        }

        public override void MouseDoubleClick(int X, int Y, MouseButton Button)
        {
            base.MouseDoubleClick(X, Y, Button);
            if (Button == MouseButton.Left)
            {
                pollForState(true);
            }
        }

        public override void MouseTripleClick(int X, int Y, MouseButton Button)
        {
            base.MouseTripleClick(X, Y, Button);
            if (Button == MouseButton.Left)
            {
                pollForState(true);
            }
        }

        public bool Checked
        {
            get
            {
                return GetState();
            }
        }

        private bool GetState()
        {
            //Get the state
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Checked", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            return GUI.m_APE.GetValueFromMessage();
        }
    }
}
