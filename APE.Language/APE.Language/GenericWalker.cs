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
using Parameter = APE.Communication.APEIPC.Parameter;
using System.Threading;
using System.Drawing.Imaging;
using System.Security.Principal;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    //TODO actually make this the walker not the textbox of the walker
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Capstone.Controls.Walker.WalkerControl
    /// </summary>
    public sealed class GUIGenericWalker : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIGenericWalker(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        public void SetText(string text)
        {
            Stopwatch timer;
            string CurrentText;

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                CurrentText = GUI.m_APE.GetWindowText(Identity.Handle);

                if (CurrentText != "")
                {
                    base.MouseDoubleClick(MouseButton.Left);

                    string selectedText;

                    //wait for .selectedText to = Text
                    timer = Stopwatch.StartNew();
                    do
                    {
                        //Get the selectedText property
                        GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                        GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "SelectedText", MemberTypes.Property);
                        GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store1);
                        GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                        GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                        //Get the value(s) returned MUST be done straight after the WaitForMessages call
                        selectedText = GUI.m_APE.GetValueFromMessage(1);

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to select all the text in the TextBox");
                        }

                        Thread.Sleep(15);
                    }
                    while (CurrentText != selectedText);
                    timer.Stop();
                }

                //Send first 2 characters
                base.SendKeys(text.Substring(0, 2));

                //Wait for popup
                string PopupState;
                timer = Stopwatch.StartNew();
                do
                {
                    //Get the selectedText property
                    GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "Parent", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store1, APEIPC.DataStores.Store2, "PopupState", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store2, APEIPC.DataStores.Store3, "ToString", MemberTypes.Method);
                    GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store3);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    PopupState = GUI.m_APE.GetValueFromMessage(1);
                }
                while (PopupState != "Open");
                timer.Stop();

                //Send rest of characters
                base.SendKeys(text.Substring(2));

                //wait for .Text to == text
                timer = Stopwatch.StartNew();
                do
                {
                    CurrentText = GUI.m_APE.GetWindowText(Identity.Handle);

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to set the text of the TextBox");
                    }

                    Thread.Sleep(15);
                }
                while (CurrentText != text);
                timer.Stop();

                GUI.Log("Press Enter to set the value", LogItemTypeEnum.Action);
                base.SendKeys("{Enter}");
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
