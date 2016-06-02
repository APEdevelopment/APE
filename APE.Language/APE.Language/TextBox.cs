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
    /// System.Windows.Forms.TextBox
    /// </summary>
    public sealed class GUITextBox : GUIFocusableObject, IGUITextBox
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITextBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Sets the text of the textbox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text of the textbox to</param>
        public void SetText(string text)
        {
            Stopwatch timer;
            string currentText;

            currentText = GUI.m_APE.GetWindowText(Identity.Handle);
            if (currentText == text)
            {
                GUI.Log("Ensure " + m_DescriptionOfControl + " is set to " + text, LogItemTypeEnum.Action);
                return;
            }

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                currentText = GUI.m_APE.GetWindowText(Identity.Handle);

                if (currentText != "")
                {
                    base.MouseDoubleClick(MouseButton.Left);

                    string selectedText;

                    //wait for .selectedText to = Text
                    timer = Stopwatch.StartNew();
                    do
                    {
                        int start = 0;
                        int end = 0;
                        IntPtr Return;
                        IntPtr Result;

                        currentText = GUI.m_APE.GetWindowText(Identity.Handle);
                        Return = NM.SendMessageTimeout(Identity.Handle, NM.EM_GETSEL, ref start, ref end, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out Result);
                        selectedText = currentText.Substring(start, end);

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to select all the text in the TextBox");
                        }

                        Thread.Sleep(15);
                    }
                    while (currentText != selectedText);
                    timer.Stop();
                }
                
                base.SendKeys(text);
                
                //wait for .Text to == text
                timer = Stopwatch.StartNew();
                do
                {
                    currentText = GUI.m_APE.GetWindowText(Identity.Handle);

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to set the text of the TextBox");
                    }

                    Thread.Sleep(15);
                }
                while (currentText != text);
                timer.Stop();
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
