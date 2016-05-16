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
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.ComboBox
    /// LatentZero.Utility.Controls.GUIComboBox
    /// </summary>
    public sealed class GUIComboBox : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIComboBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Checks if the specified item exists in the combobox
        /// </summary>
        /// <param name="Item">The item to check if it exists</param>
        /// <returns></returns>
        public bool ItemExists(string Item)
        {
            int Index = ItemIndex(Item);

            if (Index == NM.CB_ERR)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>
        /// Selects the specified item in the combobox
        /// </summary>
        /// <param name="Item">The item to select</param>
        public void ItemSelect(string Item)
        {
            Stopwatch timer;

            GUI.Log("Select [" + Item + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            //TODO check if already set (no need to set it again)

            //Get the style
            GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "DropDownStyle", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store1, APEIPC.DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store3, "DroppedDown", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store2);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string Style = GUI.m_APE.GetValueFromMessage(1);
            bool DroppedDown = GUI.m_APE.GetValueFromMessage(2);

            IntPtr ListBox = IntPtr.Zero;

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                if (Style == "Simple")
                {
                    //get the Simple mode listbox child window
                    GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "childListBox", MemberTypes.Field);
                    GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store1, APEIPC.DataStores.Store2, "Handle", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store2);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    ListBox = (IntPtr)GUI.m_APE.GetValueFromMessage(1);
                }
                else
                {
                    if (!DroppedDown)
                    {
                        //show the dropdown
                        base.MouseSingleClickInternal(Width - 5, -1, MouseButton.Left, MouseKeyModifier.None);
                    }
                    
                    //find the dropdown
                    Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut);
                    GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "dropDownHandle", MemberTypes.Field);
                    GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store1);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    dynamic DroppedDownHandle = GUI.m_APE.GetValueFromMessage(1);

                    if (DroppedDownHandle == null)
                    {
                        throw new Exception("Failed to find the ComboBox dropdown");
                    }

                    ListBox = DroppedDownHandle;
                }

                //locate the item
                int Index = ItemIndex(Item);
                if (Index == NM.CB_ERR)
                {
                    throw new Exception("Failed to find the ComboBox item");
                }

                IntPtr MessageResult;
                IntPtr SendResult;
                NM.RECT ItemRect = new NM.RECT();

                //Locate the rect of the item
                SendResult = NM.SendMessageTimeout(ListBox, NM.ListBoxMessages.LB_GETITEMRECT, new IntPtr(Index), ref ItemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                if (SendResult == IntPtr.Zero || MessageResult.ToInt64() == NM.LB_ERR)  //Failed
                {
                    throw new Exception("Failed to access the combobox");
                }

                NM.RECT ClientRect;
                NM.GetClientRect(ListBox, out ClientRect);

                //scroll the item into view if needed
                if (((ItemRect.Bottom - ItemRect.Top) / 2) + ItemRect.Top > ClientRect.Bottom || ((ItemRect.Bottom - ItemRect.Top) / 2) + ItemRect.Top < ClientRect.Top)
                {
                    SendResult = NM.SendMessageTimeout(ListBox, NM.ListBoxMessages.LB_SETTOPINDEX, new IntPtr(Index), ref ItemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                    if (SendResult == IntPtr.Zero || MessageResult.ToInt64() == NM.LB_ERR)  //Failed
                    {
                        throw new Exception("Failed to access the combobox");
                    }

                    //Locate the rect of the item
                    SendResult = NM.SendMessageTimeout(ListBox, NM.ListBoxMessages.LB_GETITEMRECT, new IntPtr(Index), ref ItemRect, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
                    if (SendResult == IntPtr.Zero || MessageResult.ToInt64() == NM.LB_ERR)  //Failed
                    {
                        throw new Exception("Failed to access the combobox");
                    }
                }

                //click the item
                IntPtr temp = Identity.Handle;
                try
                {
                    Identity.Handle = ListBox;
                    base.MouseSingleClickInternal(-1, ((ItemRect.Bottom - ItemRect.Top) / 2) + ItemRect.Top, MouseButton.Left, MouseKeyModifier.None);
                }
                finally
                {
                    Identity.Handle = temp;
                }

                //wait for .Text to == text
                string CurrentText;
                timer = Stopwatch.StartNew();
                do
                {
                    CurrentText = GUI.m_APE.GetWindowText(Identity.Handle);

                    if (CurrentText == Item)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to set the text of the ComboBox");
                    }

                    Thread.Sleep(15);
                }
                while (true);
                timer.Stop();
            }
            finally
            {
                Input.Unblock();
            }
        }

        /// <summary>
        /// Sets the text portion of the combobox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text portion of the combobox to</param>
        public void SetText(string text)
        {
            //Get the style
            GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "DropDownStyle", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store1, APEIPC.DataStores.Store2, "ToString", MemberTypes.Method);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string Style = GUI.m_APE.GetValueFromMessage(1);

            if (Style == "DropDownList")
            {
                throw new Exception("ComboBox Edit control is not editable");
            }

            //get the editbox child window
            GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "childEdit", MemberTypes.Field);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store1, APEIPC.DataStores.Store2, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            IntPtr EditBox = (IntPtr)GUI.m_APE.GetValueFromMessage(1);

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                GUITextBox comboboxTextBox = new GUITextBox(m_ParentForm, m_DescriptionOfControl + " textbox", new Identifier(Identifiers.Handle, EditBox), new Identifier(Identifiers.TechnologyType, "Windows Native"));
                comboboxTextBox.SetText(text);
            }
            finally
            {
                Input.Unblock();
            }
        }

        private int ItemIndex(string Item)
        {
            IntPtr MessageResult;
            IntPtr SendResult;

            SendResult = NM.SendMessageTimeout(Identity.Handle, NM.ComboBoxMessages.CB_FINDSTRINGEXACT, new IntPtr(-1), Item, NM.SendMessageTimeoutFlags.SMTO_NORMAL, GUI.m_APE.TimeOut, out MessageResult);
            if (SendResult == IntPtr.Zero || MessageResult.ToInt64() == NM.CB_ERR) //Failed
            {
                throw new Exception("Failed to access the combobox");
            }
            else
            {
                return MessageResult.ToInt32();
            }
        }
    }
}
