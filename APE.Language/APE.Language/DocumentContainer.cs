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
    /// TD.SandDock.DocumentContainer
    /// </summary>
    public sealed class GUIDocumentContainer : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIDocumentContainer(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        private int DockControls()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "GetLength", MemberTypes.Method, new Parameter(GUI.m_APE, 0));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int NumberOfItems = GUI.m_APE.GetValueFromMessage();

            return NumberOfItems;
        }

        public int ItemCount()
        {
            int Count = 0;
            int NumberOfDockControls = DockControls();

            for (int Tab = 0; Tab < NumberOfDockControls; Tab++)
            {
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Parent", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store5, "Handle", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                IntPtr ParentHandle = GUI.m_APE.GetValueFromMessage();

                if (ParentHandle == this.Handle)
                {
                    Count++;
                }
            }

            return Count;
        }

        public bool ItemExists(string Item)
        {
            int NumberOfDockControls = DockControls();

            for (int Tab = 0; Tab < NumberOfDockControls; Tab++)
            {
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Parent", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store5, DataStores.Store6, "Handle", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddMessageGetValue(DataStores.Store6);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string TabText = GUI.m_APE.GetValueFromMessage();
                IntPtr ParentHandle = GUI.m_APE.GetValueFromMessage();

                if (ParentHandle == this.Handle)
                {
                    if (TabText == Item)
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        public void ItemSelect(string Item)
        {
            ItemSelect(Item, MouseButton.Left);
        }

        public void ItemSelect(string Item, MouseButton Button)
        {
            GUI.Log("Select [" + Item + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "GetLength", MemberTypes.Method, new Parameter(GUI.m_APE, 0));
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int NumberOfItems = GUI.m_APE.GetValueFromMessage();

            for (int Tab = 0; Tab < NumberOfItems; Tab++)
            {
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "TabText", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string TabText = GUI.m_APE.GetValueFromMessage();

                if (TabText == Item)
                {
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetDockControls", MemberTypes.Method);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "GetValue", MemberTypes.Method, new Parameter(GUI.m_APE, Tab));
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "TabBounds", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store5, "X", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store6, "Y", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store7, "Width", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store8, "Height", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store5);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store6);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store7);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store8);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int X = GUI.m_APE.GetValueFromMessage();
                    int Y = GUI.m_APE.GetValueFromMessage();
                    int Width = GUI.m_APE.GetValueFromMessage();
                    int Height = GUI.m_APE.GetValueFromMessage();
                    
                    base.MouseSingleClickInternal(X + (Width / 2), Y + (Height / 2), Button, MouseKeyModifier.None);

                    //Wait for the active document to be the tab we selected
                    string ActiveTab;
                    Stopwatch timer = Stopwatch.StartNew();
                    do
                    {
                        ActiveTab = ItemActive();

                        if (ActiveTab != TabText)
                        {
                            break;
                        }

                        if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                        {
                            throw new Exception("Failed to select document container tab");
                        }

                        Thread.Sleep(50);
                    }
                    while (true);
                    timer.Stop();

                    return;
                }
            }

            throw new Exception("Failed to find DocumentContainer tab");
        }

        public string ItemActive()
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Manager", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "ActiveTabbedDocument", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "TabText", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string TabText = GUI.m_APE.GetValueFromMessage();

            return TabText;
        }

        public void ItemRemove(string Item)
        {
            if (ItemActive() != Item)
            {
                ItemSelect(Item);
            }

            //Get the size of the window
            NM.tagRect ClientRect;
            NM.GetClientRect(Identity.Handle, out ClientRect);

            int InitialItems = ItemCount();
            int CurrentItems = InitialItems;

            GUI.Log("Remove [" + Item + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            //Click 10 pixels in from the right hand side of the window and 10 pixels down
            base.MouseSingleClickInternal(ClientRect.right -10, 10, MouseButton.Left, MouseKeyModifier.None);

            //Wait for the number of items to decrease by one
            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (InitialItems > 1)
                {
                    CurrentItems = ItemCount();
                }
                else
                {
                    if (!this.Exists)
                    {
                        CurrentItems = 0;
                    }
                }

                if (CurrentItems == InitialItems - 1)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to remove document container tab");
                }

                Thread.Sleep(50);
            }
            while (true);
            timer.Stop();
        }
    }
}
