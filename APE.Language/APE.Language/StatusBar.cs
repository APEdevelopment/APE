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
    /// System.Windows.Forms.StatusBar
    /// </summary>
    public sealed class GUIStatusBar : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIStatusBar(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        public int PanelIndex(String PanelName)
        {
            int Items;

            //Get the number of items
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            Items = GUI.m_APE.GetValueFromMessage(1);

            //Loop through looking for the item we want
            for (int Item = 0; Item < Items; Item++)
            {
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Item));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //get the values returned
                string CurrentPanelName = GUI.m_APE.GetValueFromMessage(1);

                if (CurrentPanelName == PanelName)
                {
                    //found it
                    return Item;
                }
            }

            //Failed to find it
            throw new Exception("Failed to find panel with name [" + PanelName + "]");
        }

        public string PanelName(int PanelIndex)
        {
            //query the panel text 0 based index
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, PanelIndex));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            return GUI.m_APE.GetValueFromMessage(1);
        }

        public void PanelPollForText(int PanelIndex, string Text)
        {
            //query the panel text 0 based index
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, PanelIndex));
            GUI.m_APE.AddMessagePollMember(DataStores.Store2, "Text", MemberTypes.Property, new Parameter(GUI.m_APE, Text));
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
        }

        public string PanelText(int PanelIndex)
        {
            //query the panel text 0 based index
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, PanelIndex));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);

            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);

            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

            //get the values returned
            return GUI.m_APE.GetValueFromMessage(1);
        }

        public string PanelText(string PanelName)
        {
            int Items;

            //Get the number of items
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //get the values returned
            Items = GUI.m_APE.GetValueFromMessage(1);

            //Loop through looking for the item we want
            for (int Item = 0; Item < Items; Item++)
            {
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Item));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "Text", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //get the values returned
                string CurrentPanelName = GUI.m_APE.GetValueFromMessage(1);
                string CurrentPanelText = GUI.m_APE.GetValueFromMessage(2);

                if (CurrentPanelName == PanelName)
                {
                    //found it
                    return CurrentPanelText;
                }
            }

            //Failed to find it
            throw new Exception("Failed to find panel with name [" + PanelName + "]");
        }
    }
}
