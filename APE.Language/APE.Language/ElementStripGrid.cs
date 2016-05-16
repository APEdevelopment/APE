using System;
using System.Drawing;
using System.Text;
using System.Reflection;
using APE.Capture;
using APE.Communication;
using Parameter = APE.Communication.APEIPC.Parameter;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;
using NM = APE.Native.NativeMethods;
using System.Xml;

namespace APE.Language
{
    // TODO this should probably extend GridControl with most of the methods in there


    public sealed class GUIElementStripGrid : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIElementStripGrid(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        public int TitleRows()
        {
            GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "GetColumnInfoXML", MemberTypes.Method);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store1);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string ColumnInfoXML = GUI.m_APE.GetValueFromMessage(1);

            XmlDocument columnDocument = new XmlDocument();
            columnDocument.LoadXml(ColumnInfoXML);

            int maxLevels = -1;
            int currentLevel = -1;
            foreach (XmlNode node in columnDocument.SelectSingleNode("Columns").ChildNodes)
            {
                if (int.TryParse(node.Attributes.GetNamedItem("ColumnHeaderLevel").Value, out currentLevel))
                {
                    if (currentLevel > maxLevels)
                    {
                        maxLevels = currentLevel;
                    }
                }
            }

            return maxLevels + 1;
        }

        public int Rows()
        {
            GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "NumRows", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store1);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Rows = GUI.m_APE.GetValueFromMessage(1);

            return Rows + TitleRows();
        }

        public int FixedRows()
        {
            GUI.m_APE.AddMessageFindByHandle(APEIPC.DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(APEIPC.DataStores.Store0, APEIPC.DataStores.Store1, "NumFrozenRows", MemberTypes.Method);
            GUI.m_APE.AddMessageGetValue(APEIPC.DataStores.Store1);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedRows = GUI.m_APE.GetValueFromMessage(1);

            return FixedRows + TitleRows();
        }
    }
}
