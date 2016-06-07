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
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    //ToolStripButton			-> rendered
    //ToolStripLabel			-> rendered
    //ToolStripSplitButton		-> rendered
    //ToolStripDropDownButton	-> rendered
    //ToolStripSeparator		-> rendered
    //ToolStripComboBox
    //ToolStripTextBox
    //ToolStripProgressBar

    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.ToolStrip
    /// </summary>
    public class GUIToolStrip : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStrip(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        //TODO support overflow

        public GUIToolStripButton GetButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripButton(this, descriptionOfControl, identParams);
        }

        public GUIToolStripLabel GetLabel(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripLabel(this, descriptionOfControl, identParams);
        }

        public GUIToolStripSplitButton GetSplitButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripSplitButton(this, descriptionOfControl, identParams);
        }

        public GUIToolStripMenu GetMenu(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripMenu(this, descriptionOfControl, identParams);
        }

        public GUIToolStripDropDownButton GetDropDownButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripDropDownButton(this, descriptionOfControl, identParams);
        }

        public GUIToolStripSeparator GetSeparator(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripSeparator(this, descriptionOfControl, identParams);
        }

        public GUIToolStripComboBox GetComboBox(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripComboBox(this, descriptionOfControl, identParams);
        }

        public GUIToolStripTextBox GetTextBox(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripTextBox(this, descriptionOfControl, identParams);
        }

        public GUIToolStripProgressBar GetProgressBar(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripProgressBar(this, descriptionOfControl, identParams);
        }
    }

    public abstract class GUIToolStripObject
    {
        protected GUIToolStrip m_ParentToolStrip;
        protected string m_DescriptionOfControl;
        protected ControlIdentifier m_Identity;
        protected int m_Index = 0;

        protected GUIToolStripObject(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
        {
            m_ParentToolStrip = parentToolStrip;
            m_DescriptionOfControl = descriptionOfControl;

            foreach (Identifier i in identParams)
            {
                switch (i.IdentifierType)
                {
                    case Identifiers.Handle:
                        m_Identity.Handle = i.IdentifierValue;
                        break;
                    case Identifiers.Name:
                        m_Identity.Name = i.IdentifierValue;
                        break;
                    case Identifiers.TechnologyType:
                        m_Identity.TechnologyType = i.IdentifierValue;
                        break;
                    case Identifiers.TypeNameSpace:
                        m_Identity.TypeNameSpace = i.IdentifierValue;
                        break;
                    case Identifiers.TypeName:
                        m_Identity.TypeName = i.IdentifierValue;
                        break;
                    case Identifiers.ModuleName:
                        m_Identity.ModuleName = i.IdentifierValue;
                        break;
                    case Identifiers.AssemblyName:
                        m_Identity.AssemblyName = i.IdentifierValue;
                        break;
                    case Identifiers.Index:
                        m_Identity.Index = i.IdentifierValue;
                        break;
                    case Identifiers.Text:
                        m_Identity.Text = i.IdentifierValue;
                        break;
                }
            }

            UpdateIndex();
        }

        internal bool ItemMatchIdentifier(int item)
        {
            string Name = null;
            string TypeNameSpace = null;
            string TypeName = null;
            string Text = null;
            string ModuleName = null;
            string AssemblyName = null;

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, item));
            if (m_Identity.Name != null)
            {
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
            }
            if (m_Identity.TypeNameSpace != null || m_Identity.TypeName != null || m_Identity.AssemblyName != null || m_Identity.ModuleName != null)
            {
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                if (m_Identity.TypeNameSpace != null)
                {
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store5, "Namespace", MemberTypes.Property);
                }
                if (m_Identity.TypeName != null)
                {
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store6, "Name", MemberTypes.Property);
                }
                if (m_Identity.AssemblyName != null)
                {
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store7, "Assembly", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store7, DataStores.Store8, "GetName", MemberTypes.Method);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store8, DataStores.Store9, "Name", MemberTypes.Property);
                }
                if (m_Identity.ModuleName != null)
                {
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store7, "Module", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store7, DataStores.Store8, "Name", MemberTypes.Property);
                }
            }
            if (m_Identity.Text != null)
            {
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store7, "Text", MemberTypes.Property);
            }

            if (m_Identity.Name != null)
            {
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            }
            if (m_Identity.Text != null)
            {
                GUI.m_APE.AddMessageGetValue(DataStores.Store7);
            }
            if (m_Identity.TypeNameSpace != null)
            {
                GUI.m_APE.AddMessageGetValue(DataStores.Store5);
            }
            if (m_Identity.TypeName != null)
            {
                GUI.m_APE.AddMessageGetValue(DataStores.Store6);
            }
            if (m_Identity.ModuleName != null)
            {
                GUI.m_APE.AddMessageGetValue(DataStores.Store8);
            }
            if (m_Identity.AssemblyName != null)
            {
                GUI.m_APE.AddMessageGetValue(DataStores.Store9);
            }

            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            if (m_Identity.Name != null)
            {
                Name = GUI.m_APE.GetValueFromMessage();
            }
            if (m_Identity.Text != null)
            {
                Text = GUI.m_APE.GetValueFromMessage();
            }
            if (m_Identity.TypeNameSpace != null)
            {
                TypeNameSpace = GUI.m_APE.GetValueFromMessage();
            }
            if (m_Identity.TypeName != null)
            {
                TypeName = GUI.m_APE.GetValueFromMessage();
            }
            if (m_Identity.ModuleName != null)
            {
                ModuleName = GUI.m_APE.GetValueFromMessage();
            }
            if (m_Identity.AssemblyName != null)
            {
                AssemblyName = GUI.m_APE.GetValueFromMessage();
            }

            if (m_Identity.Name != null)
            {
                if (Name != m_Identity.Name)
                {
                    return false;
                }
            }

            if (m_Identity.TechnologyType != null)
            {
                if ("Windows Forms (WinForms)" != m_Identity.TechnologyType)
                {
                    return false;
                }
            }

            if (m_Identity.TypeNameSpace != null)
            {
                if (TypeNameSpace != m_Identity.TypeNameSpace)
                {
                    return false;
                }
            }

            if (m_Identity.TypeName != null)
            {
                if (TypeName != m_Identity.TypeName)
                {
                    return false;
                }
            }

            if (m_Identity.ModuleName != null)
            {
                if (ModuleName != m_Identity.ModuleName)
                {
                    return false;
                }
            }

            if (m_Identity.AssemblyName != null)
            {
                if (AssemblyName != m_Identity.AssemblyName)
                {
                    return false;
                }
            }

            if (m_Identity.Text != null)
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(Text, m_Identity.Text))
                {
                    return false;
                }
            }

            if (m_Identity.Index > 0)
            {
                if (item != m_Identity.Index)
                {
                    return false;
                }
            }

            return true;
        }

        internal void UpdateIndex()
        {
            bool match = false;

            //Get the number of items
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int items = GUI.m_APE.GetValueFromMessage();

            //The old index is larger than the current maximum index so reset the index back to 0
            if (m_Index >= items)
            {
                m_Index = 0;
            }

            //Check the index of the item when it was found first
            match = ItemMatchIdentifier(m_Index);

            if (!match)
            {
                //Look through all the items
                for (int item = 0; item < items; item++)
                {
                    //No need to recheck the value of m_Index as we did it above
                    if (item != m_Index)
                    {
                        match = ItemMatchIdentifier(item);
                        if (match)
                        {
                            m_Index = item;
                            return;
                        }
                    }
                }

                throw new Exception("Failed to find " + m_DescriptionOfControl);
            }
        }
    }

    public abstract class GUIToolStripRenderedObject : GUIToolStripObject
    {
        protected GUIToolStripRenderedObject(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }

        public virtual void MouseSingleClick(MouseButton button)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl, LogItemTypeEnum.Action);
            m_ParentToolStrip.MouseSingleClickInternal(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2), button, MouseKeyModifier.None);
        }

        public virtual void MouseDoubleClick(MouseButton button)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + button.ToString() + " click on " + m_DescriptionOfControl, LogItemTypeEnum.Action);
            m_ParentToolStrip.MouseDoubleClickInternal(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2), button, MouseKeyModifier.None);
        }

        public virtual string Text
        {
            get
            {
                UpdateIndex();

                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemText = GUI.m_APE.GetValueFromMessage();
                return itemText;
            }
        }

        public void PollForText(string Text)
        {
            UpdateIndex();

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessagePollMember(DataStores.Store2, "Text", MemberTypes.Property, new Parameter(GUI.m_APE, Text));
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
        }

        public virtual string ToolTipText
        {
            get
            {
                UpdateIndex();

                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ToolTipText", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemToolTipText = GUI.m_APE.GetValueFromMessage();

                return itemToolTipText;
            }
        }

        public virtual void SaveImage(string filename)
        {
            UpdateIndex();

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Image", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Height", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Save", MemberTypes.Method, new Parameter(GUI.m_APE, filename));
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic height = GUI.m_APE.GetValueFromMessage();

            if (height == null)
            {
                throw new Exception(m_DescriptionOfControl + " does not have an image");
            }
        }

        internal Rectangle ItemBounds()
        {
            UpdateIndex();

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "Bounds", MemberTypes.Property);
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
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int x = GUI.m_APE.GetValueFromMessage();
            int y = GUI.m_APE.GetValueFromMessage();
            int width = GUI.m_APE.GetValueFromMessage();
            int height = GUI.m_APE.GetValueFromMessage();

            Rectangle bounds = new Rectangle(x, y, width, height);
            return bounds;
        }

        internal IntPtr GetDropDown()
        {
            UpdateIndex();

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "DropDown", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic handle = GUI.m_APE.GetValueFromMessage();

            if (handle == null)
            {
                throw new Exception("Could not find dropdown of " + m_DescriptionOfControl);
            }

            return handle;
                
        }
    }

    public sealed class GUIToolStripButton : GUIToolStripRenderedObject
    {
        public GUIToolStripButton(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    public sealed class GUIToolStripLabel : GUIToolStripRenderedObject
    {
        public GUIToolStripLabel(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    public sealed class GUIToolStripSplitButton : GUIToolStripRenderedObject
    {
        MenuUtils m_MenuUtils = new MenuUtils();

        public GUIToolStripSplitButton(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }

        public void Select(string SplitDropDownItem)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + MouseButton.Left.ToString() + " click on " + m_DescriptionOfControl, LogItemTypeEnum.Action);
            m_ParentToolStrip.MouseSingleClickInternal(bounds.Width - 3, bounds.Y + (bounds.Height / 2), MouseButton.Left, MouseKeyModifier.None);

            GUI.Log("Select [" + SplitDropDownItem + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            string[] DropDownItems = SplitDropDownItem.Split(new char[] { '\\' });
            int MenuIndex = 0;
            IntPtr Handle= GetDropDown();

            Input.Block(m_ParentToolStrip.m_ParentForm.Handle, m_Identity.Handle);
            try
            {
                for (int Item = 0; Item < DropDownItems.Length; Item++)
                {
                    if (Item > 0)
                    {
                        Handle = m_MenuUtils.GetDropDown(m_ParentToolStrip.m_ParentForm.Handle, Handle, MenuIndex);
                    }

                    MenuIndex = m_MenuUtils.GetIndexOfMenuItem(m_ParentToolStrip.m_ParentForm.Handle, Handle, DropDownItems[Item]);
                    m_MenuUtils.ClickMenuItem(m_ParentToolStrip.m_ParentForm.Handle, Handle, MenuIndex, DropDownItems[Item], ref m_Identity);
                }
            }
            finally
            {
                Input.Unblock();
            }
        }
    }

    public sealed class GUIToolStripMenu : GUIToolStripDropDownButton
    {
        public GUIToolStripMenu(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    public class GUIToolStripDropDownButton : GUIToolStripRenderedObject
    {
        MenuUtils m_MenuUtils = new MenuUtils();

        public GUIToolStripDropDownButton(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }

        public void Select(string DropDownItem)
        {
            this.MouseSingleClick(MouseButton.Left);

            GUI.Log("Select [" + DropDownItem + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            string[] DropDownItems = DropDownItem.Split(new char[] { '\\' });
            int MenuIndex = 0;
            IntPtr Handle = GetDropDown();

            Input.Block(m_ParentToolStrip.m_ParentForm.Handle, m_Identity.Handle);
            try
            {
                for (int Item = 0; Item < DropDownItems.Length; Item++)
                {
                    if (Item > 0)
                    {
                        Handle = m_MenuUtils.GetDropDown(m_ParentToolStrip.m_ParentForm.Handle, Handle, MenuIndex);
                    }

                    MenuIndex = m_MenuUtils.GetIndexOfMenuItem(m_ParentToolStrip.m_ParentForm.Handle, Handle, DropDownItems[Item]);
                    m_MenuUtils.ClickMenuItem(m_ParentToolStrip.m_ParentForm.Handle, Handle, MenuIndex, DropDownItems[Item], ref m_Identity);
                }
            }
            finally
            {
                Input.Unblock();
            }
        }
    }

    public sealed class GUIToolStripSeparator : GUIToolStripRenderedObject
    {
        public GUIToolStripSeparator(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    public sealed class GUIToolStripComboBox : GUIToolStripObject
    {
        private GUIComboBox m_ComboBox;

        public GUIToolStripComboBox(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
            IntPtr comboBoxHandle = ItemComboBoxHandle(descriptionOfControl, base.m_Identity);
            m_ComboBox = new GUIComboBox(parentToolStrip.m_ParentForm, descriptionOfControl, new Identifier(Identifiers.Handle, comboBoxHandle));
        }

        public bool ItemExists(string item)
        {
            return m_ComboBox.ItemExists(item);
        }

        public void ItemSelect(string item)
        {
            m_ComboBox.ItemSelect(item);
        }

        public void SetText(string text)
        {
            m_ComboBox.SetText(text);
        }

        internal IntPtr ItemComboBoxHandle(string descriptionOfControl, ControlIdentifier identity)
        {
            UpdateIndex();

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ComboBox", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic comboBoxHandle = GUI.m_APE.GetValueFromMessage();

            if (comboBoxHandle != null)
            {
                return comboBoxHandle;
            }

            throw new Exception("Failed to find the combobox of " + m_DescriptionOfControl);
        }
    }

    public sealed class GUIToolStripTextBox : GUIToolStripObject
    {
        private GUITextBox m_TextBox;

        public GUIToolStripTextBox(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
            IntPtr textBoxHandle = ItemTextBoxHandle(descriptionOfControl, base.m_Identity);
            m_TextBox = new GUITextBox(parentToolStrip.m_ParentForm, descriptionOfControl, new Identifier(Identifiers.Handle, textBoxHandle));

        }

        public void SetText(string text)
        {
            m_TextBox.SetText(text);
        }

        public void MouseSingleClick(MouseButton button)
        {
            m_TextBox.MouseSingleClick(button);
        }

        public void MouseSingleClick(int x, int y, MouseButton button)
        {
            m_TextBox.MouseSingleClick(x, y, button);
        }

        public void MouseSingleClick(int x, int y, MouseButton button, MouseKeyModifier keys)
        {
            m_TextBox.MouseSingleClick(x, y, button, keys);
        }

        public void MouseDoubleClick(MouseButton button)
        {
            m_TextBox.MouseDoubleClick(button);
        }

        public void MouseDoubleClick(int x, int y, MouseButton button)
        {
            m_TextBox.MouseDoubleClick(x, y, button);
        }

        public void MouseDoubleClick(int x, int y, MouseButton button, MouseKeyModifier keys)
        {
            m_TextBox.MouseDoubleClick(x, y, button, keys);
        }

        internal IntPtr ItemTextBoxHandle(string descriptionOfControl, ControlIdentifier identity)
        {
            UpdateIndex();

            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "TextBox", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic textBoxHandle = GUI.m_APE.GetValueFromMessage();

            if (textBoxHandle != null)
            {
                return textBoxHandle;
            }

            throw new Exception("Failed to find the textbox of " + descriptionOfControl);
        }
    }

    public sealed class GUIToolStripProgressBar : GUIToolStripObject
    {
        private GUIProgressBar m_ProgressBar;

        public GUIToolStripProgressBar(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
            IntPtr progressBarHandle = ItemProgressBarHandle(descriptionOfControl, base.m_Identity);
            m_ProgressBar = new GUIProgressBar(parentToolStrip.m_ParentForm, descriptionOfControl, new Identifier(Identifiers.Handle, progressBarHandle));
        }

        public int Minimum
        {
            get
            {
                return m_ProgressBar.Minimum;
            }
        }

        public int Maximum
        {
            get
            {
                return m_ProgressBar.Maximum;
            }
        }

        public int Value
        {
            get
            {
                return m_ProgressBar.Value;
            }
        }

        internal IntPtr ItemProgressBarHandle(string descriptionOfControl, ControlIdentifier identity)
        {
            UpdateIndex();
            
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, m_ParentToolStrip.m_ParentForm.Handle, m_ParentToolStrip.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "ProgressBar", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic progressBarHandle = GUI.m_APE.GetValueFromMessage();

            if (progressBarHandle != null)
            {
                return progressBarHandle;
            }
             
            throw new Exception("Failed to find the progressbar of " + descriptionOfControl);
        }
    }
}
