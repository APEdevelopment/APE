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
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LatentZero.Utility.Controls.TitleFrame
    /// </summary>
    public sealed class GUITitleFrame : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITitleFrame(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Returns a GUITitleFrameButton object which can be used to automate a titel frame button
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUITitleFrameButton object</returns>
        public GUITitleFrameButton GetButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUITitleFrameButton(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Gets the currently displayed subtitle
        /// </summary>
        /// <returns>The currently displayed subtitle</returns>
        public string Subtitle()
        {
            Input.WaitForInputIdle(Identity.Handle, 250);

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetCurrentSubtitle", MemberTypes.Method);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            string subtitle = GUI.m_APE.GetValueFromMessage();

            return subtitle;
        }

        /// <summary>
        /// Determines if the title frame can be maximised
        /// </summary>
        /// <returns>True if the title frame can be maximised otherwise false</returns>
        public bool CanMaximise()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "CanMaximise", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            bool canMaximise = GUI.m_APE.GetValueFromMessage();

            return canMaximise;
        }

        /// <summary>
        /// Determines if the titel frame can be maximised
        /// </summary>
        /// <returns>True if the title frame is maximised otherwise false</returns>
        public bool IsMaximised()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "IsMaximised", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            bool isMaximised = GUI.m_APE.GetValueFromMessage();

            return isMaximised;
        }
    }

    /// <summary>
    /// Abstract base object which all other title frame objects extend
    /// </summary>
    public abstract class GUITitleFrameObject
    {
        /// <summary>
        /// The identity of the title frame object
        /// </summary>
        protected ControlIdentifier m_Identity;
        /// <summary>
        /// The index of the title frame object within the TitleFrame
        /// </summary>
        protected int m_Index = 0;
        /// <summary>
        /// The TitleFrame which the title frame object belongs to
        /// </summary>
        protected GUITitleFrame m_ParentTitleFrame;
        /// <summary>
        /// The human readable description of the title frame object
        /// </summary>
        protected string m_DescriptionOfControl;

        /// <summary>
        /// Constructor used for TitleFrame objects
        /// </summary>
        /// <param name="parentTitleFrame">The parent title frame the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        protected GUITitleFrameObject(GUITitleFrame parentTitleFrame, string descriptionOfControl, params Identifier[] identParams)
        {
            m_ParentTitleFrame = parentTitleFrame;
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
                    default:
                        throw new Exception("Unsupported identifier: " + i.ToString());
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

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_ParentTitleFrame.m_ParentForm.Handle, m_ParentTitleFrame.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_toolButtons", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, item));
            if (m_Identity.Name != null)
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
            }
            if (m_Identity.TypeNameSpace != null || m_Identity.TypeName != null || m_Identity.AssemblyName != null || m_Identity.ModuleName != null)
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                if (m_Identity.TypeNameSpace != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Namespace", MemberTypes.Property);
                }
                if (m_Identity.TypeName != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store6, "Name", MemberTypes.Property);
                }
                if (m_Identity.AssemblyName != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store7, "Assembly", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "GetName", MemberTypes.Method);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store8, DataStores.Store9, "Name", MemberTypes.Property);
                }
                if (m_Identity.ModuleName != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store7, "Module", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "Name", MemberTypes.Property);
                }
            }
            if (m_Identity.Text != null)
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store7, "Text", MemberTypes.Property);
            }

            if (m_Identity.Name != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            }
            if (m_Identity.Text != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            }
            if (m_Identity.TypeNameSpace != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            }
            if (m_Identity.TypeName != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            }
            if (m_Identity.ModuleName != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            }
            if (m_Identity.AssemblyName != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
            }

            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_ParentTitleFrame.m_ParentForm.Handle, m_ParentTitleFrame.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_toolButtons", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
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

                throw new Exception("Failed to find the " + m_DescriptionOfControl);
            }
        }
    }

    /// <summary>
    /// Abstract extension of the GUITitleFrameObject object which all other rendered TitleFrame objects extend
    /// </summary>
    public abstract class GUITitleFrameRenderedObject : GUITitleFrameObject
    {
        /// <summary>
        /// Constructor used for rendered title frame objects
        /// </summary>
        /// <param name="parentTitleFrame">The parent title frame the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        protected GUITitleFrameRenderedObject(GUITitleFrame parentTitleFrame, string descriptionOfControl, params Identifier[] identParams)
            : base(parentTitleFrame, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the title frame object
        /// </summary>
        /// <param name="button">The button to click</param>
        public virtual void MouseSingleClick(MouseButton button)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + button.ToString() + " click on the " + m_DescriptionOfControl, LogItemType.Action);
            m_ParentTitleFrame.MouseSingleClickInternal(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2), button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the title frame object
        /// </summary>
        /// <param name="button">The button to double click</param>
        public virtual void MouseDoubleClick(MouseButton button)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + button.ToString() + " click on the " + m_DescriptionOfControl, LogItemType.Action);
            m_ParentTitleFrame.MouseDoubleClickInternal(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2), button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Moves the mouse cursor to the middle of the title frame object
        /// </summary>
        public virtual void MouseMove()
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Move the mouse over the " + m_DescriptionOfControl, LogItemType.Action);
            m_ParentTitleFrame.MouseMove(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2));
        }

        /// <summary>
        /// The tool tip text of the title frame object
        /// </summary>
        public virtual string ToolTipText
        {
            get
            {
                UpdateIndex();

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_ParentTitleFrame.m_ParentForm.Handle, m_ParentTitleFrame.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_toolButtons", MemberTypes.Field);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ToolTip", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemToolTipText = GUI.m_APE.GetValueFromMessage();

                return itemToolTipText;
            }
        }

        internal Rectangle ItemBounds()
        {
            UpdateIndex();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_ParentTitleFrame.m_ParentForm.Handle, m_ParentTitleFrame.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_toolButtons", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, m_Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "m_rect", MemberTypes.Field);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "X", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store6, "Y", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store7, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store8, "Height", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int x = GUI.m_APE.GetValueFromMessage();
            int y = GUI.m_APE.GetValueFromMessage();
            int width = GUI.m_APE.GetValueFromMessage();
            int height = GUI.m_APE.GetValueFromMessage();

            Rectangle bounds = new Rectangle(x, y, width, height);
            return bounds;
        }
    }

    /// <summary>
    /// Automation class used to automate title frame buttons
    /// </summary>
    public sealed class GUITitleFrameButton : GUITitleFrameRenderedObject
    {
        /// <summary>
        /// Constructor used for title frame buttons
        /// </summary>
        /// <param name="parentTitleFrame">The parent title frame the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITitleFrameButton(GUITitleFrame parentTitleFrame, string descriptionOfControl, params Identifier[] identParams)
            : base(parentTitleFrame, descriptionOfControl, identParams)
        {
        }
    }
}
