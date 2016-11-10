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

        /// <summary>
        /// Returns a GUIToolStripButton object which can be used to automate a toolstrip button
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripButton object</returns>
        public GUIToolStripButton GetButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripButton(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripLabel object which can be used to automate a toolstrip label
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripLabel object</returns>
        public GUIToolStripLabel GetLabel(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripLabel(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GetSplitButton object which can be used to automate a toolstrip split button
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GetSplitButton object</returns>
        public GUIToolStripSplitButton GetSplitButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripSplitButton(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripMenu object which can be used to automate a toolstrip menu
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripMenu object</returns>
        public GUIToolStripMenu GetMenu(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripMenu(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripDropDownButton object which can be used to automate a toolstrip drop down button
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripDropDownButton object</returns>
        public GUIToolStripDropDownButton GetDropDownButton(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripDropDownButton(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripSeparator object which can be used to automate a toolstrip separator
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripSeparator object</returns>
        public GUIToolStripSeparator GetSeparator(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripSeparator(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripComboBox object which can be used to automate a toolstrip combobox
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripComboBox object</returns>
        public GUIToolStripComboBox GetComboBox(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripComboBox(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripTextBox object which can be used to automate a toolstrip textbox
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripTextBox object</returns>
        public GUIToolStripTextBox GetTextBox(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripTextBox(this, descriptionOfControl, identParams);
        }

        /// <summary>
        /// Returns a GUIToolStripProgressBar object which can be used to automate a toolstrip progress bar
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        /// <returns>The GUIToolStripProgressBar object</returns>
        public GUIToolStripProgressBar GetProgressBar(string descriptionOfControl, params Identifier[] identParams)
        {
            return new GUIToolStripProgressBar(this, descriptionOfControl, identParams);
        }
    }

    /// <summary>
    /// Abstract base object which all other toolstrip objects extend
    /// </summary>
    public abstract class GUIToolStripObject
    {
        /// <summary>
        /// The identity of the toolstrip object
        /// </summary>
        protected ControlIdentifier Identity;
        /// <summary>
        /// The index of the toolstrip object within the toolstrip
        /// </summary>
        protected int Index = 0;
        /// <summary>
        /// The toolstrip which the toolstrip object belongs to
        /// </summary>
        protected GUIToolStrip ParentToolStrip;

        /// <summary>
        /// Constructor used for toolstrip objects
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        protected GUIToolStripObject(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
        {
            ParentToolStrip = parentToolStrip;
            Identity = GUI.BuildIdentity(null, descriptionOfControl, identParams);
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

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, item));
            if (Identity.Name != null)
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
            }
            if (Identity.TypeNameSpace != null || Identity.TypeName != null || Identity.AssemblyName != null || Identity.ModuleName != null)
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                if (Identity.TypeNameSpace != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Namespace", MemberTypes.Property);
                }
                if (Identity.TypeName != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store6, "Name", MemberTypes.Property);
                }
                if (Identity.AssemblyName != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store7, "Assembly", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "GetName", MemberTypes.Method);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store8, DataStores.Store9, "Name", MemberTypes.Property);
                }
                if (Identity.ModuleName != null)
                {
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store7, "Module", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "Name", MemberTypes.Property);
                }
            }
            if (Identity.Text != null)
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store7, "Text", MemberTypes.Property);
            }

            if (Identity.Name != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            }
            if (Identity.Text != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            }
            if (Identity.TypeNameSpace != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
            }
            if (Identity.TypeName != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            }
            if (Identity.ModuleName != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            }
            if (Identity.AssemblyName != null)
            {
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
            }

            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            if (Identity.Name != null)
            {
                Name = GUI.m_APE.GetValueFromMessage();
            }
            if (Identity.Text != null)
            {
                Text = GUI.m_APE.GetValueFromMessage();
            }
            if (Identity.TypeNameSpace != null)
            {
                TypeNameSpace = GUI.m_APE.GetValueFromMessage();
            }
            if (Identity.TypeName != null)
            {
                TypeName = GUI.m_APE.GetValueFromMessage();
            }
            if (Identity.ModuleName != null)
            {
                ModuleName = GUI.m_APE.GetValueFromMessage();
            }
            if (Identity.AssemblyName != null)
            {
                AssemblyName = GUI.m_APE.GetValueFromMessage();
            }

            if (Identity.Name != null)
            {
                if (Name != Identity.Name)
                {
                    return false;
                }
            }

            if (Identity.TechnologyType != null)
            {
                if ("Windows Forms (WinForms)" != Identity.TechnologyType)
                {
                    return false;
                }
            }

            if (Identity.TypeNameSpace != null)
            {
                if (TypeNameSpace != Identity.TypeNameSpace)
                {
                    return false;
                }
            }

            if (Identity.TypeName != null)
            {
                if (TypeName != Identity.TypeName)
                {
                    return false;
                }
            }

            if (Identity.ModuleName != null)
            {
                if (ModuleName != Identity.ModuleName)
                {
                    return false;
                }
            }

            if (Identity.AssemblyName != null)
            {
                if (AssemblyName != Identity.AssemblyName)
                {
                    return false;
                }
            }

            if (Identity.Text != null)
            {
                if (!System.Text.RegularExpressions.Regex.IsMatch(Text, Identity.Text))
                {
                    return false;
                }
            }

            if (Identity.Index > 0)
            {
                if (item != Identity.Index)
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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int items = GUI.m_APE.GetValueFromMessage();

            //The old index is larger than the current maximum index so reset the index back to 0
            if (Index >= items)
            {
                Index = 0;
            }

            //Check the index of the item when it was found first
            match = ItemMatchIdentifier(Index);

            if (!match)
            {
                //Look through all the items
                for (int item = 0; item < items; item++)
                {
                    //No need to recheck the value of m_Index as we did it above
                    if (item != Index)
                    {
                        match = ItemMatchIdentifier(item);
                        if (match)
                        {
                            Index = item;
                            return;
                        }
                    }
                }

                throw new Exception("Failed to find " + Identity.Description);
            }
        }
    }

    /// <summary>
    /// Abstract extension of the GUIToolStripObject object which all other rendered toolstrip objects extend
    /// </summary>
    public abstract class GUIToolStripRenderedObject : GUIToolStripObject
    {
        /// <summary>
        /// Constructor used for rendered toolstrip objects
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        protected GUIToolStripRenderedObject(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Perform a single left mouse click in the middle of the toolstrip object
        /// </summary>
        public virtual void SingleClick()
        {
            SingleClick(MouseButton.Left);
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the toolstrip object
        /// </summary>
        /// <param name="button">The button to click</param>
        public virtual void SingleClick(MouseButton button)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description, LogItemType.Action);
            ParentToolStrip.SingleClickInternal(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2), button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double left mouse click in the middle of the toolstrip object
        /// </summary>
        public virtual void DoubleClick()
        {
            DoubleClick(MouseButton.Left);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the toolstrip object
        /// </summary>
        /// <param name="button">The button to double click</param>
        public virtual void DoubleClick(MouseButton button)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description, LogItemType.Action);
            ParentToolStrip.DoubleClickInternal(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2), button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Moves the mouse cursor to the middle of the toolstrip object
        /// </summary>
        public virtual void MoveTo()
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Move the mouse over the " + Identity.Description, LogItemType.Action);
            ParentToolStrip.MoveTo(bounds.X + (bounds.Width / 2), bounds.Y + (bounds.Height / 2));
        }

        /// <summary>
        /// Gets the toolstrip objects current text
        /// </summary>
        public virtual string Text
        {
            get
            {
                UpdateIndex();

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemText = GUI.m_APE.GetValueFromMessage();
                return itemText;
            }
        }

        /// <summary>
        /// Polls for the toolstrip object to have the specified text
        /// </summary>
        /// <param name="text">The text to wait for the toolstrip object to have</param>
        public void PollForText(string text)
        {
            UpdateIndex();

            //AddMessagePollMember will not work here
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (text == Text)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to match text within timeout for the " + Identity.Description);
                }

                Thread.Sleep(50);
            }
        }

        /// <summary>
        /// The tool tip text of the toolstrip object
        /// </summary>
        public virtual string ToolTipText
        {
            get
            {
                UpdateIndex();

                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ToolTipText", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemToolTipText = GUI.m_APE.GetValueFromMessage();

                return itemToolTipText;
            }
        }

        /// <summary>
        /// Gets the image of the toolstrip object
        /// </summary>
        /// <returns></returns>
        public virtual Image Image()
        {
            UpdateIndex();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Image", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            Image itemImage = GUI.m_APE.GetValueFromMessage();
            return itemImage;
        }

        internal Rectangle ItemBounds()
        {
            UpdateIndex();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "Bounds", MemberTypes.Property);
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

        internal IntPtr GetDropDown()
        {
            UpdateIndex();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "DropDown", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic handle = GUI.m_APE.GetValueFromMessage();

            if (handle == null)
            {
                throw new Exception("Could not find dropdown of " + Identity.Description);
            }

            return handle;
                
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip buttons
    /// </summary>
    public sealed class GUIToolStripButton : GUIToolStripRenderedObject
    {
        /// <summary>
        /// Constructor used for toolstrip buttons
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripButton(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip labels
    /// </summary>
    public sealed class GUIToolStripLabel : GUIToolStripRenderedObject
    {
        /// <summary>
        /// Constructor used for toolstrip labels
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripLabel(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip split buttons
    /// </summary>
    public sealed class GUIToolStripSplitButton : GUIToolStripRenderedObject
    {
        MenuUtils m_MenuUtils = new MenuUtils();

        /// <summary>
        /// Constructor used for toolstrip split buttons
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripSplitButton(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Selects the specified item in the split drop down
        /// </summary>
        /// <param name="splitDropDownItem">The item to select from the split drop down</param>
        public void SingleClickItem(string splitDropDownItem)
        {
            Rectangle bounds = ItemBounds();
            GUI.Log("Single " + MouseButton.Left.ToString() + " click on " + Identity.Description, LogItemType.Action);
            ParentToolStrip.SingleClickInternal(bounds.Width - 3, bounds.Y + (bounds.Height / 2), MouseButton.Left, MouseKeyModifier.None);

            GUI.Log("Select [" + splitDropDownItem + "] from " + Identity.Description, LogItemType.Action);

            string[] DropDownItems = splitDropDownItem.Split(new char[] { '\\' });
            int MenuIndex = 0;
            IntPtr Handle= GetDropDown();

            Input.Block(ParentToolStrip.ParentForm.Handle, Identity.Handle);
            try
            {
                for (int Item = 0; Item < DropDownItems.Length; Item++)
                {
                    if (Item > 0)
                    {
                        Handle = m_MenuUtils.GetDropDown(ParentToolStrip.ParentForm.Handle, Handle, MenuIndex);
                    }

                    MenuIndex = m_MenuUtils.GetIndexOfMenuItem(ParentToolStrip.ParentForm.Handle, Handle, DropDownItems[Item]);
                    m_MenuUtils.ClickMenuItem(ParentToolStrip.ParentForm.Handle, Handle, MenuIndex, DropDownItems[Item], ref Identity);
                }
            }
            finally
            {
                Input.Unblock();
            }
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip menus
    /// </summary>
    public sealed class GUIToolStripMenu : GUIToolStripDropDownButton
    {
        /// <summary>
        /// Constructor used for toolstrip menus
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripMenu(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip drop down buttons
    /// </summary>
    public class GUIToolStripDropDownButton : GUIToolStripRenderedObject
    {
        MenuUtils m_MenuUtils = new MenuUtils();

        /// <summary>
        /// Constructor used for toolstrip drop down buttons
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripDropDownButton(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Selects the specified item in the drop down
        /// </summary>
        /// <param name="dropDownItem">The item to select from the drop down</param>
        public void SingleClickItem(string dropDownItem)
        {
            this.SingleClick(MouseButton.Left);

            GUI.Log("Select [" + dropDownItem + "] from " + Identity.Description, LogItemType.Action);

            string[] DropDownItems = dropDownItem.Split(new char[] { '\\' });
            int MenuIndex = 0;
            IntPtr Handle = GetDropDown();

            Input.Block(ParentToolStrip.ParentForm.Handle, Identity.Handle);
            try
            {
                for (int Item = 0; Item < DropDownItems.Length; Item++)
                {
                    if (Item > 0)
                    {
                        Handle = m_MenuUtils.GetDropDown(ParentToolStrip.ParentForm.Handle, Handle, MenuIndex);
                    }

                    MenuIndex = m_MenuUtils.GetIndexOfMenuItem(ParentToolStrip.ParentForm.Handle, Handle, DropDownItems[Item]);
                    m_MenuUtils.ClickMenuItem(ParentToolStrip.ParentForm.Handle, Handle, MenuIndex, DropDownItems[Item], ref Identity);
                }
            }
            finally
            {
                Input.Unblock();
            }
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip separators
    /// </summary>
    public sealed class GUIToolStripSeparator : GUIToolStripRenderedObject
    {
        /// <summary>
        /// Constructor used for toolstrip separators
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripSeparator(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip comboboxes
    /// </summary>
    public sealed class GUIToolStripComboBox : GUIToolStripObject
    {
        private GUIComboBox ComboBox;

        /// <summary>
        /// Constructor used for toolstrip comboboxes
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripComboBox(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
            IntPtr comboBoxHandle = ItemComboBoxHandle(descriptionOfControl, base.Identity);
            ComboBox = new GUIComboBox(parentToolStrip.ParentForm, descriptionOfControl, new Identifier(Identifiers.Handle, comboBoxHandle));
        }

        /// <summary>
        /// Checks if the specified item exists in the toolstrip combobox
        /// </summary>
        /// <param name="item">The item to check if it exists</param>
        /// <returns></returns>
        public bool ItemExists(string item)
        {
            return ComboBox.ItemExists(item);
        }

        /// <summary>
        /// Selects the specified item in the toolstrip combobox
        /// </summary>
        /// <param name="item">The item to select</param>
        public void SingleClickItem(string item)
        {
            ComboBox.SingleClickItem(item);
        }

        /// <summary>
        /// Sets the text portion of the toolstrip combobox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text portion of the combobox to</param>
        public void SetText(string text)
        {
            ComboBox.SetText(text);
        }

        internal IntPtr ItemComboBoxHandle(string descriptionOfControl, ControlIdentifier identity)
        {
            UpdateIndex();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ComboBox", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic comboBoxHandle = GUI.m_APE.GetValueFromMessage();

            if (comboBoxHandle != null)
            {
                return comboBoxHandle;
            }

            throw new Exception("Failed to find the combobox of " + Identity.Description);
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip textboxes
    /// </summary>
    public sealed class GUIToolStripTextBox : GUIToolStripObject
    {
        private GUITextBox TextBox;

        /// <summary>
        /// Constructor used for toolstrip textboxes
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripTextBox(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
            IntPtr textBoxHandle = ItemTextBoxHandle(descriptionOfControl, base.Identity);
            TextBox = new GUITextBox(parentToolStrip.ParentForm, descriptionOfControl, new Identifier(Identifiers.Handle, textBoxHandle));

        }

        /// <summary>
        /// Sets the text of the toolstrip textbox to the specified text by sending keystrokes
        /// </summary>
        /// <param name="text">The text to set the text of the textbox to</param>
        public void SetText(string text)
        {
            TextBox.SetText(text);
        }

        /// <summary>
        /// Perform a single left mouse click in the middle of the toolstrip textbox
        /// </summary>
        public void SingleClick()
        {
            SingleClick(MouseButton.Left);
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the toolstrip textbox
        /// </summary>
        /// <param name="button">The button to click</param>
        public void SingleClick(MouseButton button)
        {
            TextBox.SingleClick(button);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the toolstrip textbox
        /// </summary>
        /// <param name="x">How far from the left edge of the control to click the mouse</param>
        /// <param name="y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        public void SingleClick(int x, int y, MouseButton button)
        {
            TextBox.SingleClick(x, y, button);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the toolstrip textbox while pressing the specified key
        /// </summary>
        /// <param name="x">How far from the left edge of the control to click the mouse</param>
        /// <param name="y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        /// <param name="keys">The key to hold while clicking</param>
        public void SingleClick(int x, int y, MouseButton button, MouseKeyModifier keys)
        {
            TextBox.SingleClick(x, y, button, keys);
        }

        /// <summary>
        /// Perform a double left mouse click in the middle of the toolstrip textbox
        /// </summary>
        public void DoubleClick()
        {
            DoubleClick(MouseButton.Left);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the toolstrip textbox
        /// </summary>
        /// <param name="button">The button to double click</param>
        public void DoubleClick(MouseButton button)
        {
            TextBox.DoubleClick(button);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the toolstrip textbox
        /// </summary>
        /// <param name="x">How far from the left edge of the control to double click the mouse</param>
        /// <param name="y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        public void DoubleClick(int x, int y, MouseButton button)
        {
            TextBox.DoubleClick(x, y, button);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the toolstrip textbox while pressing the specified key
        /// </summary>
        /// <param name="x">How far from the left edge of the control to double click the mouse</param>
        /// <param name="y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        /// <param name="keys">The key to hold while double clicking</param>
        public void DoubleClick(int x, int y, MouseButton button, MouseKeyModifier keys)
        {
            TextBox.DoubleClick(x, y, button, keys);
        }

        internal IntPtr ItemTextBoxHandle(string descriptionOfControl, ControlIdentifier identity)
        {
            UpdateIndex();

            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "TextBox", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic textBoxHandle = GUI.m_APE.GetValueFromMessage();

            if (textBoxHandle != null)
            {
                return textBoxHandle;
            }

            throw new Exception("Failed to find the textbox of " + descriptionOfControl);
        }
    }

    /// <summary>
    /// Automation class used to automate toolstrip progressbars
    /// </summary>
    public sealed class GUIToolStripProgressBar : GUIToolStripObject
    {
        private GUIProgressBar ProgressBar;

        /// <summary>
        /// Constructor used for toolstrip progress bars
        /// </summary>
        /// <param name="parentToolStrip">The parent toolstrip the object belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIToolStripProgressBar(GUIToolStrip parentToolStrip, string descriptionOfControl, params Identifier[] identParams)
            : base(parentToolStrip, descriptionOfControl, identParams)
        {
            IntPtr progressBarHandle = ItemProgressBarHandle(descriptionOfControl, base.Identity);
            ProgressBar = new GUIProgressBar(parentToolStrip.ParentForm, descriptionOfControl, new Identifier(Identifiers.Handle, progressBarHandle));
        }

        /// <summary>
        /// Gets the minimum value of the progressbar
        /// </summary>
        public int Minimum
        {
            get
            {
                return ProgressBar.Minimum;
            }
        }

        /// <summary>
        /// Gets the maximum value of the progressbar
        /// </summary>
        public int Maximum
        {
            get
            {
                return ProgressBar.Maximum;
            }
        }

        /// <summary>
        /// Gets the current value of the progressbar
        /// </summary>
        public int Value
        {
            get
            {
                return ProgressBar.Value;
            }
        }

        internal IntPtr ItemProgressBarHandle(string descriptionOfControl, ControlIdentifier identity)
        {
            UpdateIndex();
            
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, ParentToolStrip.ParentForm.Handle, ParentToolStrip.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, Index));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "ProgressBar", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Handle", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
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
