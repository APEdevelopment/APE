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
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;
using System.Drawing;
using APE.Communication;
using NM = APE.Native.NativeMethods;
using System.Reflection;

namespace APE.Language
{
    /// <summary>
    /// Abstract base object which all other winforms objects extend
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class GUIObject
    {
        /// <summary>
        /// The identity of the control
        /// </summary>
        protected ControlIdentifier Identity;
        /// <summary>
        /// The parent form the control
        /// </summary>
        protected internal GUIForm ParentForm;

        ///// <summary>
        ///// The human readable description of the control
        ///// </summary>
        private MenuUtils m_MenuUtils = new MenuUtils();
        private AnimationUtils m_AnimationUtils = new AnimationUtils();

        /// <summary>
        /// Constructor used for form controls
        /// </summary>
        /// <param name="descriptionOfControl">The human readable description of the form</param>
        /// <param name="identParams">The identifier(s) of the form</param>
        protected GUIObject(string descriptionOfControl, params Identifier[] identParams)
            : this(null, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Constructor used for everything which isn't a Form
        /// </summary>
        /// <param name="parentForm">The form the control belongs to</param>
        /// <param name="descriptionOfControl">The human readable description of the control</param>
        /// <param name="identParams">The identifier(s) of the control</param>
        protected GUIObject(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
        {
            ParentForm = parentForm;
            Identity = GUI.BuildIdentity(parentForm, descriptionOfControl, identParams);

            GUI.m_APE.AddFirstMessageFindByProperty(Identity);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            GUI.m_APE.DecodeControl(1, out Identity);

            Input.WaitForInputIdle(Handle, GUI.m_APE.TimeOut);
        }

        /// <summary>
        /// Waits for the specified control to not be visible
        /// </summary>
        public void WaitForControlToNotBeVisible()
        {
            //Wait for the control to not be visible
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Control " + this.Description + " failed to become nonvisible");
                }

                if (!NM.IsWindowVisible(this.Handle))
                {
                    break;
                }
                
                Thread.Sleep(15);
            }
        }

        /// <summary>
        /// Selects the specified item on the currently displayed context menu
        /// </summary>
        /// <param name="ContextMenuItem">The text of the menu path to select, sub menu items are delimited by the \ character for instance File\Exit</param>
        public void ContextMenuSelect(string ContextMenuItem)
        {
            GUI.Log("Select context menu item [" + ContextMenuItem + "]", LogItemType.Action);

            string[] Menus = ContextMenuItem.Split(new char[] { '\\' });
            int MenuIndex = 0;
            IntPtr Handle;
            IntPtr MenuParent;

            Handle = GetContextMenu(Identity.ParentHandle, Identity.Handle);
            //WaitForAnimation(Handle, false);

            if (Input.ActiveWindow(Handle))
            {
                MenuParent = Handle;
            }
            else
            {
                MenuParent = Identity.ParentHandle;
            }

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                for (int Item = 0; Item < Menus.Length; Item++)
                {
                    if (Item > 0)
                    {
                        Handle = m_MenuUtils.GetDropDown(Identity.ParentHandle, Handle, MenuIndex);
                        //WaitForAnimation(Handle, false);
                    }
                    MenuIndex = m_MenuUtils.GetIndexOfMenuItem(Identity.ParentHandle, Handle, Menus[Item]);
                    m_MenuUtils.ClickMenuItem(MenuParent, Handle, MenuIndex, Menus[Item], ref Identity);
                }
            }
            finally
            {
                Input.Unblock();
            }
        }

        private IntPtr GetContextMenu(IntPtr Parent, IntPtr Control)
        {
            //TODO implement support for ContextMenu (as well as the ContextMenuStrip thats currently implemented)

            IntPtr contextMenuHandle = IntPtr.Zero;

            //If all controls set the contextmenu / contextmenustrip property we could find it using the below
            //but some don't so we use an alternative method to find them
            ////Get the ContextMenuStrip handle
            //GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
            //GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "ContextMenuStrip", MemberTypes.Property);
            //GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            //GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            //GUI.m_APE.SendMessages(EventSet.APE);
            //GUI.m_APE.WaitForMessages(EventSet.APE);
            ////get the values returned
            //contextMenuHandle = GUI.m_APE.GetValueFromMessage();

            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to find context menu");
                }

                //Get the ContextMenuStrip handle
                GUI.m_APE.AddFirstMessageGetContextMenuStrip(Control);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                //get the values returned
                contextMenuHandle = GUI.m_APE.GetValueFromMessage();

                Thread.Sleep(15);
            }
            while (contextMenuHandle == IntPtr.Zero);

            return contextMenuHandle;
        }

        /// <summary>
        /// Gets the windows classname (not that helpful in the .NET world)
        /// </summary>
        public string ClassName
        {
            get
            {
                return NM.GetClassName(Identity.Handle);
            }
        }

        /// <summary>
        /// The description of the control
        /// </summary>
        public string Description
        {
            get
            {
                return Identity.Description;
            }
        }

        /// <summary>
        /// Moves the mouse cursor to the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to move the mouse</param>
        /// <param name="Y">How far from the top edge of the control to move the mouse</param>
        public void MouseMove(int X, int Y)
        {
            Input.MouseMove(Identity.Handle, X, Y);
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to click</param>
        public virtual void MouseSingleClick(MouseButton button)
        {
            MouseSingleClick(-1, -1, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        public virtual void MouseSingleClick(int X, int Y, MouseButton button)
        {
            MouseSingleClick(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        /// <param name="keys">The key to hold while clicking</param>
        public virtual void MouseSingleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log("Single " + button.ToString() + " click on " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseSingleClickInternal(X, Y, button, keys);
        }

        internal void MouseSingleClickInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseSingleClick(Identity.ParentHandle, Identity.Handle, X, Y, button, keys);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to double click</param>
        public virtual void MouseDoubleClick(MouseButton button)
        {
            MouseDoubleClick(-1, -1, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        public virtual void MouseDoubleClick(int X, int Y, MouseButton button)
        {
            MouseDoubleClick(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        /// <param name="keys">The key to hold while double clicking</param>
        public virtual void MouseDoubleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log("Double " + button.ToString() + " click on " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseDoubleClickInternal(X, Y, button, keys);
        }

        internal void MouseDoubleClickInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseDoubleClick(Identity.ParentHandle, Identity.Handle, X, Y, button, keys);
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to triple click</param>
        public virtual void MouseTripleClick(MouseButton button)
        {
            MouseTripleClick(-1, -1, button, MouseKeyModifier.None);
        }

        /// <summary>
        ///  Perform a triple mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="button">The button to triple click</param>
        public virtual void MouseTripleClick(int X, int Y, MouseButton button)
        {
            MouseTripleClick(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="button">The button to triple click</param>
        /// <param name="keys">The key to hold while triple clicking</param>
        public virtual void MouseTripleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log("Triple " + button.ToString() + " click on " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseTripleClickInternal(X, Y, button, keys);
        }

        internal void MouseTripleClickInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseTripleClick(Identity.ParentHandle, Identity.Handle, X, Y, button, keys);
        }

        /// <summary>
        /// Perform a mouse down with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse down</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse down</param>
        /// <param name="button">The button to press</param>
        public virtual void MouseDown(int X, int Y, MouseButton button)
        {
            MouseDown(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse down with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse down</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse down</param>
        /// <param name="button">The button to press</param>
        /// <param name="keys">The key to hold while performing a mouse down</param>
        public virtual void MouseDown(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log(button.ToString() + " mouse down on " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseDownInternal(X, Y, button, keys);
        }

        internal void MouseDownInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseDown(Identity.ParentHandle, Identity.Handle, X, Y, button, keys);
        }

        /// <summary>
        /// Perform a mouse up with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse up</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse up</param>
        /// <param name="button">The button to release</param>
        public void MouseUp(int X, int Y, MouseButton button)
        {
            MouseUp(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse up with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse up</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse up</param>
        /// <param name="button">The button to release</param>
        /// <param name="keys">The key to hold while performing a mouse up</param>
        public void MouseUp(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log(button.ToString() + " mouse up on " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseUpInternal(X, Y, button, keys);
        }

        internal void MouseUpInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseUp(Identity.ParentHandle, Identity.Handle, X, Y, button, keys);
        }

        /// <summary>
        /// Gets whether the control is currently enabled
        /// </summary>
        public bool IsEnabled
        {
            get
            {
                return NM.IsWindowEnabled(Identity.Handle);
            }
        }

        /// <summary>
        /// Gets whether the control currently exists
        /// </summary>
        public bool Exists
        {
            get
            {
                return NM.IsWindow(Identity.Handle);
            }
        }

        /// <summary>
        /// Gets the extended window style of the control
        /// </summary>
        public long ExtendedStyle
        {
            get
            {
                return (long)NM.GetWindowLongPtr(Identity.Handle, NM.GWL.GWL_EXSTYLE);
            }
        }

        /// <summary>
        /// Gets the height of the control
        /// </summary>
        public int Height
        {
            get
            {
                NM.tagRect WindowSize;
                NM.GetClientRect(Identity.Handle, out WindowSize);
                return WindowSize.bottom;
            }
        }

        /// <summary>
        /// Gets the controls window handle
        /// </summary>
        public IntPtr Handle
        {
            get
            {
                return Identity.Handle;
            }
        }

        /// <summary>
        /// Gets the id of the control
        /// </summary>
        public int Id
        {
            get
            {
                return NM.GetDlgCtrlID(Identity.Handle);
            }
        }

        /// <summary>
        /// Gets the left edge of the control relative to the screen
        /// </summary>
        public int Left
        {
            get
            {
                NM.tagRect WindowSize;
                NM.GetWindowRect(Identity.Handle, out WindowSize);
                return WindowSize.left;
            }
        }

        /// <summary>
        /// Get the .Name property of the control
        /// </summary>
        public string Name
        {
            get
            {
                return Identity.Name;
            }
        }

        /// <summary>
        /// Gets the window style of the control
        /// </summary>
        public long Style
        {
            get
            {
                return (long)NM.GetWindowLongPtr(Identity.Handle, NM.GWL.GWL_STYLE);
            }
        }

        /// <summary>
        /// Gets the windows current text
        /// </summary>
        public virtual string Text
        {
            get
            {
                if (Identity.TechnologyType == "Windows Forms (WinForms)")
                {
                    //Get the Text property
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Text", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string text = GUI.m_APE.GetValueFromMessage();
                    return text;
                }
                else
                {
                    return GUI.m_APE.GetWindowTextViaWindowMessage(Identity.Handle);
                }
            }
        }

        /// <summary>
        /// Gets the topedge of the control relative to the screen
        /// </summary>
        public int Top
        {
            get
            {
                NM.tagRect WindowSize;
                NM.GetWindowRect(Identity.Handle, out WindowSize);
                return WindowSize.top;
            }
        }

        /// <summary>
        /// Gets the name of the .NET type of the control
        /// </summary>
        public string TypeName
        {
            get
            {
                return Identity.TypeName;
            }
        }

        /// <summary>
        /// Gets whether the control is currently visible
        /// </summary>
        public bool IsVisible
        {
            get
            {
                return NM.IsWindowVisible(Identity.Handle);
            }
        }

        /// <summary>
        /// Gets the width of the control
        /// </summary>
        public int Width
        {
            get
            {
                NM.tagRect WindowSize;
                NM.GetClientRect(Identity.Handle, out WindowSize);
                return WindowSize.right;
            }
        }

        /// <summary>
        /// Gets the technology type of the control, currently this can be one of the following:
        ///   Windows Forms (WinForms)
        ///   Windows Native
        /// </summary>
        public string TechnologyType
        {
            get
            {
                return Identity.TechnologyType;
            }
        }

        /// <summary>
        /// Gets the namespace of the .NET type of the control
        /// </summary>
        public string TypeNameSpace
        {
            get
            {
                return Identity.TypeNameSpace;
            }
        }

        /// <summary>
        /// Gets the name of the module the control is part of
        /// </summary>
        public string ModuleName
        {
            get
            {
                return Identity.ModuleName;
            }
        }

        /// <summary>
        /// Gets the name of the assembly the control is part of
        /// </summary>
        public string AssemblyName
        {
            get
            {
                return Identity.AssemblyName;
            }
        }

        internal void WaitForAnimation(IntPtr Handle, bool ClearClientArea, AnimationUtils.WaitForAnimationSource Source)
        {
            m_AnimationUtils.WaitForAnimation(Handle, ClearClientArea, Source);
        }

        private string GetPointText(int X, int Y)
        {
            string point = "";

            if (X == -1 && Y == -1)
            {
                point = "";
            }
            else
            {
                point = " at point " + X.ToString() + ", " + Y.ToString();
            }

            return point;
        }

        private string GetKeyModifierText(MouseKeyModifier Keys)
        {
            string keyModifiers = "";

            if (Keys.HasFlag(MouseKeyModifier.Control))
            {
                keyModifiers = " holding the control key";
            }

            if (Keys.HasFlag(MouseKeyModifier.Shift))
            {
                if (Keys.HasFlag(MouseKeyModifier.Control))
                {
                    keyModifiers += " and the shift key";
                }
                else
                {
                    keyModifiers = " holding the shift key";
                }
            }

            return keyModifiers;
        }
    }
}
