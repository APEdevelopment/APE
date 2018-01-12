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
using System.Windows.Forms;

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
        /// Gets the background image of the picturebox
        /// </summary>
        /// <returns>The background image</returns>
        public Image BackgroundImage()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "BackgroundImage", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            Image theBackgroundImage = GUI.m_APE.GetValueFromMessage();
            return theBackgroundImage;
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
                    throw new Exception(this.Description + " failed to become nonvisible");
                }

                if (!NM.IsWindowVisible(this.Handle))
                {
                    break;
                }
                
                Thread.Sleep(15);
            }

            // Small sleep to let focus switch
            Thread.Sleep(20);
        }

        

        /// <summary>
        /// Gets the currently displayed context menu of the control
        /// </summary>
        /// <returns>The context menu</returns>
        public GUIContextMenu GetContextMenu()
        {
            //TODO implement support for ContextMenu and native context menus (as well as the ContextMenuStrip thats currently implemented)
            if (GUIContextMenuStrip.ContextMenuExists(Identity.ParentHandle, Identity.Handle))
            {
                
                GUIContextMenuStrip contextMenu = new GUIContextMenuStrip(Identity.ParentHandle, Description + " conext menu", new Identifier(Identifiers.Handle, Identity.Handle));
                return contextMenu;
            }
            throw new Exception("Failed to find context menu");
        }

        /// <summary>
        /// Deprecated will be removed in a future version of APE, use ContextMenu instead.  
        /// Selects the specified item on the currently displayed context menu
        /// </summary>
        /// <param name="contextMenuItem">The text of the menu path to select, sub menu items are delimited by the \ character for instance File\Exit</param>
        [EditorBrowsable(EditorBrowsableState.Never)]
        public void ContextMenuSelect(string contextMenuItem)
        {
            GUI.Log("ContextMenuSelect is deprecated, please use GetContextMenu().SingleClickItem instead", LogItemType.Warning);
            GUIContextMenuStrip contextMenu = new GUIContextMenuStrip(Identity.ParentHandle, Description + " conext menu", new Identifier(Identifiers.Handle, Identity.Handle));
            contextMenu.SingleClickItem(contextMenuItem);
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
        public void MoveTo(int X, int Y)
        {
            Input.ClickCommon(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y);
        }

        /// <summary>
        /// Perform a left mouse single click in the middle of the control
        /// </summary>
        public virtual void SingleClick()
        {
            SingleClick(-1, -1, MouseButton.Left, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to click</param>
        public virtual void SingleClick(MouseButton button)
        {
            SingleClick(-1, -1, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        public virtual void SingleClick(int X, int Y, MouseButton button)
        {
            SingleClick(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        /// <param name="keys">The key to hold while clicking</param>
        public virtual void SingleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log("Single " + button.ToString() + " click on the " + Identity.Description + point + keyModifiers, LogItemType.Action);
            SingleClickInternal(X, Y, button, keys);
        }

        internal void SingleClickInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseSingleClick(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y, button, keys);
        }

        /// <summary>
        /// Perform a left mouse double click in the middle of the control
        /// </summary>
        public virtual void DoubleClick()
        {
            DoubleClick(-1, -1, MouseButton.Left, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to double click</param>
        public virtual void DoubleClick(MouseButton button)
        {
            DoubleClick(-1, -1, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        public virtual void DoubleClick(int X, int Y, MouseButton button)
        {
            DoubleClick(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        /// <param name="keys">The key to hold while double clicking</param>
        public virtual void DoubleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log("Double " + button.ToString() + " click on the " + Identity.Description + point + keyModifiers, LogItemType.Action);
            DoubleClickInternal(X, Y, button, keys);
        }

        internal void DoubleClickInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseDoubleClick(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y, button, keys);
        }

        /// <summary>
        /// Perform a left mouse triple click in the middle of the control
        /// </summary>
        public virtual void TripleClick()
        {
            TripleClick(-1, -1, MouseButton.Left, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to triple click</param>
        public virtual void TripleClick(MouseButton button)
        {
            TripleClick(-1, -1, button, MouseKeyModifier.None);
        }

        /// <summary>
        ///  Perform a triple mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="button">The button to triple click</param>
        public virtual void TripleClick(int X, int Y, MouseButton button)
        {
            TripleClick(X, Y, button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="button">The button to triple click</param>
        /// <param name="keys">The key to hold while triple clicking</param>
        public virtual void TripleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(keys);

            GUI.Log("Triple " + button.ToString() + " click on the " + Identity.Description + point + keyModifiers, LogItemType.Action);
            TripleClickInternal(X, Y, button, keys);
        }

        internal void TripleClickInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseTripleClick(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y, button, keys);
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

            GUI.Log(button.ToString() + " mouse down on the " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseDownInternal(X, Y, button, keys);
        }

        internal void MouseDownInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseDown(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y, button, keys);
            if (!Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception(Identity.Description + " did not go idle within timeout");
            }

            // If we are doing separate calls to mouse down and up then its very likely we want to drag so make sure we are in dragmode
            NM.tagRect WindowRect;
            NM.tagRect ClientRect;
            NM.GetWindowRect(Handle, out WindowRect);
            NM.GetClientRect(Handle, out ClientRect);

            int middleOfClientAreaX = (ClientRect.right / 2);
            int middleOfClientAreaY = (ClientRect.bottom / 2);

            int dragTriggerHeight = SystemInformation.DragSize.Height;
            int dragTriggerWidth = SystemInformation.DragSize.Width;

            if (middleOfClientAreaX < dragTriggerWidth || middleOfClientAreaY < dragTriggerHeight)
            {
                throw new Exception(Description + " is to small to reliably enter drag mode");
            }

            // Move the mouse a few times to make sure we are in drag mode
            for (int i = 0; i < 10; i++)
            {
                Input.MoveMouse(WindowRect.left, WindowRect.top);
                Input.MoveMouse(WindowRect.left + middleOfClientAreaX, WindowRect.top + middleOfClientAreaY);
            }
            if (!Input.WaitForInputIdle(Identity.Handle, GUI.m_APE.TimeOut))
            {
                throw new Exception(Identity.Description + " did not go idle within timeout");
            }
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

            GUI.Log(button.ToString() + " mouse up on the " + Identity.Description + point + keyModifiers, LogItemType.Action);
            MouseUpInternal(X, Y, button, keys);
        }

        internal void MouseUpInternal(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            Input.MouseUp(Identity.ParentHandle, Identity.Handle, Identity.Description, X, Y, button, keys);
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
                // Try to scroll it into view
                GUI.m_APE.AddFirstMessageScrollControlIntoView(Handle);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);

                NM.tagRect clipBox = NM.GetClipBox(Handle);
                return clipBox.bottom;
                //NM.tagRect WindowSize;
                //NM.GetClientRect(Identity.Handle, out WindowSize);
                //return WindowSize.bottom;
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
                // Try to scroll it into view
                GUI.m_APE.AddFirstMessageScrollControlIntoView(Handle);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);

                NM.tagRect clipBox = NM.GetClipBox(Handle);
                return clipBox.right;
                //NM.tagRect WindowSize;
                //NM.GetClientRect(Identity.Handle, out WindowSize);
                //return WindowSize.right;
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
