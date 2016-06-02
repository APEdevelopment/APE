using System;
using System.Collections.Generic;
using System.Text;
using System.ComponentModel;
using System.Threading;
using System.Diagnostics;
using System.Drawing;
using APE.Communication;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Abstract base object which all other winforms object extend
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public abstract class GUIObject
    {
        /// <summary>
        /// Gets the identity of the control
        /// </summary>
        protected ControlIdentifier Identity;
        private MenuUtils m_MenuUtils = new MenuUtils();
        private AnimationUtils m_AnimationUtils = new AnimationUtils();
        internal GUIForm m_ParentForm = null;
        internal string m_DescriptionOfControl = null;

        /// <summary>
        /// Constructor used for form controls
        /// </summary>
        protected GUIObject(string descriptionOfControl, params Identifier[] identParams)
            : this(null, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Constructor used for everything which isn't a Form
        /// </summary>
        protected GUIObject(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
        {
            m_DescriptionOfControl = descriptionOfControl;
            if (parentForm != null)
            {
                m_ParentForm = parentForm;
                Identity.ParentHandle = m_ParentForm.Handle;
            }

            foreach (Identifier i in identParams)
            {
                switch (i.IdentifierType)
                {
                    case Identifiers.Handle:
                        Identity.Handle = i.IdentifierValue;
                        break;
                    case Identifiers.Name:
                        Identity.Name = i.IdentifierValue;
                        break;
                    case Identifiers.TechnologyType:
                        Identity.TechnologyType = i.IdentifierValue;
                        break;
                    case Identifiers.TypeNameSpace:
                        Identity.TypeNameSpace = i.IdentifierValue;
                        break;
                    case Identifiers.TypeName:
                        Identity.TypeName = i.IdentifierValue;
                        break;
                    case Identifiers.ModuleName:
                        Identity.ModuleName = i.IdentifierValue;
                        break;
                    case Identifiers.AssemblyName:
                        Identity.AssemblyName = i.IdentifierValue;
                        break;
                    case Identifiers.Index:
                        Identity.Index = i.IdentifierValue;
                        break;
                    case Identifiers.Text:
                        Identity.Text = i.IdentifierValue;
                        break;
                }
            }

            switch (Identity.TechnologyType)
            {
                case "Windows Native":
                    if (Identity.Handle == IntPtr.Zero)
                    {
                        Stopwatch timer = Stopwatch.StartNew();
                        do
                        {
                            if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                            {
                                throw new Exception("Failed to find native control");
                            }

                            Identity.Handle = NM.FindWindow(null, Identity.Text);

                            if (Identity.Handle != IntPtr.Zero)
                            {
                                if (NM.IsWindowVisible(Identity.Handle))
                                {
                                    break;
                                }
                            }

                            Thread.Sleep(15);
                        }
                        while (true);
                    }
                    else
                    {
                        if (!NM.IsWindowVisible(Identity.Handle))
                        {
                            throw new Exception("Failed to find native control");
                        }
                    }
                    break;
                default:
                    GUI.m_APE.AddMessageFindByProperty(Identity);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);

                    GUI.m_APE.DecodeControl(1, out Identity);
                    break;
            }
        }

        /// <summary>
        /// Selects the specified item on the currently displayed context menu
        /// </summary>
        /// <param name="ContextMenuItem">The text of the menu path to select, sub menu items are delimited by the \ character for instance File\Exit</param>
        public void ContextMenuSelect(string ContextMenuItem)
        {
            GUI.Log("Select context menu item [" + ContextMenuItem + "]", LogItemTypeEnum.Action);

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

            dynamic ContextMenuHandle;

            //If all controls set the contextmenu / contextmenustrip property we could find it using the below
            //but some don't so we use an alternative method to find them
            ////Get the ContextMenuStrip handle
            //GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Parent, Control);
            //GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "ContextMenuStrip", MemberTypes.Property);
            //GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Handle", MemberTypes.Property);
            //GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            //GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            //GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            ////get the values returned
            //ContextMenuHandle = GUI.m_APE.GetValueFromMessage(1);

            Stopwatch timer = Stopwatch.StartNew();
            do
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to find context menu");
                }

                //Get the ContextMenuStrip handle
                GUI.m_APE.AddMessageGetContextMenuStrip(Control);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //get the values returned
                ContextMenuHandle = GUI.m_APE.GetValueFromMessage(1);

                Thread.Sleep(15);
            }
            while (ContextMenuHandle == null || ContextMenuHandle == IntPtr.Zero);

            return ContextMenuHandle;
        }

        /// <summary>
        /// Gets the windows classname (not that helpful in the .NET world)
        /// </summary>
        public string ClassName
        {
            get
            {
                StringBuilder ClassName = new StringBuilder(1024);
                NM.GetClassName(Identity.Handle, ClassName, ClassName.Capacity - 1);
                return ClassName.ToString();
            }
        }

        /// <summary>
        /// Moves the mouse cursor to the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to move the mouse</param>
        /// <param name="Y">How far from the top edge of the control to move the mouse</param>
        public void MouseMove(int X, int Y)
        {
            GUI.m_APE.MouseMove(Identity.Handle, X, Y);
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="Button">The button to click</param>
        public virtual void MouseSingleClick(MouseButton Button)
        {
            MouseSingleClick(-1, -1, Button, MouseKeyModifier.None);
        }

        /// <summary>
        ///  Perform a mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="Button">The button to click</param>
        public virtual void MouseSingleClick(int X, int Y, MouseButton Button)
        {
            MouseSingleClick(X, Y, Button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="Button">The button to click</param>
        /// <param name="Keys">The key to hold while clicking</param>
        public virtual void MouseSingleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(Keys);

            GUI.Log("Single " + Button.ToString() + " click on " + m_DescriptionOfControl + point + keyModifiers, LogItemTypeEnum.Action);
            MouseSingleClickInternal(X, Y, Button, Keys);
        }

        internal void MouseSingleClickInternal(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            Input.MouseSingleClick(Identity.ParentHandle, Identity.Handle, X, Y, Button, Keys);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="Button">The button to double click</param>
        public virtual void MouseDoubleClick(MouseButton Button)
        {
            MouseDoubleClick(-1, -1, Button, MouseKeyModifier.None);
        }

        /// <summary>
        ///  Perform a double mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="Button">The button to double click</param>
        public virtual void MouseDoubleClick(int X, int Y, MouseButton Button)
        {
            MouseDoubleClick(X, Y, Button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="Button">The button to double click</param>
        /// <param name="Keys">The key to hold while double clicking</param>
        public virtual void MouseDoubleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(Keys);

            GUI.Log("Double " + Button.ToString() + " click on " + m_DescriptionOfControl + point + keyModifiers, LogItemTypeEnum.Action);
            MouseDoubleClickInternal(X, Y, Button, Keys);
        }

        internal void MouseDoubleClickInternal(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            Input.MouseDoubleClick(Identity.ParentHandle, Identity.Handle, X, Y, Button, Keys);
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="Button">The button to triple click</param>
        public virtual void MouseTripleClick(MouseButton Button)
        {
            MouseTripleClick(-1, -1, Button, MouseKeyModifier.None);
        }

        /// <summary>
        ///  Perform a triple mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="Button">The button to triple click</param>
        public virtual void MouseTripleClick(int X, int Y, MouseButton Button)
        {
            MouseTripleClick(X, Y, Button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="Button">The button to triple click</param>
        /// <param name="Keys">The key to hold while triple clicking</param>
        public virtual void MouseTripleClick(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(Keys);

            GUI.Log("Triple " + Button.ToString() + " click on " + m_DescriptionOfControl + point + keyModifiers, LogItemTypeEnum.Action);
            MouseTripleClickInternal(X, Y, Button, Keys);
        }

        internal void MouseTripleClickInternal(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            Input.MouseTripleClick(Identity.ParentHandle, Identity.Handle, X, Y, Button, Keys);
        }

        /// <summary>
        /// Perform a mouse down with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse down</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse down</param>
        /// <param name="Button">The button to press</param>
        public virtual void MouseDown(int X, int Y, MouseButton Button)
        {
            MouseDown(X, Y, Button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse down with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse down</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse down</param>
        /// <param name="Button">The button to press</param>
        /// <param name="Keys">The key to hold while performing a mouse down</param>
        public virtual void MouseDown(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(Keys);

            GUI.Log(Button.ToString() + " mouse down on " + m_DescriptionOfControl + point + keyModifiers, LogItemTypeEnum.Action);
            MouseDownInternal(X, Y, Button, Keys);
        }

        internal void MouseDownInternal(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            Input.MouseDown(Identity.ParentHandle, Identity.Handle, X, Y, Button, Keys);
        }

        /// <summary>
        /// Perform a mouse up with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse up</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse up</param>
        /// <param name="Button">The button to release</param>
        public void MouseUp(int X, int Y, MouseButton Button)
        {
            MouseUp(X, Y, Button, MouseKeyModifier.None);
        }

        /// <summary>
        /// Perform a mouse up with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse up</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse up</param>
        /// <param name="Button">The button to release</param>
        /// <param name="Keys">The key to hold while performing a mouse up</param>
        public void MouseUp(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            string point = GetPointText(X, Y);
            string keyModifiers = GetKeyModifierText(Keys);

            GUI.Log(Button.ToString() + " mouse up on " + m_DescriptionOfControl + point + keyModifiers, LogItemTypeEnum.Action);
            MouseUpInternal(X, Y, Button, Keys);
        }

        internal void MouseUpInternal(int X, int Y, MouseButton Button, MouseKeyModifier Keys)
        {
            Input.MouseUp(Identity.ParentHandle, Identity.Handle, X, Y, Button, Keys);
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
                NM.RECT WindowSize;
                NM.GetClientRect(Identity.Handle, out WindowSize);
                return WindowSize.Bottom;
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
                NM.RECT WindowSize;
                NM.GetWindowRect(Identity.Handle, out WindowSize);
                return WindowSize.Left;
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
        /// Gets the windows title bar current text
        /// </summary>
        public virtual string Text
        {
            get
            {
                //TODO does this work cross application?
                return GUI.m_APE.GetWindowText(Identity.Handle);
            }
        }

        /// <summary>
        /// Gets the topedge of the control relative to the screen
        /// </summary>
        public int Top
        {
            get
            {
                NM.RECT WindowSize;
                NM.GetWindowRect(Identity.Handle, out WindowSize);
                return WindowSize.Top;
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
                NM.RECT WindowSize;
                NM.GetClientRect(Identity.Handle, out WindowSize);
                return WindowSize.Right;
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
