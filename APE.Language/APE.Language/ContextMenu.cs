using System;
using System.Diagnostics;
using System.Threading;
using NM = APE.Native.NativeMethods;
using APE.Communication;
using System.Reflection;
using System.ComponentModel;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate context menu controls
    /// </summary>
    public abstract class GUIContextMenu : GUIFocusableObject
    {
        /// <summary>
        /// Constructor used for context menus
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        internal GUIContextMenu(string descriptionOfControl, params Identifier[] identParams)
            : base(descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Selects the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to select from the context menu</param>
        public void SingleClickItem(string contextMenuItem)
        {
            SingleClickItem(contextMenuItem, ItemIdentifier.Text);
        }

        /// <summary>
        /// Selects the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to select from the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        public void SingleClickItem(string contextMenuItem, ItemIdentifier itemIdentifier)
        {
            string logMessageAction = "Single Left click on the item " + contextMenuItem + " from the " + Identity.Description;
            SingleClickItemInternal(contextMenuItem, logMessageAction, itemIdentifier);
        }

        internal void CloseContextMenu(IntPtr handle)
        {
            GUI.Log("Press the left alt key", LogItemType.Action);

            // Close the menu
            Input.SendAltKey();

            WaitForControlToNotBeVisible();

            // Small sleep to let focus switch
            Thread.Sleep(20);
        }

        internal abstract void SingleClickItemInternal(string contextMenuItem, string logMessageAction, ItemIdentifier itemIdentifier);

        /// <summary>
        /// Checks the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to check from the context menu</param>
        public void CheckItem(string contextMenuItem)
        {
            CheckItem(contextMenuItem, ItemIdentifier.Text);
        }

        /// <summary>
        /// Checks the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to check from the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        public abstract void CheckItem(string contextMenuItem, ItemIdentifier itemIdentifier);

        /// <summary>
        /// Unchecks the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to uncheck from the context menu</param>
        public void UncheckItem(string contextMenuItem)
        {
            UncheckItem(contextMenuItem, ItemIdentifier.Text);
        }

        /// <summary>
        /// Unchecks the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to uncheck from the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        public abstract void UncheckItem(string contextMenuItem, ItemIdentifier itemIdentifier);

        /// <summary>
        /// Determines if the specified item in the context menu is enabled
        /// </summary>
        /// <param name="contextMenuItem">The item to get the enabled state of in the context menu</param>
        /// <returns>True if the item is enabled otherwise false</returns>
        public bool ItemIsEnabled(string contextMenuItem)
        {
            return ItemIsEnabled(contextMenuItem, ItemIdentifier.Text);
        }

        /// <summary>
        /// Determines if the specified item in the context menu is enabled
        /// </summary>
        /// <param name="contextMenuItem">The item to get the enabled state of in the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        /// <returns>True if the item is enabled otherwise false</returns>
        public abstract bool ItemIsEnabled(string contextMenuItem, ItemIdentifier itemIdentifier);

        /// <summary>
        /// Determines if the specified item in the context menu is checked
        /// </summary>
        /// <param name="contextMenuItem">The item to get the checked state of in the context menu</param>
        /// <returns>True if the item is checked otherwise false</returns>
        public bool ItemIsChecked(string contextMenuItem)
        {
            return ItemIsChecked(contextMenuItem, ItemIdentifier.Text);
        }

        /// <summary>
        /// Determines if the specified item in the context menu is checked
        /// </summary>
        /// <param name="contextMenuItem">The item to get the checked state of in the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        /// <returns>True if the item is checked otherwise false</returns>
        public abstract bool ItemIsChecked(string contextMenuItem, ItemIdentifier itemIdentifier);
    }


    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.ContextMenuStrip
    /// </summary>
    [EditorBrowsable(EditorBrowsableState.Never)]
    public sealed class GUIContextMenuStrip : GUIContextMenu
    {
        private MenuUtils m_MenuUtils = new MenuUtils();

        /// <summary>
        /// Constructor used for context menus
        /// </summary>
        /// <param name="parentHandle">Parent form</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the identifier for the handle returned by GetContextMenu</param>
        internal GUIContextMenuStrip(IntPtr parentHandle, string descriptionOfControl, params Identifier[] identParams)
            : base(descriptionOfControl, new Identifier(Identifiers.Handle, GetContextMenu(parentHandle, descriptionOfControl)))
        {
        }

        internal static bool ContextMenuExists(IntPtr parent)
        {
            IntPtr handle = GetContextMenuInternal(parent);
            if (handle == IntPtr.Zero)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        internal override void SingleClickItemInternal(string contextMenuItem, string logMessageAction, ItemIdentifier itemIdentifier)
        {
            GUI.Log(logMessageAction, LogItemType.Action);

            string[] menus = contextMenuItem.Split(GUI.MenuDelimiterAsArray, StringSplitOptions.None);
            int menuIndex = 0;
            IntPtr handle = Identity.Handle;

            Input.Block();
            try
            {
                bool hasDropDownItems = false;
                for (int item = 0; item < menus.Length; item++)
                {
                    if (item > 0)
                    {
                        handle = m_MenuUtils.GetDropDown(Identity.ParentHandle, handle, menuIndex);
                    }
                    menuIndex = m_MenuUtils.GetIndexOfMenuItem(Identity.ParentHandle, handle, menus[item], itemIdentifier);
                    hasDropDownItems = m_MenuUtils.HasDropDownItems(handle, handle, menuIndex);
                    m_MenuUtils.ClickMenuItem(handle, handle, Description, menuIndex, menus[item]);
                }

                if (hasDropDownItems)
                {
                    CloseContextMenu(handle);
                }
            }
            catch
            {
                Input.Reset();  //Reset the mouse blocking
                throw;
            }
            finally
            {
                Input.Unblock();
            }
        }

        private static IntPtr GetContextMenu(IntPtr parent, string description)
        {
            IntPtr contextMenuHandle = IntPtr.Zero;

            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                contextMenuHandle = GetContextMenuInternal(parent);

                if (contextMenuHandle != IntPtr.Zero)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException("Failed to find context menu of the " + description);
                }

                Thread.Sleep(15);
            }
            return contextMenuHandle;
        }

        private static IntPtr GetContextMenuInternal(IntPtr parent)
        {
            //Get the ContextMenuStrip handle
            GUI.m_APE.AddFirstMessageGetContextMenuStrip(parent);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            IntPtr contextMenuHandle = GUI.m_APE.GetValueFromMessage();
            return contextMenuHandle;
        }

        /// <summary>
        /// Checks the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to check from the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        public override void CheckItem(string contextMenuItem, ItemIdentifier itemIdentifier)
        {
            bool isChecked = ItemIsChecked(contextMenuItem, itemIdentifier);

            if (isChecked)
            {
                GUI.Log("Ensure item " + contextMenuItem + "from the " + Identity.Description + " is checked", LogItemType.Action);
            }
            else
            {
                string logMessageAction = "Check item " + contextMenuItem + " from the " + Identity.Description;
                SingleClickItemInternal(contextMenuItem, logMessageAction, itemIdentifier);

                //Check it has been checked
                isChecked = ItemIsChecked(contextMenuItem, itemIdentifier);

                if (!isChecked)
                {
                    throw GUI.ApeException("Failed to check item " + contextMenuItem + " from the " + Identity.Description);
                }
            }
        }

        /// <summary>
        /// Unchecks the specified item in the context menu by clicking on it
        /// </summary>
        /// <param name="contextMenuItem">The item to uncheck from the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        public override void UncheckItem(string contextMenuItem, ItemIdentifier itemIdentifier)
        {
            bool isChecked = ItemIsChecked(contextMenuItem, itemIdentifier);

            if (!isChecked)
            {
                GUI.Log("Ensure item " + contextMenuItem + "from the " + Identity.Description + " is unchecked", LogItemType.Action);
            }
            else
            {
                string logMessageAction = "Uncheck item " + contextMenuItem + " from the " + Identity.Description;
                SingleClickItemInternal(contextMenuItem, logMessageAction, itemIdentifier);

                //Check it has been unchecked
                isChecked = ItemIsChecked(contextMenuItem, itemIdentifier);

                if (isChecked)
                {
                    throw GUI.ApeException("Failed to uncheck item " + contextMenuItem + " from the " + Identity.Description);
                }
            }
        }

        /// <summary>
        /// Determines if the specified item in the context menu is enabled
        /// </summary>
        /// <param name="contextMenuItem">The item to get the enabled state of in the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        /// <returns>True if the item is enabled otherwise false</returns>
        public override bool ItemIsEnabled(string contextMenuItem, ItemIdentifier itemIdentifier)
        {
            if (!IsEnabled)
            {
                return false;
            }

            string[] menus = contextMenuItem.Split(GUI.MenuDelimiterAsArray, StringSplitOptions.None);
            int menuIndex = 0;
            IntPtr handle = Identity.Handle;
            bool isEnabled = false;
            
            for (int item = 0; item < menus.Length; item++)
            {
                if (item > 0)
                {
                    handle = m_MenuUtils.GetDropDown(Identity.ParentHandle, handle, menuIndex);
                }
                menuIndex = m_MenuUtils.GetIndexOfMenuItem(Identity.ParentHandle, handle, menus[item], itemIdentifier);
                isEnabled = m_MenuUtils.MenuItemEnabled(handle, handle, menuIndex, menus[item]);

                if (!isEnabled)
                {
                    break;
                }
            }
            return isEnabled;
        }

        /// <summary>
        /// Determines if the specified item in the context menu is checked
        /// </summary>
        /// <param name="contextMenuItem">The item to get the checked state of in the context menu</param>
        /// <param name="itemIdentifier">The property to identify the item by</param>
        /// <returns>True if the item is checked otherwise false</returns>
        public override bool ItemIsChecked(string contextMenuItem, ItemIdentifier itemIdentifier)
        {
            string[] menus = contextMenuItem.Split(GUI.MenuDelimiterAsArray, StringSplitOptions.None);
            int menuIndex = 0;
            IntPtr handle = Identity.Handle;
            bool isChecked = false;

            for (int item = 0; item < menus.Length; item++)
            {
                if (item > 0)
                {
                    handle = m_MenuUtils.GetDropDown(Identity.ParentHandle, handle, menuIndex);
                }
                menuIndex = m_MenuUtils.GetIndexOfMenuItem(Identity.ParentHandle, handle, menus[item], itemIdentifier);
                if (item == contextMenuItem.Length - 1)
                {
                    isChecked = m_MenuUtils.MenuItemChecked(handle, handle, menuIndex, menus[item]);
                }
            }
            return isChecked;
        }
    }
}
