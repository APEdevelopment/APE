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
    /// System.Windows.Forms.MenuStrip
    /// </summary>
    public sealed class GUIMenuStrip : GUIToolStrip
    {
        MenuUtils m_MenuUtils = new MenuUtils();

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIMenuStrip(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Selects the specified menu item
        /// </summary>
        /// <param name="MenuItem">The text of the menu path to select, sub menu items are delimited by the \ character for instance File\Exit</param>
        public void Select(string MenuItem)
        {
            GUI.Log("Select [" + MenuItem + "] from " + m_DescriptionOfControl, LogItemTypeEnum.Action);

            string[] Menus = MenuItem.Split(new char[] { '\\' });
            int MenuIndex = 0;
            IntPtr Handle;

            Handle = Identity.Handle;

            Input.Block(Identity.ParentHandle, Identity.Handle);
            try
            {
                for (int Item = 0; Item < Menus.Length; Item++)
                {
                    if (Item > 0)
                    {
                        Handle = m_MenuUtils.GetDropDown(Identity.ParentHandle, Handle, MenuIndex);
                    }

                    MenuIndex = m_MenuUtils.GetIndexOfMenuItem(Identity.ParentHandle, Handle, Menus[Item]);
                    m_MenuUtils.ClickMenuItem(Identity.ParentHandle, Handle, MenuIndex, Menus[Item], ref Identity);
                }
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
