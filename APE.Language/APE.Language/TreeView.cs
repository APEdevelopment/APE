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
    /// System.Windows.Forms.TreeView
    /// </summary>
    public sealed class GUITreeView : GUIFocusableObject
    {
        private bool m_ShowPlusMinus = true;
        private bool m_CheckBoxes = false;

        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUITreeView(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        private IntPtr GetRootNode(string NodeText)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetNodeCount", MemberTypes.Method, new Parameter(GUI.m_APE, false));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store2, "ShowPlusMinus", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store3, "CheckBoxes", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int ChildNodes = GUI.m_APE.GetValueFromMessage();
            m_ShowPlusMinus = GUI.m_APE.GetValueFromMessage();
            m_CheckBoxes = GUI.m_APE.GetValueFromMessage();

            //if (CheckBoxes)
            //{
            //    using (Graphics g = GUI.CreateGraphics())
            //    {
            //        CheckBoxModifier = CheckBoxRenderer.GetGlyphSize(g, CheckBoxState.UncheckedNormal).Width;
            //    }
            //}

            for (int ChildItem = 0; ChildItem < ChildNodes; ChildItem++)
            {
                // Get the text of the root node
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Nodes", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, ChildItem));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "Handle", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string Text = GUI.m_APE.GetValueFromMessage();
                IntPtr NewNodeHandle = GUI.m_APE.GetValueFromMessage();

                if (Text == NodeText)
                {
                    return NewNodeHandle;
                }
            }

            throw new Exception("Failed to find root node");
        }

        private IntPtr GetChildNode(IntPtr Node, string NodeText)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "GetNodeCount", MemberTypes.Method, new Parameter(GUI.m_APE, false));
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int ChildNodes = GUI.m_APE.GetValueFromMessage();

            for (int ChildItem = 0; ChildItem < ChildNodes; ChildItem++)
            {
                // Get the text of the child node
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Nodes", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Item", MemberTypes.Property, new Parameter(GUI.m_APE, ChildItem));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Text", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Handle", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string Text = GUI.m_APE.GetValueFromMessage();
                IntPtr NewNodeHandle = GUI.m_APE.GetValueFromMessage();

                if (Text == NodeText)
                {
                    return NewNodeHandle;
                }
            }

            throw new Exception("Failed to find child node");
        }

        private int MakeLParam(int LoWord, int HiWord)
        {
            return ((HiWord << 16) | (LoWord & 0xffff));
        }

        private void ExpandNode(IntPtr Node)
        {
            int Position = 0;
            Stopwatch timer;
            int Left = 0;
            int Top = 0;
            int Right = 0;
            int Bottom = 0;
            NM.tagRect ClientRect = new NM.tagRect();

            // check if already expanded
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "IsExpanded", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool IsExpanded = GUI.m_APE.GetValueFromMessage();

            if (IsExpanded)
            {
                return;
            }

            for (int i = 0; i < 2; i++)
            {
                // Get bounds
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "EnsureVisible", MemberTypes.Method);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddMessageGetValue(DataStores.Store7);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                Left = GUI.m_APE.GetValueFromMessage();
                Top = GUI.m_APE.GetValueFromMessage();
                Right = GUI.m_APE.GetValueFromMessage();
                Bottom = GUI.m_APE.GetValueFromMessage();

                NM.GetClientRect(Identity.Handle, out ClientRect);

                if (Left < 0 || Top < 0 || Bottom > ClientRect.bottom)
                {
                    // Go round again
                }
                else
                {
                    break;
                }
            }

            if (Right > ClientRect.right)
            {
                Right = ClientRect.right;
            }

            bool WaitForMove = false;
            if (Bottom > ClientRect.bottom)
            {
                WaitForMove = true;
                Bottom = ClientRect.bottom;
            }

            // Expand it
            if (m_ShowPlusMinus)
            {
                int CheckBoxModifier = 0;
                if (m_CheckBoxes)
                {
                    CheckBoxModifier = 12;
                }

                // Scroll the expand icon in to view if need be
                int X = Left - 7 - CheckBoxModifier - 5;
                if (X < 0)
                {
                    //get current scroll bar pos
                    Position = NM.GetScrollPos(Identity.Handle, NM.SBS_HORZ);

                    //update the pos
                    Position = Position + X;
                    NM.SetScrollPos(Identity.Handle, NM.SBS_HORZ, Position, false);

                    //send the message to move
                    IntPtr MessageResult;
                    IntPtr SendResult = NM.SendMessageTimeout(Handle, NM.WM_HSCROLL, new IntPtr(MakeLParam((int)NM.SB_THUMBPOSITION, Position)), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, (uint)GUI.GetTimeOuts(), out MessageResult);

                    base.MouseSingleClickInternal(5, Top + ((Bottom - Top) / 2), MouseButton.Left, MouseKeyModifier.None);
                }
                else
                {
                    base.MouseSingleClickInternal(Left - 7 - CheckBoxModifier, Top + ((Bottom - Top) / 2), MouseButton.Left, MouseKeyModifier.None);
                }
            }
            else
            {
                base.MouseDoubleClickInternal(Right - 4, Top + ((Bottom - Top) / 2), MouseButton.Left, MouseKeyModifier.None);
            }

            // Wait for it to be expanded
            bool IsNowExpanded = false;
            timer = Stopwatch.StartNew();
            do
            {
                // Get if expanded
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "IsExpanded", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                IsNowExpanded = GUI.m_APE.GetValueFromMessage();

                if (IsNowExpanded)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Node failed to expand");
                }

                Thread.Sleep(15);
            }
            while (true);

            //left clicking on a node which is partially off the bottom of the window will scroll it
            if (WaitForMove)
            {
                timer = Stopwatch.StartNew();
                do
                {
                    // Get bounds
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Bounds", MemberTypes.Property);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Bottom", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    // Get the value(s) returned MUST be done straight after the WaitForMessages call
                    Bottom = GUI.m_APE.GetValueFromMessage();

                    if (Bottom >= ClientRect.bottom)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Node failed to scroll fully into view");
                    }

                    Thread.Sleep(15);
                }
                while (true);
            }

            // After clicking sleep for a bit to give time for the treeview to update
            Thread.Sleep(30);
        }

        private void ClickNode(IntPtr Node, MouseButton Button)
        {
            Stopwatch timer;
            bool IsSelected = false;
            int Left = 0;
            int Top = 0;
            int Right = 0;
            int Bottom = 0;
            NM.tagRect ClientRect = new NM.tagRect();

            if (Button == MouseButton.Left)
            {
                // Get state and bounds
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "IsSelected", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                IsSelected = GUI.m_APE.GetValueFromMessage();

                // If we are left clicking and its already selected don't reselect
                if (IsSelected)
                {
                    return;
                }
            }

            timer = Stopwatch.StartNew();
            do
            {
                // Get state and bounds
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "EnsureVisible", MemberTypes.Method);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddMessageGetValue(DataStores.Store7);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                Left = GUI.m_APE.GetValueFromMessage();
                Top = GUI.m_APE.GetValueFromMessage();
                Right = GUI.m_APE.GetValueFromMessage();
                Bottom = GUI.m_APE.GetValueFromMessage();

                NM.GetClientRect(Identity.Handle, out ClientRect);

                if (Left < 0 || Top < 0 || Bottom > ClientRect.bottom)
                {
                    // Go round again
                }
                else
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to select node");
                }
            }
            while (true);

            if (Right > ClientRect.right)
            {
                Right = ClientRect.right;
            }

            if (Button == MouseButton.Left)
            {
                // Click it
                base.MouseSingleClickInternal(Right - 2, Top + ((Bottom - Top) / 2), Button, MouseKeyModifier.None);

                // Wait for it to be selected
                bool IsNowSelected = false;
                timer = Stopwatch.StartNew();
                do
                {
                    // Get if selected
                    GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                    GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "IsSelected", MemberTypes.Property);
                    GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                    GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                    // Get the value(s) returned MUST be done straight after the WaitForMessages call
                    IsNowSelected = GUI.m_APE.GetValueFromMessage();

                    if (IsNowSelected)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw new Exception("Failed to select node");
                    }

                    Thread.Sleep(15);
                }
                while (true);

                // After clicking sleep for a bit to give time for the treeview to update
                Thread.Sleep(30);
            }
            else
            {
                // Click it
                base.MouseSingleClickInternal(Right - 2, Top + ((Bottom - Top) / 2), Button, MouseKeyModifier.None);
            }
        }

        private void CheckNode(IntPtr Node, bool Check)
        {
            int Position = 0;
            Stopwatch timer;
            bool IsChecked = false;
            int Left = 0;
            int Top = 0;
            int Right = 0;
            int Bottom = 0;
            NM.tagRect ClientRect = new NM.tagRect();

            // Check for checkbox style
            if (!m_CheckBoxes)
            {
                throw new Exception("TreeView does not have checkbox style");
            }

            // Check for if its already in the correct state
            // Get state and bounds
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Checked", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            IsChecked = GUI.m_APE.GetValueFromMessage();

            if (IsChecked == Check)
            {
                return;
            }

            // Treeviews have a habbit of scrolling the last selected node back into view so clicking the
            // Node we want to check to select it is a good idea although not in thoery needed
            ClickNode(Node, MouseButton.Left);

            do
            {
                // Get state and bounds
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "EnsureVisible", MemberTypes.Method);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddMessageGetValue(DataStores.Store7);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                Left = GUI.m_APE.GetValueFromMessage();
                Top = GUI.m_APE.GetValueFromMessage();
                Right = GUI.m_APE.GetValueFromMessage();
                Bottom = GUI.m_APE.GetValueFromMessage();

                NM.GetClientRect(Identity.Handle, out ClientRect);

                if (Left < 0 || Top < 0 || Bottom > ClientRect.bottom)
                {
                    // Go round again
                }
                else
                {
                    break;
                }
            }
            while (true);

            if (Right > ClientRect.right)
            {
                Right = ClientRect.right;
            }

            // Scroll the checkbox in to view if need be
            int X = Left - 7 - 5;
            if (X < 0)
            {
                //get current scroll bar pos
                Position = NM.GetScrollPos(Identity.Handle, NM.SBS_HORZ);

                //update the pos
                Position = Position + X;
                NM.SetScrollPos(Identity.Handle, NM.SBS_HORZ, Position, true);

                //send the message to move
                IntPtr MessageResult;
                IntPtr SendResult = NM.SendMessageTimeout(Handle, NM.WM_HSCROLL, new IntPtr(MakeLParam((int)NM.SB_THUMBPOSITION, Position)), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, (uint)GUI.GetTimeOuts(), out MessageResult);

                base.MouseSingleClickInternal(5, Top + ((Bottom - Top) / 2), MouseButton.Left, MouseKeyModifier.None);
            }
            else
            {
                base.MouseSingleClickInternal(Left - 7, Top + ((Bottom - Top) / 2), MouseButton.Left, MouseKeyModifier.None);
            }

            // Wait for it to be checked / unchecked
            bool IsNowChecked = !Check;
            timer = Stopwatch.StartNew();
            do
            {
                // Get the checked state
                GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Checked", MemberTypes.Property);
                GUI.m_APE.AddMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
                GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                IsNowChecked = GUI.m_APE.GetValueFromMessage();

                if (IsNowChecked == Check)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw new Exception("Failed to check / uncheck node");
                }

                Thread.Sleep(15);
            }
            while (true);

            // After clicking sleep for a bit to give time for the treeview to update
            Thread.Sleep(30);
        }

        public void Check(string TreeViewItem)
        {
            GUI.Log("Check treeview item [" + TreeViewItem + "]", LogItemTypeEnum.Action);
            CheckInternal(TreeViewItem, true);
        }

        public void Uncheck(string TreeViewItem)
        {
            GUI.Log("Uncheck treeview item [" + TreeViewItem + "]", LogItemTypeEnum.Action);
            CheckInternal(TreeViewItem, false);
        }

        private void CheckInternal(string TreeViewItem, bool Check)
        {
            string[] Nodes = TreeViewItem.Split(new char[] { '\\' });
            IntPtr NodeHandle = IntPtr.Zero;

            for (int Item = 0; Item < Nodes.Length; Item++)
            {
                if (Item == 0)
                {
                    NodeHandle = GetRootNode(Nodes[Item]);
                }
                else
                {
                    NodeHandle = GetChildNode(NodeHandle, Nodes[Item]);
                }

                if (Item < Nodes.Length - 1)
                {
                    // Expand the node
                    ExpandNode(NodeHandle);
                }
                else
                {
                    // Check the checkbox on the node 
                    CheckNode(NodeHandle, Check);
                }
            }
        }

        public void Select(string TreeViewItem)
        {
            Select(TreeViewItem, MouseButton.Left);
        }

        public void Select(string TreeViewItem, MouseButton Button)
        {
            switch (Button)
            {
                case MouseButton.Left:
                    GUI.Log("Left click treeview item [" + TreeViewItem + "]", LogItemTypeEnum.Action);
                    break;
                case MouseButton.Middle:
                    GUI.Log("Middle click treeview item [" + TreeViewItem + "]", LogItemTypeEnum.Action);
                    break;
                case MouseButton.Right:
                    GUI.Log("Right click treeview item [" + TreeViewItem + "]", LogItemTypeEnum.Action);
                    break;
            }

            string[] Nodes = TreeViewItem.Split(new char[] { '\\' });
            IntPtr NodeHandle = IntPtr.Zero;

            for (int Item = 0; Item < Nodes.Length; Item++)
            {
                if (Item == 0)
                {
                    NodeHandle = GetRootNode(Nodes[Item]);
                }
                else
                {
                    NodeHandle = GetChildNode(NodeHandle, Nodes[Item]);
                }

                if (Item < Nodes.Length - 1)
                {
                    // Expand the node
                    ExpandNode(NodeHandle);
                }
                else
                {
                    // Select the node
                    ClickNode(NodeHandle, Button);
                }
            }
        }
    }
}
