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
        private int m_ImageWidth = 16;
        private int m_ImageHeight = 16;
        private int m_StateImageWidth = 16;
        private int m_StateImageHeight = 16;

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
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetNodeCount", MemberTypes.Method, new Parameter(GUI.m_APE, false));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "ShowPlusMinus", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store3, "CheckBoxes", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store4, "ImageList", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ImageSize", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store7, "Height", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store4, "StateImageList", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "ImageSize", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store8, "Width", MemberTypes.Property);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store9, "Height", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int ChildNodes = GUI.m_APE.GetValueFromMessage();
            m_ShowPlusMinus = GUI.m_APE.GetValueFromMessage();
            m_CheckBoxes = GUI.m_APE.GetValueFromMessage();
            dynamic imageWidth = GUI.m_APE.GetValueFromMessage();
            dynamic imageHeight = GUI.m_APE.GetValueFromMessage();
            dynamic stateImageWidth = GUI.m_APE.GetValueFromMessage();
            dynamic stateImageHeight = GUI.m_APE.GetValueFromMessage();
            if (imageWidth != null)
            {
                m_ImageWidth = imageWidth;
                m_ImageHeight = imageHeight;
            }
            if (stateImageWidth != null)
            {
                m_StateImageWidth = stateImageWidth;
                m_StateImageHeight = stateImageHeight;
            }

            for (int ChildItem = 0; ChildItem < ChildNodes; ChildItem++)
            {
                // Get the text of the root node
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Nodes", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, ChildItem));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string Text = GUI.m_APE.GetValueFromMessage();
                IntPtr NewNodeHandle = GUI.m_APE.GetValueFromMessage();

                if (Text == NodeText)
                {
                    return NewNodeHandle;
                }
            }

            throw GUI.ApeException("Failed to find root node");
        }

        private IntPtr GetChildNode(IntPtr Node, string NodeText)
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "GetNodeCount", MemberTypes.Method, new Parameter(GUI.m_APE, false));
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int ChildNodes = GUI.m_APE.GetValueFromMessage();

            for (int ChildItem = 0; ChildItem < ChildNodes; ChildItem++)
            {
                // Get the text of the child node
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Nodes", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "<Indexer>", MemberTypes.Property, new Parameter(GUI.m_APE, ChildItem));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Text", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Handle", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string Text = GUI.m_APE.GetValueFromMessage();
                IntPtr NewNodeHandle = GUI.m_APE.GetValueFromMessage();

                if (Text == NodeText)
                {
                    return NewNodeHandle;
                }
            }

            throw GUI.ApeException("Failed to find child node");
        }

        private int MakeLParam(int LoWord, int HiWord)
        {
            return ((HiWord << 16) | (LoWord & 0xffff));
        }

        private void ExpandNode(IntPtr Node)
        {
            int position = 0;
            Stopwatch timer;
            int left = 0;
            int top = 0;
            int right = 0;
            int bottom = 0;
            int imageIndex = -1;
            string imageKey = "";
            NM.tagRect clientRect = new NM.tagRect();

            // check if already expanded
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "IsExpanded", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            bool isExpanded = GUI.m_APE.GetValueFromMessage();

            if (isExpanded)
            {
                return;
            }

            for (int i = 0; i < 2; i++)
            {
                // Get bounds
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "EnsureVisible", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store8, "ImageIndex", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store9, "ImageKey", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                left = GUI.m_APE.GetValueFromMessage();
                top = GUI.m_APE.GetValueFromMessage();
                right = GUI.m_APE.GetValueFromMessage();
                bottom = GUI.m_APE.GetValueFromMessage();
                imageIndex = GUI.m_APE.GetValueFromMessage();
                imageKey = GUI.m_APE.GetValueFromMessage();

                NM.GetClientRect(Identity.Handle, out clientRect);

                if (left < 0 || top < 0 || bottom > clientRect.bottom)
                {
                    // Go round again
                }
                else
                {
                    break;
                }
            }

            if (right > clientRect.right)
            {
                right = clientRect.right;
            }

            bool waitForMove = false;
            if (bottom > clientRect.bottom)
            {
                waitForMove = true;
                bottom = clientRect.bottom;
            }

            // Expand it
            if (m_ShowPlusMinus)
            {
                int stateImageWidth = 0;
                if (m_CheckBoxes)
                {
                    stateImageWidth = m_StateImageWidth;
                }

                int imageWidth = 0;
                if (imageIndex != -1 || imageKey != "")
                {
                    imageWidth = m_ImageWidth;
                }

                int expandImageWidth = m_ImageWidth;

                // Scroll the expand icon in to view if need be
                int x = left - imageWidth - stateImageWidth - expandImageWidth;
                if (x != 0)
                {
                    //get current scroll bar pos
                    position = NM.GetScrollPos(Identity.Handle, NM.SBS_HORZ);

                    //update the pos
                    position = position + x;
                    NM.SetScrollPos(Identity.Handle, NM.SBS_HORZ, position, false);

                    //send the message to move
                    IntPtr MessageResult;
                    IntPtr SendResult = NM.SendMessageTimeout(Handle, NM.WM_HSCROLL, new IntPtr(MakeLParam((int)NM.SB_THUMBPOSITION, position)), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, (uint)GUI.GetTimeOut(), out MessageResult);

                    //update the bounds
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    // Get the value(s) returned MUST be done straight after the WaitForMessages call
                    left = GUI.m_APE.GetValueFromMessage();
                    top = GUI.m_APE.GetValueFromMessage();
                    right = GUI.m_APE.GetValueFromMessage();
                    bottom = GUI.m_APE.GetValueFromMessage();
                }
                
                base.SingleClickInternal(left - imageWidth - stateImageWidth - (expandImageWidth / 2), top + ((bottom - top) / 2), MouseButton.Left, MouseKeyModifier.None);
            }
            else
            {
                base.DoubleClickInternal(right - 4, top + ((bottom - top) / 2), MouseButton.Left, MouseKeyModifier.None);
            }

            // Wait for it to be expanded
            bool isNowExpanded = false;
            timer = Stopwatch.StartNew();
            do
            {
                // Get if expanded
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "IsExpanded", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                isNowExpanded = GUI.m_APE.GetValueFromMessage();

                if (isNowExpanded)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException("Node failed to expand");
                }

                Thread.Sleep(15);
            }
            while (true);

            //left clicking on a node which is partially off the bottom of the window will scroll it
            if (waitForMove)
            {
                timer = Stopwatch.StartNew();
                do
                {
                    // Get bounds
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Bounds", MemberTypes.Property);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Bottom", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    // Get the value(s) returned MUST be done straight after the WaitForMessages call
                    bottom = GUI.m_APE.GetValueFromMessage();

                    if (bottom >= clientRect.bottom)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw GUI.ApeException("Node failed to scroll fully into view");
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
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "IsSelected", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
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
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "EnsureVisible", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                Left = GUI.m_APE.GetValueFromMessage();
                Top = GUI.m_APE.GetValueFromMessage();
                Right = GUI.m_APE.GetValueFromMessage();
                Bottom = GUI.m_APE.GetValueFromMessage();

                NM.GetClientRect(Identity.Handle, out ClientRect);

                if (Top == 0)
                {
                    if (Bottom > ClientRect.bottom)
                    {
                        Bottom = ClientRect.bottom;
                    }
                }

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
                    throw GUI.ApeException("Failed to select node");
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
                base.SingleClickInternal(Right - 2, Top + ((Bottom - Top) / 2), Button, MouseKeyModifier.None);

                // Wait for it to be selected
                bool IsNowSelected = false;
                timer = Stopwatch.StartNew();
                do
                {
                    // Get if selected
                    GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "IsSelected", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    // Get the value(s) returned MUST be done straight after the WaitForMessages call
                    IsNowSelected = GUI.m_APE.GetValueFromMessage();

                    if (IsNowSelected)
                    {
                        break;
                    }

                    if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                    {
                        throw GUI.ApeException("Failed to select node");
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
                base.SingleClickInternal(Right - 2, Top + ((Bottom - Top) / 2), Button, MouseKeyModifier.None);
            }
        }

        private void CheckNode(IntPtr Node, bool Check)
        {
            int position = 0;
            Stopwatch timer;
            bool isChecked = false;
            int left = 0;
            int top = 0;
            int right = 0;
            int bottom = 0;
            int imageIndex = -1;
            string imageKey = "";
            NM.tagRect clientRect = new NM.tagRect();

            // Check for checkbox style
            if (!m_CheckBoxes)
            {
                throw GUI.ApeException("TreeView does not have checkbox style");
            }

            // Check for if its already in the correct state
            // Get state and bounds
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
            GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Checked", MemberTypes.Property);
            GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            isChecked = GUI.m_APE.GetValueFromMessage();

            if (isChecked == Check)
            {
                return;
            }

            // Treeviews have a habbit of scrolling the last selected node back into view so clicking the
            // Node we want to check to select it is a good idea although not in thoery needed
            ClickNode(Node, MouseButton.Left);

            do
            {
                // Get state and bounds
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "EnsureVisible", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store8, "ImageIndex", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store9, "ImageKey", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store9);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                left = GUI.m_APE.GetValueFromMessage();
                top = GUI.m_APE.GetValueFromMessage();
                right = GUI.m_APE.GetValueFromMessage();
                bottom = GUI.m_APE.GetValueFromMessage();
                imageIndex = GUI.m_APE.GetValueFromMessage();
                imageKey = GUI.m_APE.GetValueFromMessage();

                NM.GetClientRect(Identity.Handle, out clientRect);

                if (top == 0)
                {
                    if (bottom > clientRect.bottom)
                    {
                        bottom = clientRect.bottom;
                    }
                }
                
                if (left < 0 || top < 0 || bottom > clientRect.bottom)
                {
                    // Go round again
                }
                else
                {
                    break;
                }
            }
            while (true);

            if (right > clientRect.right)
            {
                right = clientRect.right;
            }

            int imageWidth = 0;
            if (imageIndex != -1 || imageKey != "")
            {
                imageWidth = m_ImageWidth;
            }

            int stateImageWidth = m_StateImageWidth;

            // Scroll the checkbox in to view if need be
            int x = left - imageWidth - stateImageWidth;
            if (x != 0)
            {
                //get current scroll bar pos
                position = NM.GetScrollPos(Identity.Handle, NM.SBS_HORZ);

                //update the pos
                position = position + x;
                NM.SetScrollPos(Identity.Handle, NM.SBS_HORZ, position, true);

                //send the message to move
                IntPtr MessageResult;
                IntPtr SendResult = NM.SendMessageTimeout(Handle, NM.WM_HSCROLL, new IntPtr(MakeLParam((int)NM.SB_THUMBPOSITION, position)), IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, (uint)GUI.GetTimeOut(), out MessageResult);

                //update the bounds
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store3, "Bounds", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store4, "Left", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store5, "Top", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store6, "Right", MemberTypes.Property);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store3, DataStores.Store7, "Bottom", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store7);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                left = GUI.m_APE.GetValueFromMessage();
                top = GUI.m_APE.GetValueFromMessage();
                right = GUI.m_APE.GetValueFromMessage();
                bottom = GUI.m_APE.GetValueFromMessage();
            }

            base.SingleClickInternal(left - imageWidth - (stateImageWidth / 2), top + ((bottom - top) / 2), MouseButton.Left, MouseKeyModifier.None);

            // Wait for it to be checked / unchecked
            bool isNowChecked = !Check;
            timer = Stopwatch.StartNew();
            do
            {
                // Get the checked state
                GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "NodeFromHandle", MemberTypes.Method, new Parameter(GUI.m_APE, Node));
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Checked", MemberTypes.Property);
                GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                GUI.m_APE.SendMessages(EventSet.APE);
                GUI.m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                isNowChecked = GUI.m_APE.GetValueFromMessage();

                if (isNowChecked == Check)
                {
                    break;
                }

                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException("Failed to check / uncheck node");
                }

                Thread.Sleep(15);
            }
            while (true);

            // After clicking sleep for a bit to give time for the treeview to update
            Thread.Sleep(30);
        }

        /// <summary>
        /// Checks the specified item in the tree view 
        /// </summary>
        /// <param name="treeViewItem">The item to check</param>
        public void CheckItem(string treeViewItem)
        {
            GUI.Log("Check treeview item [" + treeViewItem + "]", LogItemType.Action);
            CheckInternal(treeViewItem, true);
        }

        /// <summary>
        /// Unchecks the specified item in the tree view 
        /// </summary>
        /// <param name="treeViewItem">The item to uncheck</param>
        public void UncheckItem(string treeViewItem)
        {
            GUI.Log("Uncheck treeview item [" + treeViewItem + "]", LogItemType.Action);
            CheckInternal(treeViewItem, false);
        }

        private void CheckInternal(string TreeViewItem, bool Check)
        {
            string[] Nodes = TreeViewItem.Split(GUI.TreeViewDelimiterAsArray, StringSplitOptions.None);
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

        /// <summary>
        /// Selects the specified item in the tree view
        /// </summary>
        /// <param name="treeViewItem">The item to select</param>
        public void SingleClickItem(string treeViewItem)
        {
            SingleClickItem(treeViewItem, MouseButton.Left);
        }

        /// <summary>
        /// Clicks on the specified item in the tree view with the specified mouse button
        /// </summary>
        /// <param name="treeViewItem">The item to click on</param>
        /// <param name="button">The mouse button to click with</param>
        public void SingleClickItem(string treeViewItem, MouseButton button)
        {
            switch (button)
            {
                case MouseButton.Left:
                    GUI.Log("Left click treeview item [" + treeViewItem + "]", LogItemType.Action);
                    break;
                case MouseButton.Middle:
                    GUI.Log("Middle click treeview item [" + treeViewItem + "]", LogItemType.Action);
                    break;
                case MouseButton.Right:
                    GUI.Log("Right click treeview item [" + treeViewItem + "]", LogItemType.Action);
                    break;
            }

            string[] Nodes = treeViewItem.Split(GUI.TreeViewDelimiterAsArray, StringSplitOptions.None);
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
                    ClickNode(NodeHandle, button);
                }
            }
        }
    }
}
