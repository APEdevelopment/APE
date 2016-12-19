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
using System.Runtime.InteropServices;
using NM = APE.Native.NativeMethods;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    public partial class APEIPC
    {
        private Type m_ScrollableControl = typeof(WF.ScrollableControl);

        //
        //  GetTitleBarItemRectangle
        //

        unsafe public void AddFirstMessageGetTitleBarItemRectangle(IntPtr handle, NM.TitleBarStateElement item)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetTitleBarItemRectangle;

            Parameter handleParam = new Parameter(this, handle);
            Parameter itemParam = new Parameter(this, (int)item);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetTitleBarItemRectangle(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GetTitleBarItemRectangle must be first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);

            // p2  = TitleBarStateElement
            NM.TitleBarStateElement item = (NM.TitleBarStateElement)GetParameterInt32(ptrMessage, 1);

            CleanUpMessage(ptrMessage);

            NM.TITLEBARINFOEX tbi = new NM.TITLEBARINFOEX();
            tbi.cbSize = Marshal.SizeOf(typeof(NM.TITLEBARINFOEX));

            // Send the WM_GETTITLEBARINFOEX message
            IntPtr result;
            IntPtr smtReturn;
            smtReturn = NM.SendMessageTimeout(handle, NM.WM_GETTITLEBARINFOEX, IntPtr.Zero, ref tbi, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out result);

            uint state;
            int left;
            int top;
            int right;
            int bottom;

            switch (item)
            {
                case NM.TitleBarStateElement.TitleBar:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.TitleBar];
                    top = tbi.rcTitleBar.top;
                    left = tbi.rcTitleBar.left;
                    bottom = tbi.rcTitleBar.bottom;
                    right = tbi.rcTitleBar.right;
                    break;
                case NM.TitleBarStateElement.Minimize:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Minimize];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].right;
                    break;
                case NM.TitleBarStateElement.Maximize:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Maximize];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].right;
                    break;
                case NM.TitleBarStateElement.Help:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Help];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Help].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Help].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Help].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Help].right;
                    break;
                case NM.TitleBarStateElement.Close:
                    state = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Close];
                    top = tbi.rgrect[(int)NM.TitleBarStateElement.Close].top;
                    left = tbi.rgrect[(int)NM.TitleBarStateElement.Close].left;
                    bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Close].bottom;
                    right = tbi.rgrect[(int)NM.TitleBarStateElement.Close].right;
                    break;
                default:
                    throw new Exception("Unknown titlebar element: " + item.ToString());
            }

            AddReturnValue(new Parameter(this, state));
            AddReturnValue(new Parameter(this, top));
            AddReturnValue(new Parameter(this, left));
            AddReturnValue(new Parameter(this, bottom));
            AddReturnValue(new Parameter(this, right));
        }

        //
        //  ScrollControlIntoView
        //

        unsafe public void AddFirstMessageScrollControlIntoView(IntPtr handle)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.ScrollControlIntoView;

            Parameter HandleParam = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void ScrollControlIntoView(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("ScrollControlIntoView must be first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);

            WF.Control childControl = WF.Control.FromHandle(handle);
            if (childControl != null)
            {
                WF.Control control = childControl.Parent;
                while (control != null)
                {
                    if (control.GetType().IsSubclassOf(m_ScrollableControl))
                    {
                        ((WF.ScrollableControl)control).ScrollControlIntoView(childControl);
                    }

                    control = control.Parent;
                }
            }
            //TODO WPF etc
        }
    }
}
