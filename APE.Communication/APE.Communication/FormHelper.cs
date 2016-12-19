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

namespace APE.Communication
{
    public partial class APEIPC
    {
        //
        //  GetTitleBarItemRectangle
        //

        unsafe public void AddFirstMessageGetTitleBarItemRectangle(IntPtr Handle, NM.TitleBarStateElement Item)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetTitleBarItemRectangle;

            Parameter HandleParam = new Parameter(this, Handle);
            Parameter ItemParam = new Parameter(this, (int)Item);

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
            IntPtr Handle = GetParameterIntPtr(ptrMessage, 0);

            // p2  = TitleBarStateElement
            NM.TitleBarStateElement Item = (NM.TitleBarStateElement)GetParameterInt32(ptrMessage, 1);

            CleanUpMessage(ptrMessage);

            NM.TITLEBARINFOEX tbi = new NM.TITLEBARINFOEX();
            tbi.cbSize = Marshal.SizeOf(typeof(NM.TITLEBARINFOEX));

            // Send the WM_GETTITLEBARINFOEX message
            IntPtr Result;
            IntPtr Return;
            Return = NM.SendMessageTimeout(Handle, NM.WM_GETTITLEBARINFOEX, IntPtr.Zero, ref tbi, NM.SendMessageTimeoutFlags.SMTO_NORMAL, TimeOut, out Result);

            uint State;
            int Left;
            int Top;
            int Right;
            int Bottom;

            switch (Item)
            {
                case NM.TitleBarStateElement.TitleBar:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.TitleBar];
                    Top = tbi.rcTitleBar.top;
                    Left = tbi.rcTitleBar.left;
                    Bottom = tbi.rcTitleBar.bottom;
                    Right = tbi.rcTitleBar.right;
                    break;
                case NM.TitleBarStateElement.Minimize:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Minimize];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Minimize].right;
                    break;
                case NM.TitleBarStateElement.Maximize:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Maximize];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Maximize].right;
                    break;
                case NM.TitleBarStateElement.Help:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Help];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Help].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Help].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Help].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Help].right;
                    break;
                case NM.TitleBarStateElement.Close:
                    State = (uint)tbi.rgstate[(int)NM.TitleBarStateElement.Close];
                    Top = tbi.rgrect[(int)NM.TitleBarStateElement.Close].top;
                    Left = tbi.rgrect[(int)NM.TitleBarStateElement.Close].left;
                    Bottom = tbi.rgrect[(int)NM.TitleBarStateElement.Close].bottom;
                    Right = tbi.rgrect[(int)NM.TitleBarStateElement.Close].right;
                    break;
                default:
                    throw new Exception("Unknown titlebar element: " + Item.ToString());
            }

            AddReturnValue(new Parameter(this, State));
            AddReturnValue(new Parameter(this, Top));
            AddReturnValue(new Parameter(this, Left));
            AddReturnValue(new Parameter(this, Bottom));
            AddReturnValue(new Parameter(this, Right));
        }
    }
}
