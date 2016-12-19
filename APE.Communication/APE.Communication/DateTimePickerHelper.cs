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
using NM = APE.Native.NativeMethods;

namespace APE.Communication
{
    public partial class APEIPC
    {

        //
        //  GetDateTimePickerCheckboxRectangle
        //

        unsafe public void AddFirstMessageGetDateTimePickerCheckboxRectangle(IntPtr handle)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetDateTimePickerCheckboxRectangle;

            Parameter handleParam = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetDateTimePickerCheckboxRectangle(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GetDateTimePickerCheckboxRectangle must be first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);

            CleanUpMessage(ptrMessage);

            NM.DATETIMEPICKERINFO dtpi = NM.GetDateTimePickerInfo(handle);
            AddReturnValue(new Parameter(this, dtpi.rcCheck.top));
            AddReturnValue(new Parameter(this, dtpi.rcCheck.left));
            AddReturnValue(new Parameter(this, dtpi.rcCheck.bottom));
            AddReturnValue(new Parameter(this, dtpi.rcCheck.right));
        }

        //
        //  GetDateTimePickerButtonRectangle
        //

        unsafe public void AddFirstMessageGetDateTimePickerButtonRectangle(IntPtr handle)
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetDateTimePickerButtonRectangle;

            Parameter handleParam = new Parameter(this, handle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetDateTimePickerButtonRectangle(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GetDateTimePickerButtonRectangle must be first message");
            }

            // p1  = handle
            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);

            CleanUpMessage(ptrMessage);

            NM.DATETIMEPICKERINFO dtpi = NM.GetDateTimePickerInfo(handle);
            AddReturnValue(new Parameter(this, dtpi.rcButton.top));
            AddReturnValue(new Parameter(this, dtpi.rcButton.left));
            AddReturnValue(new Parameter(this, dtpi.rcButton.bottom));
            AddReturnValue(new Parameter(this, dtpi.rcButton.right));
        }
    }
}
