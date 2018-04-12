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
using System.ComponentModel;
using System.Diagnostics;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security;
using System.Threading;
using NM = APE.Native.NativeMethods;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    public partial class APEIPC
    {
        //
        //  AddGenericWalkerSelectedHandler
        //

        private bool m_GenericWalkerSelected = false;
        private Delegate m_GenericWalkerSelectedHandler = null;
        private EventInfo m_GenericWalkerSelectedEventInfo = null;
        private dynamic m_GenericWalkerControl = null;
        private IntPtr m_GenericWalkerParent;

        /// <summary>
        /// Calls into the AUT to add a selected handler to the specific control
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the control object</param>
        unsafe public void AddQueryMessageAddGenericWalkerSelectedHandler(DataStores sourceStore, IntPtr parentFormHandle)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the generic walker before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->Action = MessageAction.AddGenericWalkerSelectedHandler;

            Parameter parentFormHandleParameter = new Parameter(this, parentFormHandle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Adds a selected handler to the control
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void AddGenericWalkerSelectedHandler(Message* ptrMessage)
        {
            m_GenericWalkerControl = GetObjectFromDatastore(ptrMessage->SourceStore);
            m_GenericWalkerParent = GetParameterIntPtr(ptrMessage, 0);

            m_GenericWalkerSelected = false;

            if (m_GenericWalkerControl != null)
            {
                if (m_GenericWalkerSelectedHandler == null)
                {
                    MethodInfo genericWalkerSelectedMethodInfo = this.GetType().GetMethod("GenericWalkerSelectedHandler", BindingFlags.NonPublic | BindingFlags.Instance);
                    if (genericWalkerSelectedMethodInfo == null)
                    {
                        throw new Exception("Failed to find method GenericWalkerSelectedHandler");
                    }
                    m_GenericWalkerSelectedEventInfo = m_GenericWalkerControl.GetType().GetEvent("Selected");
                    if (m_GenericWalkerSelectedEventInfo == null)
                    {
                        throw new Exception("Failed to find event Selected");
                    }
                    m_GenericWalkerSelectedHandler = Delegate.CreateDelegate(m_GenericWalkerSelectedEventInfo.EventHandlerType, this, genericWalkerSelectedMethodInfo);
                }

                m_GenericWalkerSelectedEventInfo.AddEventHandler(m_GenericWalkerControl, m_GenericWalkerSelectedHandler);
            }

            CleanUpMessage(ptrMessage);
        }

        private void GenericWalkerSelectedHandler(object sender, EventArgs e)
        {
            m_GenericWalkerSelected = true;
        }

        //
        //  WaitForAndRemoveGenericWalkerSelectedHandler
        //

        /// <summary>
        /// Calls into the AUT to wait for then remove the selected handler
        /// </summary>
        unsafe public void AddFirstMessageWaitForAndRemoveGenericWalkerSelectedHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.WaitForAndRemoveGenericWalkerSelectedHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Waits for the handler to set the m_GenericWalkerControl variable to true then removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void WaitForAndRemoveGenericWalkerSelectedHandler(Message* ptrMessage)
        {
            if (m_GenericWalkerControl != null)
            {
                try
                {
                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (m_GenericWalkerSelected || m_GenericWalkerControl.Disposing || m_GenericWalkerControl.IsDisposed)
                        {
                            break;
                        }

                        if (!NM.IsWindowEnabled(m_GenericWalkerParent))
                        {
                            //parent not enabled so probably a modal popup displayed so break out the loop
                            break;
                        }

                        if (timer.ElapsedMilliseconds > m_TimeOut)
                        {
                            throw new Exception("Failed to find generic walker selected event");
                        }

                        Thread.Sleep(15);
                    }
                }
                finally
                {
                    m_GenericWalkerSelectedEventInfo.RemoveEventHandler(m_GenericWalkerControl, m_GenericWalkerSelectedHandler);
                    m_GenericWalkerControl = null;
                    m_GenericWalkerParent = IntPtr.Zero;
                }
            }

            CleanUpMessage(ptrMessage);
        }
    }
}
