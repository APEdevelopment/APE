using System;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using NM = APE.Native.NativeMethods;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    public partial class APEIPC
    {
        private bool m_ToolStripItemEntered = false;
        private Delegate m_ToolStripItemEnteredHandler = null;
        private EventInfo m_ToolStripItemEnteredEventInfo = null;
        private dynamic m_ToolStripItemControl = null;
        private IntPtr m_ToolStripControlHandle;

        /// <summary>
        /// Calls into the AUT to add a mouse entered handler to the specific control
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the control object</param>
        unsafe public void AddQueryMessageToolStripItemEnteredHandler(DataStores sourceStore, IntPtr toolStripHandle)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the toolstrip before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->Action = MessageAction.AddToolStripItemEnteredHandler;

            Parameter toolStripHandleParameter = new Parameter(this, toolStripHandle);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Adds a selected handler to the control
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void AddToolStripItemEnteredHandler(Message* ptrMessage)
        {
            m_ToolStripItemControl = GetObjectFromDatastore(ptrMessage->SourceStore);
            m_ToolStripControlHandle = GetParameterIntPtr(ptrMessage, 0);

            m_ToolStripItemEntered = false;

            if (m_ToolStripItemControl != null)
            {
                if (m_ToolStripItemEnteredHandler == null)
                {
                    MethodInfo toolStripItemEnteredMethodInfo = this.GetType().GetMethod("ToolStripItemEnteredHandler", BindingFlags.NonPublic | BindingFlags.Instance);
                    if (toolStripItemEnteredMethodInfo == null)
                    {
                        throw new Exception("Failed to find method ToolStripItemEnteredHandler");
                    }
                    m_ToolStripItemEnteredEventInfo = m_ToolStripItemControl.GetType().GetEvent("MouseEnter");
                    if (m_ToolStripItemEnteredEventInfo == null)
                    {
                        throw new Exception("Failed to find event MouseEnter");
                    }
                    m_ToolStripItemEnteredHandler = Delegate.CreateDelegate(m_ToolStripItemEnteredEventInfo.EventHandlerType, this, toolStripItemEnteredMethodInfo);
                }

                m_ToolStripItemEnteredEventInfo.AddEventHandler(m_ToolStripItemControl, m_ToolStripItemEnteredHandler);
            }

            CleanUpMessage(ptrMessage);
        }

        private void ToolStripItemEnteredHandler(object sender, EventArgs e)
        {
            m_ToolStripItemEntered = true;
        }

        //
        //  WaitForAndRemoveToolStripItemEnteredHandler
        //

        /// <summary>
        /// Calls into the AUT to wait for then remove the selected handler
        /// </summary>
        unsafe public void AddFirstMessageWaitForAndRemoveToolStripItemEnteredHandler()
        {
            // Window messages 0x0400 (WM_USER) or higher are not marshalled by windows so make the call in the AUT
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.WaitForAndRemoveToolStripItemEnteredHandler;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        /// <summary>
        /// Waits for the handler to set the m_ToolStripItemEntered variable to true then removes the handler
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void WaitForAndRemoveToolStripItemEnteredHandler(Message* ptrMessage)
        {
            if (m_ToolStripItemControl != null)
            {
                try
                {
                    Stopwatch timer = Stopwatch.StartNew();
                    while (true)
                    {
                        if (m_ToolStripItemEntered || m_ToolStripItemControl.IsDisposed)
                        {
                            break;
                        }

                        if (!NM.IsWindowEnabled(m_ToolStripControlHandle))
                        {
                            //toolstrip not enabled so probably a modal popup displayed so break out the loop
                            break;
                        }

                        if (timer.ElapsedMilliseconds > m_TimeOut)
                        {
                            throw new Exception("Failed to find toolstrip item entered event");
                        }

                        Thread.Sleep(15);
                    }
                }
                finally
                {
                    m_ToolStripItemEnteredEventInfo.RemoveEventHandler(m_ToolStripItemControl, m_ToolStripItemEnteredHandler);
                    m_ToolStripItemControl = null;
                    m_ToolStripControlHandle = IntPtr.Zero;
                }
            }

            CleanUpMessage(ptrMessage);
        }
    }
}
