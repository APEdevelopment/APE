using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NM = APE.Native.NativeMethods;

namespace APE.Communication
{
    public partial class APEIPC
    {
        unsafe public void AddFirstMessageGetComboBoxExItemText(IntPtr handle, int index, string description)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();

            ptrMessage->Action = MessageAction.GetComboBoxExItemText;

            Parameter handleParam = new Parameter(this, handle);
            Parameter indexParam = new Parameter(this, index);
            Parameter descriptionParam = new Parameter(this, description);

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetComboBoxExItemText(Message* ptrMessage, int messageNumber)
        {
            //must be first message
            if (messageNumber != 1)
            {
                throw new Exception("GetComboBoxExItemText must be first message");
            }

            IntPtr handle = GetParameterIntPtr(ptrMessage, 0);
            int index = GetParameterInt32(ptrMessage, 1);
            string description = GetParameterString(ptrMessage, 2);
            
            CleanUpMessage(ptrMessage);

            IntPtr sendResult;
            IntPtr messageResult;

            //Work out if the control is unicode or not
            bool ASCII = false;
            sendResult = NM.SendMessageTimeout(handle, NM.CBEM_GETUNICODEFORMAT, IntPtr.Zero, IntPtr.Zero, NM.SendMessageTimeoutFlags.SMTO_NORMAL, this.TimeOut, out messageResult);
            if (sendResult == IntPtr.Zero) //Failed
            {
                throw new Exception("Failed to get the character format of the " + description);
            }
            if (messageResult == IntPtr.Zero)
            {
                ASCII = true;
            }

            //Get the item text
            NM.ComboBoxExItem item = new NM.ComboBoxExItem();
            item.iItem = new IntPtr(index);
            item.mask = NM.CBEIF_TEXT;
            item.pszText = new string(' ', 1001);
            item.cchTextMax = 1000;

            if (ASCII)
            {
                sendResult = NM.SendMessageTimeout(handle, NM.CBEM_GETITEMA, IntPtr.Zero, ref item, NM.SendMessageTimeoutFlags.SMTO_NORMAL, this.TimeOut, out messageResult);
                if (sendResult == IntPtr.Zero || messageResult == IntPtr.Zero) //Failed
                {
                    throw new Exception("Failed to get the text of the item in the " + description);
                }

                byte[] bytes = Encoding.Unicode.GetBytes(item.pszText);
                string ascii = Encoding.ASCII.GetString(bytes);
                AddReturnValue(new Parameter(this, ascii.Substring(0, ascii.IndexOf('\0'))));
            }
            else
            {
                sendResult = NM.SendMessageTimeout(handle, NM.CBEM_GETITEMW, IntPtr.Zero, ref item, NM.SendMessageTimeoutFlags.SMTO_NORMAL, this.TimeOut, out messageResult);
                if (sendResult == IntPtr.Zero || messageResult == IntPtr.Zero) //Failed
                {
                    throw new Exception("Failed to get the text of the item in the " + description);
                }

                AddReturnValue(new Parameter(this, item.pszText.Substring(0, item.pszText.IndexOf('\0'))));
            }
        }
    }
}
