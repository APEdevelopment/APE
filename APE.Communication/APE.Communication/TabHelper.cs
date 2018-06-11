using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace APE.Communication
{
    public partial class APEIPC
    {
        //
        //  GetTabRect
        //

        unsafe public void AddQueryMessageGetTabRect(DataStores sourceStore)
        {
            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;

            ptrMessage->Action = MessageAction.GetTabRect;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetTabRect(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);

            CleanUpMessage(ptrMessage);

            int left = -1;
            int top = -1;
            int width = -1;
            int height = -1;
            object[] parameters = { left, top, width, height };

            ParameterModifier modifiers = new ParameterModifier(4);
            modifiers[0] = true;
            modifiers[1] = true;
            modifiers[2] = true;
            modifiers[3] = true;
            ParameterModifier[] modifiersArray = { modifiers };

            sourceObject.GetType().InvokeMember("GetPositionPix", BindingFlags.InvokeMethod , null, sourceObject, parameters, modifiersArray, null, null);

            AddReturnValue(new Parameter(this, (int)parameters[0]));
            AddReturnValue(new Parameter(this, (int)parameters[1]));
            AddReturnValue(new Parameter(this, (int)parameters[2]));
            AddReturnValue(new Parameter(this, (int)parameters[3]));
        }
    }
}
