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

        unsafe public void AddQueryMessageGetTabRect(DataStores sourceStore, int tabIndex)
        {
            FirstMessageInitialise();

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;

            Parameter tabIndexParameter = new Parameter(this, tabIndex);

            ptrMessage->Action = MessageAction.GetTabRect;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneFind = true;
            m_DoneQuery = true;
            m_DoneGet = true;
        }

        private unsafe void GetTabRect(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            int tabIndex = GetParameterInt32(ptrMessage, 0);

            CleanUpMessage(ptrMessage);

            int left = -1;
            int top = -1;
            int width = -1;
            int height = -1;
            object[] parameters = { left, top, width, height };

            ComReflectInternal("GetPositionPix", sourceObject, parameters);

            AddReturnValue(new Parameter(this, left));
            AddReturnValue(new Parameter(this, top));
            AddReturnValue(new Parameter(this, width));
            AddReturnValue(new Parameter(this, height));
        }
    }
}
