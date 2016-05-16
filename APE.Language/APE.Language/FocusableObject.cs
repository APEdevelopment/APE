using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using APE.Communication;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    [EditorBrowsable(EditorBrowsableState.Never)]
    public class GUIFocusableObject : GUIObject
    {
        public GUIFocusableObject(string descriptionOfControl, params Identifier[] identParams)
            : base(descriptionOfControl, identParams)
        {
        }

        public GUIFocusableObject(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        public bool HasFocus
        {
            get
            {
                return Input.HasFocus(Identity.ParentHandle, Identity.Handle);
            }
        }

        public void SetFocus()
        {
            Input.SetFocus(Identity.ParentHandle, Identity.Handle);
        }

        protected void SendKeys(string TextToSend)
        {
            GUI.Log("Type [" + TextToSend + "] into " + m_DescriptionOfControl, LogItemTypeEnum.Action);
            SendKeysInternal(TextToSend);
        }

        protected void SendKeysInternal(string TextToSend)
        {
            Input.Block(Identity.ParentHandle, Identity.Handle);

            try
            {
                if (!HasFocus)
                {
                    SetFocus();
                }

                Input.SendKeys(Identity.Handle, TextToSend);
            }
            finally
            {
                Input.Unblock();
            }
        }
    }
}
