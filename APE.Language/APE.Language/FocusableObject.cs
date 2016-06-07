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
            GUI.Log("Type [" + TextToSend + "] into the " + m_DescriptionOfControl, LogItemTypeEnum.Action);
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
