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
using System.Drawing;
using System.Reflection;
using System.Diagnostics;
using APE.Communication;
using System.Threading;
using NM = APE.Native.NativeMethods;
using System.Collections.Generic;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// LZCASHVIEWERLib.AxLzCashViewerCtrl
    /// </summary>
    public sealed class GUIAxLzCashViewerCtrl : GUIFlexgrid
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIAxLzCashViewerCtrl(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Copies the contents of the results grid to the clipboard
        /// </summary>
        public void CopyToClipboard()
        {
            GUI.m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            if (Identity.TypeName == "AxLzCashViewerCtrl")
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "GetOcx", MemberTypes.Method);
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "CopyToClipboard", MemberTypes.Method);
            }
            else
            {
                GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store2, "CopyToClipboard", MemberTypes.Method);
            }
            GUI.m_APE.SendMessages(EventSet.APE);
            GUI.m_APE.WaitForMessages(EventSet.APE);
        }
    }
}