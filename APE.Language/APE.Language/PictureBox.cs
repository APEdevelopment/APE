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
using System.ComponentModel;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Runtime.InteropServices;
using APE.Capture;
using APE.Communication;
using System.Threading;
using System.Drawing.Imaging;
using NM = APE.Native.NativeMethods;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate controls derived from the following:
    /// System.Windows.Forms.PictureBox
    /// </summary>
    public sealed class GUIPictureBox : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUIPictureBox(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Saves the image of the picturebox to the specified file
        /// </summary>
        /// <param name="filename">Filename including path to save the image to</param>
        public void Save(string filename)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Image", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Save", MemberTypes.Method, new Parameter(GUI.m_APE, filename));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store3, "Height", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic Height = GUI.m_APE.GetValueFromMessage();

            if (Height == null)
            {
                throw new Exception("PictureBox does not have an image set");
            }
        }

        /// <summary>
        /// Saves the background image of the picturebox to the specified file
        /// </summary>
        /// <param name="filename">Filename including path to save the background image to</param>
        public void SaveBackground(string filename)
        {
            GUI.m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "BackgroundImage", MemberTypes.Property);
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Save", MemberTypes.Method, new Parameter(GUI.m_APE, filename));
            GUI.m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store3, "Height", MemberTypes.Property);
            GUI.m_APE.AddMessageGetValue(DataStores.Store3);
            GUI.m_APE.SendMessages(APEIPC.EventSet.APE);
            GUI.m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            dynamic Height = GUI.m_APE.GetValueFromMessage();

            if (Height == null)
            {
                throw new Exception("PictureBox does not have an background image set");
            }
        }

        //TODO get the image back via the mmf, should be better performance than using the disk
    }
}
