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
    /// System.Windows.Forms.Label
    /// </summary>
    public sealed class GUILabel : GUIObject
    {
        /// <summary>
        /// Constructor used for non-form controls
        /// </summary>
        /// <param name="parentForm">The top level form the control belongs to</param>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="identParams">One or more identifier object(s) used to locate the control.
        /// <para/>Normally you would just use the name identifier</param>
        public GUILabel(GUIForm parentForm, string descriptionOfControl, params Identifier[] identParams)
            : base(parentForm, descriptionOfControl, identParams)
        {
        }

        /// <summary>
        /// Gets the windows classname (not that helpful in the .NET world)
        /// </summary>
        public override string ClassName
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    return null;
                }
                else
                {
                    return base.ClassName;
                }
            }
        }

        /// <summary>
        /// Moves the mouse cursor to the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to move the mouse</param>
        /// <param name="Y">How far from the top edge of the control to move the mouse</param>
        public override void MoveTo(int X, int Y)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.MoveTo(X, Y);
            }
        }

        /// <summary>
        /// Perform a left mouse single click in the middle of the control
        /// </summary>
        public override void SingleClick()
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.SingleClick();
            } 
        }

        /// <summary>
        /// Perform a mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to click</param>
        public override void SingleClick(MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.SingleClick(button);
            }
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        public override void SingleClick(int X, int Y, MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.SingleClick(X, Y, button);
            }
        }

        /// <summary>
        /// Perform a mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to click the mouse</param>
        /// <param name="button">The button to click</param>
        /// <param name="keys">The key to hold while clicking</param>
        public override void SingleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.SingleClick(X, Y, button, keys);
            }
        }

        /// <summary>
        /// Perform a left mouse double click in the middle of the control
        /// </summary>
        public override void DoubleClick()
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.DoubleClick();
            }
        }

        /// <summary>
        /// Perform a double mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to double click</param>
        public override void DoubleClick(MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.DoubleClick(button);
            }
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        public override void DoubleClick(int X, int Y, MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.DoubleClick(X, Y, button);
            }
        }

        /// <summary>
        /// Perform a double mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to double click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to double click the mouse</param>
        /// <param name="button">The button to double click</param>
        /// <param name="keys">The key to hold while double clicking</param>
        public override void DoubleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.DoubleClick(X, Y, button, keys);
            }
        }

        /// <summary>
        /// Perform a left mouse triple click in the middle of the control
        /// </summary>
        public override void TripleClick()
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.TripleClick();
            }
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button in the middle of the control
        /// </summary>
        /// <param name="button">The button to triple click</param>
        public override void TripleClick(MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.TripleClick(button);
            }
        }

        /// <summary>
        ///  Perform a triple mouse click with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="button">The button to triple click</param>
        public override void TripleClick(int X, int Y, MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.TripleClick(X, Y, button);
            }
        }

        /// <summary>
        /// Perform a triple mouse click with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to triple click the mouse</param>
        /// <param name="Y">How far from the top edge of the control to triple click the mouse</param>
        /// <param name="button">The button to triple click</param>
        /// <param name="keys">The key to hold while triple clicking</param>
        public override void TripleClick(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.TripleClick(X, Y, button, keys);
            }
        }

        /// <summary>
        /// Perform a mouse down with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse down</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse down</param>
        /// <param name="button">The button to press</param>
        public override void MouseDown(int X, int Y, MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.MouseDown(X, Y, button);
            }
        }

        /// <summary>
        /// Perform a mouse down with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse down</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse down</param>
        /// <param name="button">The button to press</param>
        /// <param name="keys">The key to hold while performing a mouse down</param>
        public override void MouseDown(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.MouseDown(X, Y, button, keys);
            }
        }

        /// <summary>
        /// Perform a mouse up with the specified button at the specified position relative to the control
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse up</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse up</param>
        /// <param name="button">The button to release</param>
        public override void MouseUp(int X, int Y, MouseButton button)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.MouseUp(X, Y, button);
            }
        }

        /// <summary>
        /// Perform a mouse up with the specified button at the specified position relative to the control while pressing the specified key
        /// </summary>
        /// <param name="X">How far from the left edge of the control to perform a mouse up</param>
        /// <param name="Y">How far from the top edge of the control to perform a mouse up</param>
        /// <param name="button">The button to release</param>
        /// <param name="keys">The key to hold while performing a mouse up</param>
        public override void MouseUp(int X, int Y, MouseButton button, MouseKeyModifier keys)
        {
            if (Identity.TechnologyType == "Windows ActiveX")
            {
                throw new Exception("Implement me!");
            }
            else
            {
                base.MouseUp(X, Y, button, keys);
            }
        }

        /// <summary>
        /// Gets whether the control is currently enabled
        /// </summary>
        public override bool IsEnabled
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the Caption property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Enabled", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    bool enabled = GUI.m_APE.GetValueFromMessage();
                    return enabled;
                }
                else
                {
                    return base.IsEnabled;
                }
            }
        }

        /// <summary>
        /// Gets whether the control currently exists
        /// </summary>
        public override bool Exists
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //TODO do this a better way
                    try
                    {
                        bool enabled = this.IsEnabled;
                        return true;
                    }
                    catch
                    {
                        return false;
                    }
                }
                else
                {
                    return base.Exists;
                }
            }
        }

        /// <summary>
        /// Gets the extended window style of the control
        /// </summary>
        public override long ExtendedStyle
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    return 0;
                }
                else
                {
                    return base.ExtendedStyle;
                }
            }
        }

        /// <summary>
        /// Gets the height of the control
        /// </summary>
        public override int Height
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the height property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Height", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int height = GUI.m_APE.GetValueFromMessage();
                    return TwipsToPixels(height, Direction.Vertical);
                }
                else
                {
                    return base.Height;
                }
            }
        }

        /// <summary>
        /// Gets the controls window handle
        /// </summary>
        public override IntPtr Handle
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    return IntPtr.Zero;
                }
                else
                {
                    return base.Handle;
                }
            }
        }

        /// <summary>
        /// Gets the id of the control
        /// </summary>
        public override int Id
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    return 0;
                }
                else
                {
                    return base.Id;
                }
            }
        }

        /// <summary>
        /// Gets the left edge of the control relative to the screen
        /// </summary>
        public override int Left
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the left property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Left", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int left = GUI.m_APE.GetValueFromMessage();
                    return TwipsToPixels(left, Direction.Horizontal);
                }
                else
                {
                    return base.Left;
                }
            }
        }

        /// <summary>
        /// Gets the window style of the control
        /// </summary>
        public override long Style
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    return 0;
                }
                else
                {
                    return base.Style;
                }
            }
        }

        /// <summary>
        /// Gets the windows current text
        /// </summary>
        public override string Text
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the Caption property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Caption", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string text = GUI.m_APE.GetValueFromMessage();
                    return text;
                }
                else
                {
                    return base.Text;
                }
            }
        }

        /// <summary>
        /// Gets the topedge of the control relative to the screen
        /// </summary>
        public override int Top
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the top property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Top", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int top = GUI.m_APE.GetValueFromMessage();
                    return TwipsToPixels(top, Direction.Vertical);
                }
                else
                {
                    return base.Top;
                }
            }
        }

        /// <summary>
        /// Gets whether the control is currently visible
        /// </summary>
        public override bool IsVisible
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the visible property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Visible", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    bool visible = GUI.m_APE.GetValueFromMessage();
                    return visible;
                }
                else
                {
                    return base.IsVisible;
                }
            }
        }

        /// <summary>
        /// Gets the width of the control
        /// </summary>
        public override int Width
        {
            get
            {
                if (Identity.TechnologyType == "Windows ActiveX")
                {
                    //Get the width property
                    GUI.m_APE.AddFirstMessageFindByUniqueId(DataStores.Store0, Identity.UniqueId);
                    GUI.m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Width", MemberTypes.Property);
                    GUI.m_APE.AddRetrieveMessageGetValue(DataStores.Store1);
                    GUI.m_APE.SendMessages(EventSet.APE);
                    GUI.m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    int width = GUI.m_APE.GetValueFromMessage();
                    return TwipsToPixels(width, Direction.Horizontal);
                }
                else
                {
                    return base.Width;
                }
            }
        }
    }
}
