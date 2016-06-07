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
using System.Windows.Forms;
using System.Drawing;
using System.Runtime.InteropServices;
using NM = APE.Native.NativeMethods;
using System.Diagnostics;

namespace APE.Dock
{
    public class AppBar : Form
    {
        private const int PosChanged = 0x00000001;

        private AppBarEdges m_Edge;
        private UInt32 m_CallbackMessageID = 0;
        private Boolean m_IsAppbarDocked = false;
        private Size m_OriginalSize;
        private Point m_OriginalLocation;
        private FormBorderStyle m_PrevBorderStyle;

        public enum AppBarEdges
        {
            Top = 1,
            Bottom = 3,
        }

        public AppBar()
        {
            // Register a unique message as our callback message
            String uniqueMessageString = Guid.NewGuid().ToString();
            m_CallbackMessageID = NM.RegisterWindowMessage(uniqueMessageString);
            if (m_CallbackMessageID == 0)
            {
                throw new Exception("RegisterCallbackMessage failed");
            }
        }

        protected override void WndProc(ref Message msg)
        {
            if (m_IsAppbarDocked)
            {
                if (msg.Msg == m_CallbackMessageID)
                {
                    if (msg.WParam.ToInt32() == PosChanged)
                    {
                        SizeAppBar();
                    }
                }
            }

            base.WndProc(ref msg);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            if (m_IsAppbarDocked)
            {
                SizeAppBar();
            }

            base.OnSizeChanged(e);
        }

        public void AppbarNew(AppBarEdges Edge)
        {
            if (!m_IsAppbarDocked)
            {
                if (m_CallbackMessageID == 0)
                {
                    throw new Exception("CallbackMessageID is 0");
                }

                m_Edge = Edge;

                m_OriginalSize = this.Size;
                m_OriginalLocation = this.Location;
                m_PrevBorderStyle = this.FormBorderStyle;

                this.FormBorderStyle = FormBorderStyle.None;

                // prepare data structure of message
                NM.APPBARDATA msgData = new NM.APPBARDATA();
                msgData.cbSize = (UInt32)Marshal.SizeOf(msgData);
                msgData.hWnd = this.Handle;
                msgData.uCallbackMessage = m_CallbackMessageID;

                // install new appbar
                UIntPtr Return = NM.SHAppBarMessage(NM.AppBarMessages.New, ref msgData);
                if (Return == UIntPtr.Zero)
                {
                    throw new Exception("Failed to add AppBar");
                }
                m_IsAppbarDocked = true;

                SizeAppBar();
            }
        }

        public void AppbarRemove()
        {
            if (m_IsAppbarDocked)
            {
                // prepare data structure of message
                NM.APPBARDATA msgData = new NM.APPBARDATA();
                msgData.cbSize = (UInt32)Marshal.SizeOf(msgData);
                msgData.hWnd = this.Handle;

                // remove appbar
                UIntPtr Return = NM.SHAppBarMessage(NM.AppBarMessages.Remove, ref msgData);
                if (Return == UIntPtr.Zero)
                {
                    throw new Exception("Failed to remove AppBar");
                }

                this.FormBorderStyle = m_PrevBorderStyle;
                this.Size = m_OriginalSize;
                this.Location = m_OriginalLocation;

                m_IsAppbarDocked = false;
            }
        }

        private void AppbarSetPos(ref NM.tagRect appRect)
        {
            float ScreenScalingFactor;

            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr desktop = g.GetHdc();
                int LogicalScreenHeight = NM.GetDeviceCaps(desktop, NM.DeviceCap.VERTRES);
                int PhysicalScreenHeight = NM.GetDeviceCaps(desktop, NM.DeviceCap.DESKTOPVERTRES);
                g.ReleaseHdc();
                ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
            }

            // prepare data structure of message
            NM.APPBARDATA msgData = new NM.APPBARDATA();
            msgData.cbSize = (UInt32)Marshal.SizeOf(msgData);
            msgData.hWnd = this.Handle;
            msgData.uEdge = (UInt32)m_Edge;
            msgData.rc = appRect;

            // set postion for the appbar
            msgData.rc.left = (int)(Math.Round((float)msgData.rc.left * ScreenScalingFactor));
            msgData.rc.right = (int)(Math.Round((float)msgData.rc.right * ScreenScalingFactor));
            msgData.rc.top = (int)(Math.Round((float)msgData.rc.top * ScreenScalingFactor));
            msgData.rc.bottom = (int)(Math.Round((float)msgData.rc.bottom * ScreenScalingFactor));

            UIntPtr Return = NM.SHAppBarMessage(NM.AppBarMessages.SetPos, ref msgData);
            if (Return == UIntPtr.Zero)
            {
                throw new Exception("Failed to set AppBar position and size");
            }

            msgData.rc.left = (int)(Math.Round((float)msgData.rc.left / ScreenScalingFactor));
            msgData.rc.right = (int)(Math.Round((float)msgData.rc.right / ScreenScalingFactor));
            msgData.rc.top = (int)(Math.Round((float)msgData.rc.top / ScreenScalingFactor));
            msgData.rc.bottom = (int)(Math.Round((float)msgData.rc.bottom / ScreenScalingFactor));

            appRect = msgData.rc;
        }

        private void AppbarQueryPos(ref NM.tagRect appRect)
        {
            float ScreenScalingFactor;

            using (Graphics g = Graphics.FromHwnd(IntPtr.Zero))
            {
                IntPtr desktop = g.GetHdc();
                int LogicalScreenHeight = NM.GetDeviceCaps(desktop, NM.DeviceCap.VERTRES);
                int PhysicalScreenHeight = NM.GetDeviceCaps(desktop, NM.DeviceCap.DESKTOPVERTRES);
                g.ReleaseHdc();
                ScreenScalingFactor = (float)PhysicalScreenHeight / (float)LogicalScreenHeight;
            }

            // prepare data structure of message
            NM.APPBARDATA msgData = new NM.APPBARDATA();
            msgData.cbSize = (UInt32)Marshal.SizeOf(msgData);
            msgData.hWnd = Handle;
            msgData.uEdge = (UInt32)m_Edge;
            msgData.rc = appRect;

            // query postion for the appbar
            msgData.rc.left = (int)(Math.Round((float)msgData.rc.left * ScreenScalingFactor));
            msgData.rc.right = (int)(Math.Round((float)msgData.rc.right * ScreenScalingFactor));
            msgData.rc.top = (int)(Math.Round((float)msgData.rc.top * ScreenScalingFactor));
            msgData.rc.bottom = (int)(Math.Round((float)msgData.rc.bottom * ScreenScalingFactor));

            NM.SHAppBarMessage(NM.AppBarMessages.QueryPos, ref msgData);

            msgData.rc.left = (int)(Math.Round((float)msgData.rc.left / ScreenScalingFactor));
            msgData.rc.right = (int)(Math.Round((float)msgData.rc.right / ScreenScalingFactor));
            msgData.rc.top = (int)(Math.Round((float)msgData.rc.top / ScreenScalingFactor));
            msgData.rc.bottom = (int)(Math.Round((float)msgData.rc.bottom / ScreenScalingFactor));

            appRect = msgData.rc;
        }

        private void SizeAppBar()
        {
            NM.tagRect rt = new NM.tagRect();

            // Set the rect to where we would like to place the AppBar
            rt.right = SystemInformation.PrimaryMonitorSize.Width;
            rt.left = 0;
            if (m_Edge == AppBarEdges.Top)
            {
                rt.bottom = this.Size.Height;
                rt.top = 0;
            }
            else
            {
                rt.bottom = SystemInformation.PrimaryMonitorSize.Height;
                rt.top = rt.bottom - this.Size.Height;
            }

            // Query to see if there are any other AppBars in the way
            AppbarQueryPos(ref rt);

            //If our position clashes with another AppBar then then adjust our position
            if (m_Edge == AppBarEdges.Top)
            {
                rt.bottom = rt.top + this.Size.Height;
            }
            else
            {
                rt.top = rt.bottom - this.Size.Height;
            }

            AppbarSetPos(ref rt);

            this.Location = new Point(rt.left, rt.top);
            this.Size = new Size(rt.right - rt.left, rt.bottom - rt.top);
        }
    }
}