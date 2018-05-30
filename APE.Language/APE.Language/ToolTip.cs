using System;
using System.Diagnostics;
using System.Threading;
using NM = APE.Native.NativeMethods;
using System.Drawing;

namespace APE.Language
{
    /// <summary>
    /// Automation class used to automate context menu controls
    /// </summary>
    public class GUIToolTip
    {
        private string Description { get; }
        private Rectangle TipRectangle;

        /// <summary>
        /// Gets the controls window handle
        /// </summary>
        public IntPtr Handle { get; }

        /// <summary>
        /// Gets the windows current text
        /// </summary>
        public string Text { get; }

        /// <summary>
        /// Constructor used for Tooltips
        /// </summary>
        /// <param name="descriptionOfControl">A description of the control which would make sense to a human.
        /// <para/>This text is used in the logging method.  For example: OK button</param>
        /// <param name="toolTipHandle">The handle of the tooltip</param>
        /// <param name="toolTipTitle">The title text of the tooltip</param>
        /// <param name="toolTipRectangle">The rectangle of the tooltip</param>
        internal GUIToolTip(string descriptionOfControl, IntPtr toolTipHandle, string toolTipTitle, Rectangle toolTipRectangle)
        {
            Description = descriptionOfControl;
            Handle = toolTipHandle;
            Text = toolTipTitle;
            TipRectangle = toolTipRectangle;
        }

        /// <summary>
        /// Gets whether the control is currently visible
        /// </summary>
        public bool IsVisible
        {
            get
            {
                return NM.IsWindowVisible(Handle);
            }
        }

        /// <summary>
        /// Gets the left edge of the control relative to the screen
        /// </summary>
        public int Left
        {
            get
            {
                return TipRectangle.Left;
            }
        }

        /// <summary>
        /// Gets the topedge of the control relative to the screen
        /// </summary>
        public int Top
        {
            get
            {
                return TipRectangle.Top;
            }
        }

        /// <summary>
        /// Gets the width of the control
        /// </summary>
        public int Width
        {
            get
            {
                return TipRectangle.Width;
            }
        }

        /// <summary>
        /// Gets the height of the control
        /// </summary>
        public int Height
        {
            get
            {
                return TipRectangle.Height;
            }
        }

        /// <summary>
        /// Waits for the control to not be visible
        /// </summary>
        public void WaitForControlToNotBeVisible()
        {
            //Wait for the control to not be visible
            Stopwatch timer = Stopwatch.StartNew();
            while (true)
            {
                if (timer.ElapsedMilliseconds > GUI.m_APE.TimeOut)
                {
                    throw GUI.ApeException(this.Description + " failed to become nonvisible");
                }

                if (!NM.IsWindowVisible(this.Handle))
                {
                    break;
                }

                Thread.Sleep(15);
            }

            // Small sleep to let focus switch
            Thread.Sleep(20);
        }
    }
}

