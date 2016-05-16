using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;
using APE.Dock;
using System.Runtime.InteropServices;
using NM = APE.Native.NativeMethods;

namespace APE.Watcher
{
    public partial class Watcher : AppBar
    {
        public Watcher()
        {
            //InitializeComponent();
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            NM.MoveWindow(this.Handle, 0, 0, 0, 0, false);

            string[] args = Environment.GetCommandLineArgs();
            
            if (args.Length == 3)
            {
                int ParentPid = int.Parse(args[1]);
                uint DoubleClickTimer = uint.Parse(args[2]);

                //TimerResolution.SetTimerResolution(TimerResolution.GetMaximumTimerResolution());
                Process.GetProcessById(ParentPid).WaitForExit();
                //TimerResolution.UnsetTimerResolution();

                uint CurrentDoubleClickTimer = (uint)SystemInformation.DoubleClickTime;

                if (CurrentDoubleClickTimer == 1)   //APE probably messed with it so 
                {
                    if (DoubleClickTimer == 1)      //Really messed up so reset it a reasonable 340
                    {
                        DoubleClickTimer = 340;
                    }
                    NM.SetDoubleClickTime(DoubleClickTimer);   //Put it back to what it was when APE started
                }

                base.AppbarNew(AppBarEdges.Top);
                base.AppbarRemove();

                Environment.Exit(0);
            }
            else
            {
                Environment.Exit(-1);
            }
        }
    }
}
