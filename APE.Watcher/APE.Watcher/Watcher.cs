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

                base.AppbarNew(AppBarEdges.Top);    //Add and remove an appbar to make sure they are reset if APE is killed off
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
