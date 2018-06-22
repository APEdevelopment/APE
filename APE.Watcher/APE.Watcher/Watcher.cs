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
using System.Diagnostics;
using APE.Dock;
using NM = APE.Native.NativeMethods;

namespace APE.Watcher
{
    public class Watcher : AppBar
    {
        public Watcher()
        {
            this.FormBorderStyle = FormBorderStyle.None;
            NM.MoveWindow(this.Handle, 0, 0, 0, 0, false);
            this.Load += new EventHandler(this.Watcher_Load);

            string[] args = Environment.GetCommandLineArgs();

            if (args.Length == 3)
            {
                int ParentPid = int.Parse(args[1]);
                uint DoubleClickTimer = uint.Parse(args[2]);

                Process.GetProcessById(ParentPid).WaitForExit();

                uint CurrentDoubleClickTimer = (uint)SystemInformation.DoubleClickTime;

                if (CurrentDoubleClickTimer <= NM.IntraClickDelay)   //APE probably messed with it so 
                {
                    if (DoubleClickTimer <= NM.IntraClickDelay)      //Really messed up so reset it a reasonable 340
                    {
                        DoubleClickTimer = 340;
                    }
                    NM.SetDoubleClickTime(DoubleClickTimer);   //Put it back to what it was when APE started
                }

                base.AppbarNew(AppBarEdges.Top);    //Add and remove an appbar to make sure they are reset if APE is killed off
                base.AppbarRemove();

                Program.ExitCode = 0;
            }
            else
            {
                Program.ExitCode = 1;
            }
        }

        private void Watcher_Load(object sender, EventArgs e)
        {
            //Cant exit in the constructior so exit in the load event instead
            Application.Exit();
        }
    }
}
