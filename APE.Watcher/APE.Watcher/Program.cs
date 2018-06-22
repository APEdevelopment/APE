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
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APE.Watcher
{
    static class Program
    {
        private static int m_ExitCode = 0;

        internal static int ExitCode
        {
            get
            {
                return m_ExitCode;
            }
            set
            {
                if (value < 0)
                {
                    throw new Exception("exitCode must be positive");
                }
                m_ExitCode = value;
            }
        }

        [STAThread]
        private static int Main()
        {
            Application.Run(new Watcher());
            return ExitCode;
        }
    }
}
