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
using System.Runtime.InteropServices;
using NM = APE.Native.NativeMethods;

namespace APE.Watcher
{
    static class TimerResolution
    {
        private static bool m_ResolutionSet = false;
        private static uint m_TimerResolution = 0;
        private static object m_ResolutionLock = new object();

        public static uint GetMaximumTimerResolution()
        {
            uint minimumResolution;
            uint maximumResolution;
            uint currentResolution;

            NM.NtQueryTimerResolution(out minimumResolution, out maximumResolution, out currentResolution);

            return maximumResolution;
        }

        public static bool SetTimerResolution(uint timerResolutionNs)
        {
            bool Succeed = false;
            lock (m_ResolutionLock)
            {
                if (!m_ResolutionSet)
                {
                    uint maximumResolution = GetMaximumTimerResolution();

                    if (timerResolutionNs >= maximumResolution)
                    {
                        uint currentResolution;
                        uint result = NM.NtSetTimerResolution(timerResolutionNs, true, out currentResolution);

                        switch (result)
                        {
                            case NM.STATUS_SUCCESS:
                                m_TimerResolution = timerResolutionNs;
                                m_ResolutionSet = true;
                                Succeed = true;
                                break;
                            default:
                                break;
                        }
                    }
                }
            }

            return Succeed;
        }

        public static bool UnsetTimerResolution()
        {
            bool Succeed = false;
            lock (m_ResolutionLock)
            {
                if (m_ResolutionSet)
                {
                    uint currentResolution;
                    uint result = NM.NtSetTimerResolution(m_TimerResolution, false, out currentResolution);

                    switch (result)
                    {
                        case NM.STATUS_SUCCESS:
                            m_ResolutionSet = false;
                            Succeed = true;
                            break;
                        default:
                            break;
                    }
                }
            }

            return Succeed;
        }
    }
}
