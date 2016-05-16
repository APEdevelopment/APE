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
