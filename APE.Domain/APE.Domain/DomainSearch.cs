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
using mscoree;

namespace APE.Domain
{
    public static class DomainSearch
    {
        public static string GetAllAppDomainNames()
        {
            string AppDom = "";
            IntPtr Handle = IntPtr.Zero;
            CorRuntimeHost Host = new CorRuntimeHost();

            try
            {
                Host.EnumDomains(out Handle);
                while (true)
                {
                    object domain;
                    Host.NextDomain(Handle, out domain);
                    if (domain == null)
                    {
                        break;
                    }

                    if (AppDom == "")
                    {
                        if (((AppDomain)domain).IsDefaultAppDomain())
                        {
                            AppDom = "DefaultDomain";
                        }
                        else
                        {
                            AppDom = ((AppDomain)domain).FriendlyName;
                        }
                    }
                    else
                    {
                        if (((AppDomain)domain).IsDefaultAppDomain())
                        {
                            AppDom += "\tDefaultDomain";
                        }
                        else
                        {
                            AppDom += "\t" + ((AppDomain)domain).FriendlyName;
                        }
                    }
                }
            }
            finally
            {
                if (Handle != IntPtr.Zero)
                {
                    Host.CloseEnum(Handle);
                }
            }

            return AppDom;
        }

        public static AppDomain GetAppDomain(string AppDomainName)
        {
            AppDomain AppDom = null;
            IntPtr Handle = IntPtr.Zero;
            CorRuntimeHost Host = new CorRuntimeHost();

            try
            {
                Host.EnumDomains(out Handle);
                while (true)
                {
                    object domain;
                    Host.NextDomain(Handle, out domain);
                    if (domain == null)
                    {
                        break;
                    }

                    if (((AppDomain)domain).FriendlyName == AppDomainName)
                    {
                        AppDom = (AppDomain)domain;
                        break;
                    }
                }
            }
            finally
            {
                if (Handle != IntPtr.Zero)
                {
                    Host.CloseEnum(Handle);
                }
            }

            return AppDom;
        }
    }
}
