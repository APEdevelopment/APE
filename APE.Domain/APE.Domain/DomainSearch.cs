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
