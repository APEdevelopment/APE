using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to help diagnose issues with interacting with COM objects
    /// </summary>
    public partial class APEIPC
    {
        /// <summary>
        /// Gets the full typename of an object
        /// In the case of .NET objects it is in the form 'Type.Namespace'.'Type.Name'
        /// In the case of COM objects it is in the form 'TypeLibName'.'TypeInfoName'
        /// </summary>
        /// <param name="obj">The object to get the full typename of</param>
        /// <returns>The full type name</returns>
        private string GetObjectFullTypeName(object obj)
        {
            if (Marshal.IsComObject(obj))
            {
                IDispatch dispatch = (IDispatch)obj;
                if (dispatch.GetTypeInfoCount() != 1)
                {
                    throw new Exception("Failed to get runtime type information");
                }

                ITypeInfo typeInfo = dispatch.GetTypeInfo(0, 0);

                ITypeLib typeLib;
                int index;
                typeInfo.GetContainingTypeLib(out typeLib, out index);

                string typeLibName = Marshal.GetTypeLibName(typeLib);
                string typeInfoName = Marshal.GetTypeInfoName(typeInfo);

                return typeLibName + "." + typeInfoName;
            }
            else
            {
                Type typeObject = obj.GetType();
                return typeObject.Namespace + "." + typeObject.Name;
            }
        }

        /// <summary>
        /// A partial IDispatch interface
        /// </summary>
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        private interface IDispatch
        {
            int GetTypeInfoCount();
            [return: MarshalAs(UnmanagedType.Interface)]
            ITypeInfo GetTypeInfo([In, MarshalAs(UnmanagedType.U4)] int iTInfo, [In, MarshalAs(UnmanagedType.U4)] int lcid);
            void GetIDsOfNames([In] ref Guid riid, [In, MarshalAs(UnmanagedType.LPArray)] string[] rgszNames, [In, MarshalAs(UnmanagedType.U4)] int cNames, [In, MarshalAs(UnmanagedType.U4)] int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        }
    }
}
