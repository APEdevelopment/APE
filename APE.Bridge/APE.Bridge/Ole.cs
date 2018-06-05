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
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using ComTypes = System.Runtime.InteropServices.ComTypes;

namespace APE.Bridge
{
    public class Ole
    {
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("00020400-0000-0000-C000-000000000046")]
        internal interface IDispatch
        {
            uint GetTypeInfoCount();

            [return: MarshalAs(UnmanagedType.Interface)]
            ComTypes.ITypeInfo GetTypeInfo(uint iTInfo, int lcid);

            [return: MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.I4, SizeParamIndex = 2)]
            int[] GetIDsOfNames(
                ref Guid iid,
                [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr, SizeParamIndex = 2)]
                string[] names,
                uint cNames,
                int lcid);

            void Invoke(
                int dispIdMember,
                ref Guid riid,
                int lcid,
                ComTypes.INVOKEKIND wFlags,
                ref ComTypes.DISPPARAMS pDispParams,
                IntPtr pvarResult,
                IntPtr pExcepInfo,
                IntPtr puArgErr);
        }

        [ComConversionLoss]
        [Guid("00000114-0000-0000-C000-000000000046")]
        [InterfaceType(1)]
        [ComImport]
        public interface IOleWindow
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetWindow(out IntPtr phwnd);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void ContextSensitiveHelp([In] int fEnterMode);
        }

        [ComConversionLoss]
        [Guid("00000112-0000-0000-C000-000000000046")]
        [InterfaceType(1)]
        [ComImport]
        public interface IOleObject
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetClientSite([MarshalAs(UnmanagedType.Interface), In] IOleClientSite pClientSite);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetClientSite([MarshalAs(UnmanagedType.Interface)] out IOleClientSite ppClientSite);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetHostNames([MarshalAs(UnmanagedType.LPWStr), In] string szContainerApp, [MarshalAs(UnmanagedType.LPWStr), In] string szContainerObj);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Close([In] uint dwSaveOption);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetMoniker([In] uint dwWhichMoniker, [MarshalAs(UnmanagedType.Interface), In] ComTypes.IMoniker pmk);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [MarshalAs(UnmanagedType.Interface)] out ComTypes.IMoniker ppmk);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int InitFromData([MarshalAs(UnmanagedType.Interface), In] ComTypes.IDataObject pDataObject, [In] int fCreation, [In] uint dwReserved);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetClipboardData([In] uint dwReserved, [MarshalAs(UnmanagedType.Interface)] out ComTypes.IDataObject ppDataObject);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int DoVerb([In] int iVerb, [MarshalAs(UnmanagedType.LPArray), In] MSG[] lpmsg, [MarshalAs(UnmanagedType.Interface), In] IOleClientSite pActiveSite, [In] int lindex, [In] IntPtr hWndParent, [MarshalAs(UnmanagedType.LPArray), In] RECT[] lprcPosRect);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int EnumVerbs([MarshalAs(UnmanagedType.Interface)] out IEnumOLEVERB ppEnumOleVerb);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Update();

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int IsUpToDate();

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetUserClassID(out Guid pClsid);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetUserType([In] uint dwFormOfType, [In] IntPtr pszUserType);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetExtent([In] uint dwDrawAspect, [MarshalAs(UnmanagedType.LPArray), In] SIZEL[] pSizel);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetExtent([In] uint dwDrawAspect, [MarshalAs(UnmanagedType.LPArray), Out] SIZEL[] pSizel);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Advise([MarshalAs(UnmanagedType.Interface), In] ComTypes.IAdviseSink pAdvSink, out uint pdwConnection);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Unadvise([In] uint dwConnection);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void EnumAdvise([MarshalAs(UnmanagedType.Interface)] out ComTypes.IEnumSTATDATA ppenumAdvise);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int GetMiscStatus([In] uint dwAspect, out uint pdwStatus);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SetColorScheme([MarshalAs(UnmanagedType.LPArray), In] LOGPALETTE[] pLogpal);
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;
        }

        [Guid("00000104-0000-0000-C000-000000000046")]
        [InterfaceType(1)]
        [ComImport]
        public interface IEnumOLEVERB
        {
            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Next([In] uint celt, [MarshalAs(UnmanagedType.LPArray), Out] OLEVERB[] rgelt, out uint pceltFetched);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Skip([In] uint celt);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Reset();

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Clone([MarshalAs(UnmanagedType.Interface)] out IEnumOLEVERB ppEnum);
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct OLEVERB
        {
            public int lVerb;
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpszVerbName;
            public uint fuFlags;
            public uint grfAttribs;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct SIZEL
        {
            public int cx;
            public int cy;
        }

        [ComConversionLoss]
        [StructLayout(LayoutKind.Sequential, Pack = 2)]
        public struct LOGPALETTE
        {
            public ushort palVersion;
            public ushort palNumEntries;
            [ComConversionLoss]
            public IntPtr palPalEntry;
        }

        [ComConversionLoss]
        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct MSG
        {
            [ComConversionLoss]
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public uint time;
            public POINT pt;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 4)]
        public struct POINT
        {
            public int x;
            public int y;
        }

        [InterfaceType(1)]
        [Guid("00000118-0000-0000-C000-000000000046")]
        [ComImport]
        public interface IOleClientSite
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void SaveObject();

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetMoniker([In] uint dwAssign, [In] uint dwWhichMoniker, [MarshalAs(UnmanagedType.Interface)] out ComTypes.IMoniker ppmk);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void GetContainer([MarshalAs(UnmanagedType.Interface)] out IOleContainer ppContainer);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void ShowObject();

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void OnShowWindow([In] int fShow);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void RequestNewObjectLayout();
        }

        [Guid("0000011B-0000-0000-C000-000000000046")]
        [InterfaceType(1)]
        [ComImport]
        public interface IOleContainer : IParseDisplayName
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            new void ParseDisplayName([MarshalAs(UnmanagedType.Interface), In] ComTypes.IBindCtx pbc, [MarshalAs(UnmanagedType.LPWStr), In] string pszDisplayName, out uint pchEaten, [MarshalAs(UnmanagedType.Interface)] out ComTypes.IMoniker ppmkOut);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void EnumObjects([In] uint grfFlags, [MarshalAs(UnmanagedType.Interface)] out IEnumUnknown ppEnum);

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void LockContainer([In] int fLock);
        }

        [InterfaceType(1)]
        [Guid("0000011A-0000-0000-C000-000000000046")]
        [ComImport]
        public interface IParseDisplayName
        {
            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void ParseDisplayName([MarshalAs(UnmanagedType.Interface), In] ComTypes.IBindCtx pbc, [MarshalAs(UnmanagedType.LPWStr), In] string pszDisplayName, out uint pchEaten, [MarshalAs(UnmanagedType.Interface)] out ComTypes.IMoniker ppmkOut);
        }

        [InterfaceType(1)]
        [Guid("00000100-0000-0000-C000-000000000046")]
        [ComImport]
        public interface IEnumUnknown
        {
            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Next([In] uint celt, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.IUnknown), Out] object[] rgelt, out uint pceltFetched);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Skip([In] uint celt);

            [MethodImpl(MethodImplOptions.PreserveSig | MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            int Reset();

            [MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime)]
            void Clone([MarshalAs(UnmanagedType.Interface)] out IEnumUnknown ppEnum);
        }

        private static Type TypeOfTypeAttr = typeof(ComTypes.TYPEATTR);

        public static void ComTypeInformation(object comObject, out string interfaceName, out string typeLibraryName, out string className)
        {
            interfaceName = null;
            typeLibraryName = null;
            className = null;

            IDispatch comObjectAsDispatch = comObject as IDispatch;
            if (comObjectAsDispatch == null)
            {
                return;
            }
            if (comObjectAsDispatch.GetTypeInfoCount() != 1)
            {
                return;
            }

            //Get the interface name
            ComTypes.ITypeInfo comObjectTypeInformation = comObjectAsDispatch.GetTypeInfo(0, 0);
            interfaceName = Marshal.GetTypeInfoName(comObjectTypeInformation);

            //Get the type library name
            ComTypes.ITypeLib typeLibrary;
            try
            {
                comObjectTypeInformation.GetContainingTypeLib(out typeLibrary, out int index);
                typeLibraryName = Marshal.GetTypeLibName(typeLibrary);
            }
            catch
            {
                return;
            }

            //Get the class name
            IOleObject comObjectAsOleObject = comObject as IOleObject;
            if (comObjectAsOleObject == null)
            {
                return;
            }

            comObjectAsOleObject.GetUserClassID(out Guid classId);

            int typeInformationCount = typeLibrary.GetTypeInfoCount();
            ComTypes.ITypeInfo typeInformation;
            ComTypes.TYPEATTR typeAttributes;
            IntPtr pointerTypeAttr = IntPtr.Zero;

            for (int index = typeInformationCount - 1; index > -1; index--)
            {
                typeLibrary.GetTypeInfo(index, out typeInformation);
                try
                {
                    typeInformation.GetTypeAttr(out pointerTypeAttr);
                    typeAttributes = (ComTypes.TYPEATTR)Marshal.PtrToStructure(pointerTypeAttr, TypeOfTypeAttr);
                    if (classId == typeAttributes.guid)
                    {
                        className = Marshal.GetTypeInfoName(typeInformation);
                        return;
                    }
                }
                finally
                {
                    if (pointerTypeAttr != IntPtr.Zero)
                    {
                        typeInformation.ReleaseTypeAttr(pointerTypeAttr);
                        pointerTypeAttr = IntPtr.Zero;
                    }
                }
            }
        }
    }
}
