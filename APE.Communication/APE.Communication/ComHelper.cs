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
using System.Drawing;
using System.Drawing.Imaging;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace APE.Communication
{
    /// <summary>
    /// This partial class provides some in process helper methods to help diagnose issues with interacting with COM objects
    /// </summary>
    public partial class APEIPC
    {
        private delegate object ComReflectDelegate(string name, object sourceObject, object[] parameters);
        private ComReflectDelegate m_ComReflectDelegater;

        /// <summary>
        /// Initialise the sentinel grids helper delegates for use
        /// </summary>
        private void SetupComHelperDelegates()
        {
            m_ComReflectDelegater = new ComReflectDelegate(ComReflectInternal);
        }

        //
        //  GetTypeInformationActiveX
        //

        /// <summary>
        /// Get the type information and puts it in the specified datastore
        /// </summary>
        /// <param name="sourceStore">The datastore which contains the grid object</param>
        /// <param name="destinationStore">The datastore to put the resultant string into</param>
        unsafe public void AddQueryMessageGetTypeInformationActiveX(DataStores sourceStore, DataStores destinationStore)
        {
            if (!m_DoneFind)
            {
                throw new Exception("Must locate the activex control before trying to use it");
            }

            Message* ptrMessage = GetPointerToNextMessage();
            ptrMessage->SourceStore = sourceStore;
            ptrMessage->DestinationStore = destinationStore;

            ptrMessage->Action = MessageAction.GetTypeInformationActiveX;

            m_PtrMessageStore->NumberOfMessages++;
            m_DoneQuery = true;
        }

        /// <summary>
        /// Get the type information and puts it in the specified datastore
        /// </summary>
        /// <param name="ptrMessage">A pointer to the message</param>
        unsafe private void GetTypeInformationActiveX(Message* ptrMessage)
        {
            object sourceObject = GetObjectFromDatastore(ptrMessage->SourceStore);
            object destinationObject = null;

            if (sourceObject != null)
            {
                destinationObject = GetObjectFullTypeName(sourceObject);
            }

            PutObjectInDatastore(ptrMessage->DestinationStore, destinationObject);
            CleanUpMessage(ptrMessage);
        }

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
        /// Calls the specified method / property on the com object on the gui thread
        /// </summary>
        /// <param name="name">The name of the method / property to call</param>
        /// <param name="sourceObject">The com object</param>
        /// <param name="parameters">The parameters to pass to the com object</param>
        /// <returns>The object returned by the invoke</returns>
        private object ComReflectInternal(string name, object sourceObject, object[] parameters)
        {
            return sourceObject.GetType().InvokeMember(name, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, sourceObject, parameters);
        }

        /// <summary>
        /// Converts an ole picture into a .net image
        /// </summary>
        /// <param name="picture">The ole picture</param>
        /// <returns>A .net image</returns>
        internal Image OlePictureToImage(stdole.Picture picture)
        {
            Image image = null;
            switch (picture.Type)
            {
                case Ole.PICTYPE_ICON:
                    Icon icon = (Icon)Icon.FromHandle(new IntPtr(picture.Handle)).Clone();
                    image = icon.ToBitmap();
                    break;
                case Ole.PICTYPE_BITMAP:
                    image = (Image)Image.FromHbitmap(new IntPtr(picture.Handle), new IntPtr(picture.hPal)).Clone();
                    break;
                case Ole.PICTYPE_METAFILE:
                case Ole.PICTYPE_ENHMETAFILE:
                    image = (Image)new Metafile(new IntPtr(picture.Handle), true).Clone();
                    break;
                case Ole.PICTYPE_NONE:
                case Ole.PICTYPE_UNINITIALIZED:
                    break;
                default:
                    throw new Exception("Unsupported picture type: " + picture.Type.ToString());
            }
            return image;
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

        private class Ole
        {
            public const int PICTYPE_UNINITIALIZED = -1;
            public const int PICTYPE_NONE = 0;
            public const int PICTYPE_BITMAP = 1;
            public const int PICTYPE_METAFILE = 2;
            public const int PICTYPE_ICON = 3;
            public const int PICTYPE_ENHMETAFILE = 4;
        }
    }
}
