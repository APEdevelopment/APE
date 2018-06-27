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
using System.Diagnostics;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using static APE.Bridge.Ole;

namespace APE.Bridge
{
    // This COM component needs to be registered with regasm specifiying the /codebase parameter
    [Guid("B8C75701-0A4B-44CC-8E91-28A1DCC9D330")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Ax
    {
        private enum GetAncestorFlags
        {
            /// <summary>
            /// Retrieves the parent window. This does not include the owner, as it does with the GetParent function. 
            /// </summary>
            GetParent = 1,
            /// <summary>
            /// Retrieves the root window by walking the chain of parent windows.
            /// </summary>
            GetRoot = 2,
            /// <summary>
            /// Retrieves the owned root window by walking the chain of parent and owner windows returned by GetParent. 
            /// </summary>
            GetRootOwner = 3
        }

        [DllImport("user32.dll", ExactSpelling = true)]
        private static extern IntPtr GetAncestor(IntPtr hWnd, GetAncestorFlags flags);

        public class Item
        {
            public string UniqueId;
            public IntPtr ParentHandle;
            public string ContainerUniqueId;
            public IntPtr Handle;
            public string Name;
            public string TypeName;
            public string TypeNameSpace;
            public object Control;
            public bool Rendered;

            public Item(IntPtr objectPointer, IntPtr containerObjectPointer, IntPtr handle, string name, object control, bool rendered)
            {
                UniqueId = "A" + objectPointer.ToString();
                ParentHandle = GetAncestor(handle, GetAncestorFlags.GetRoot);
                ContainerUniqueId = "A" + containerObjectPointer.ToString();
                Handle = handle;
                Name = name;
                TypeName = null;
                TypeNameSpace = null;
                Control = control;
                Rendered = rendered;
            }
        }

        public static List<Item> Items = new List<Item>();
        public static readonly object AxItemsLock = new object();
        public static Form InvokeForm;
        private static bool Abort = false;

        //We use dynamic instead of IntPtr as IntPtr causes issues when using late binding
        [MethodImpl(MethodImplOptions.NoOptimization)]
        public void AddItem(dynamic objectPointer, dynamic containerObjectPointer, dynamic handle, string name, object control, bool rendered)
        {
            if (Abort)
            {
                return;
            }

            try
            {
                //We use this form to be able to invoke on to the thread which created the activex object.
                //Since the form is static if the application has more than one gui thread this won't work
                //but thats unlikely so no need to complicate the code unless I come across a case where 
                //its needed.
                if (InvokeForm == null)
                {
                    InvokeForm = new Form();
                    IntPtr ignore = InvokeForm.Handle; //Force the form handle to be created
                }

                IntPtr windowHandle = IntPtr.Zero;

                if (handle == 0)
                {
                    //If we didn't pass in a handle then we have a custom activex control which should support the IOleWindow interface if it has a window
                    IOleWindow controlAsOleWindow = control as IOleWindow;
                    if (controlAsOleWindow != null)
                    {
                        try
                        {
                            controlAsOleWindow.GetWindow(out windowHandle);
                        }
                        catch
                        {
                            //Some controls support the IOleWindow interface but don't actually have windows
                            Marshal.ReleaseComObject(control);
                            return;
                        }
                    }
                    else
                    {
                        Marshal.ReleaseComObject(control);
                        return;
                    }
                }
                else
                {
                    //We passed in a handle so we have a VB Intrinsic control (which doesn't support the IOleWindow interface)
                    windowHandle = new IntPtr(handle);
                }

                Item item = new Item(new IntPtr(objectPointer), new IntPtr(containerObjectPointer), windowHandle, name, control, rendered);
                
                //Debug.WriteLine("Checking: name: " + item.Name + " hwnd: " + item.Handle.ToString() + " parent: " + item.ParentHandle.ToString() + " address: " + ((uint)objectPointer).ToString());
                bool found = false;
                lock (AxItemsLock)
                {
                    int numberOfItems = Items.Count;
                    for (int index = 0; index < numberOfItems; index++)
                    {
                        if (Items[index].UniqueId == item.UniqueId)
                        {
                            found = true;
                            if (item.UniqueId != item.ContainerUniqueId)
                            {
                                Marshal.ReleaseComObject(Items[index].Control);
                                Items.RemoveAt(index);
                                found = false;
                            }
                            break;
                        }
                    }
                    if (!found)
                    {
                        //Debug.WriteLine("Adding: name: " + item.Name + " hwnd: " + item.Handle.ToString() + " parent: " + item.ParentHandle.ToString() + " address: " + ((uint)objectPointer).ToString());
                        Items.Add(item);
                    }
                }
            }
            catch
            {
                //If any errors occur in the code in this class then don't do any more processing so we don't leak
                Abort = true;
            }
        }

        public void RemoveAllItemsFromContainer(dynamic containerToRemoveObjectPointer)
        {
            if (Abort)
            {
                return;
            }

            try
            {
                string containerToRemoveUniqueId = "A" + containerToRemoveObjectPointer.ToString();
                lock (AxItemsLock)
                {
                    int numberOfItems = Items.Count;
                    for (int index = numberOfItems - 1; index > -1; index--)
                    {
                        if (Items[index].ContainerUniqueId == containerToRemoveUniqueId)
                        {
                            Marshal.ReleaseComObject(Items[index].Control);
                            Items.RemoveAt(index);
                        }
                    }
                }
            }
            catch
            {
                //If any errors occur in the code in this class then don't do any more processing so we don't leak
                Abort = true;
            }
        }
    }
}
