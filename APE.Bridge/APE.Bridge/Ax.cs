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
using System.Runtime.InteropServices;
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
            public IntPtr ContainerHandle;
            public IntPtr Handle;
            public string Name;
            public string TypeName;
            public string TypeNameSpace;
            public object Control;
            public bool Rendered;

            public Item(IntPtr objectPointer, IntPtr containerHandle, IntPtr handle, string name, object control, bool rendered)
            {
                UniqueId = "A" + objectPointer.ToString();
                ParentHandle = GetAncestor(handle, GetAncestorFlags.GetRoot);
                ContainerHandle = containerHandle;
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

        //We use dynamic instead of IntPtr as IntPtr causes issues when using late binding
        public void AddItem(dynamic objectPointer, dynamic containerHandle, dynamic handle, string name, object control, bool rendered)
        {
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
                        //Debug.WriteLine("No window: name: " + name + " typename: " + typeName);
                        Marshal.ReleaseComObject(control);
                        return;
                    }
                }
                else
                {
                    //Debug.WriteLine("Not an IOleWindow: name: " + name + " typename: " + typeName);
                    Marshal.ReleaseComObject(control);
                    return;
                }
            }
            else
            {
                //We passed in a handle so we have a VB Intrinsic control (which doesn't support the IOleWindow interface)
                windowHandle = new IntPtr(handle);
            }

            Item item = new Item(new IntPtr(objectPointer), new IntPtr(containerHandle), windowHandle, name, control, rendered);
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
                        //Debug.Write("Updating: name: " + Items[index].Name + " now:");
                        //Update the handles and name if need be
                        if (string.IsNullOrEmpty(Items[index].Name))
                        {
                            Items[index].Name = item.Name;
                            //Debug.Write(" name: " + Items[index].Name);
                        }
                        if (Items[index].Handle == IntPtr.Zero)
                        {
                            Items[index].Handle = item.Handle;
                            Items[index].ParentHandle = item.ParentHandle;
                            //Debug.Write(" hwnd: " + item.Handle.ToString() + " parent: " + item.ParentHandle.ToString());
                        }
                        //Debug.WriteLine("");
                        break;
                    }
                }
                if (found)
                {
                    //Already added this so decrement the RCW reference count
                    Marshal.ReleaseComObject(item.Control);
                }
                else
                {
                    //Debug.WriteLine("Adding: name: " + item.Name + " hwnd: " + item.Handle.ToString() + " parent: " + item.ParentHandle.ToString() + " address: " + ((uint)objectPointer).ToString());
                    Items.Add(item);
                }
            }
        }

        public void RemoveAllItemsFromContainer(dynamic containerHandle)
        {
            IntPtr handleOfContainerToRemove = new IntPtr(containerHandle);
            lock (AxItemsLock)
            {
                int numberOfItems = Items.Count;
                for (int index = numberOfItems - 1; index > -1; index--)
                {
                    if (Items[index].ParentHandle == handleOfContainerToRemove ||   //We check the ParentHandle to be certain that we don't leak any controls, its not really needed but doesn't hurt
                        Items[index].ContainerHandle == handleOfContainerToRemove ||
                        Items[index].Handle == handleOfContainerToRemove)
                    {
                        //Debug.WriteLine("removing: " + Items[index].Name + " parent: " + Items[index].ParentHandle.ToString());
                        Marshal.ReleaseComObject(Items[index].Control);
                        Items.RemoveAt(index);
                    }
                }
            }
        }
    }
}
