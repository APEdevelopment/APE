using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace APE.Bridge
{
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
            public object Control;
            public bool Rendered;

            public Item(IntPtr objectPointer, IntPtr containerHandle, IntPtr handle, string name, string typeName, object control, bool rendered)
            {
                UniqueId = "A" + objectPointer.ToString();
                ParentHandle = GetAncestor(handle, GetAncestorFlags.GetRoot);
                ContainerHandle = containerHandle;
                Handle = handle;
                Name = name;
                TypeName = typeName;
                Control = control;
                Rendered = rendered;
            }
        }

        public static List<Item> Items = new List<Item>();
        public static readonly object AxItemsLock = new object();

        //We use dynamic instead of IntPtr as IntPtr causes issues when using the dispatch interface

        public void AddItem(dynamic objectPointer, dynamic containerHandle, dynamic handle, string name, string typeName, object control, bool rendered)
        {
            Item item = new Item(new IntPtr(objectPointer), new IntPtr(containerHandle), new IntPtr(handle), name, typeName, control, rendered);
            bool found = false;
            lock (AxItemsLock)
            {
                int numberOfItems = Items.Count;
                for (int index = 0; index < numberOfItems; index++)
                {
                    if (Items[index].UniqueId == item.UniqueId)
                    {
                        found = true;
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
                    Debug.WriteLine("Adding: " + item.Name + " parent: " + item.ParentHandle.ToString());
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
                        Debug.WriteLine("removing: " + Items[index].Name + " parent: " + Items[index].ParentHandle.ToString());
                        Marshal.ReleaseComObject(Items[index].Control);
                        Items.RemoveAt(index);
                    }
                }
            }
        }
    }
}
