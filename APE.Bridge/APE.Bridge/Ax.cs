using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace APE.Bridge
{
    [Guid("B8C75701-0A4B-44CC-8E91-28A1DCC9D330")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Ax
    {
        public class Item
        {
            public IntPtr Handle;
            public string Name;
            public string TypeName;
            public object Control;

            public Item(IntPtr handle, string name, string typeName, object control)
            {
                Handle = handle;
                Name = name;
                TypeName = typeName;
                Control = control;
            }
        }

        public static List<Item> Items = new List<Item>();
        public static readonly object AxItemsLock = new object();

        public void AddItem(IntPtr handle, string name, string typeName, object control)
        {
            Item item = new Item(handle, name, typeName, control);
            lock (AxItemsLock)
            {
                Items.Add(item);
            }
        }

        public void RemoveItem(IntPtr handle)
        {
            lock (AxItemsLock)
            {
                int numberOfItems = Items.Count;
                for (int index = 0; index < numberOfItems; index++)
                {
                    if (Items[index].Handle == handle)
                    {
                        Items.RemoveAt(index);
                        break;
                    }
                }
            }
        }
    }
}
