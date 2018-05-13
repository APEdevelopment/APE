using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace APE.Bridge
{
    [Guid("B8C75701-0A4B-44CC-8E91-28A1DCC9D330")]
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public class Ax
    {
        public class Item
        {
            public IntPtr Handle;
            public string Name;
            public string TypeName;
            public object Control;
            public bool Rendered;

            public Item(IntPtr handle, string name, string typeName, object control, bool rendered)
            {
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
        public void AddItem(dynamic handle, string name, string typeName, object control, bool rendered)
        {
            Item item = new Item(new IntPtr(handle), name, typeName, control, rendered);
            bool found = false;
            lock (AxItemsLock)
            {
                int numberOfItems = Items.Count;
                for (int index = 0; index < numberOfItems; index++)
                {
                    if (Items[index].Handle == item.Handle &&
                        Items[index].Name == item.Name &&
                        Items[index].TypeName == item.TypeName &&
                        Items[index].Rendered == item.Rendered)
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
                    Items.Add(item);
                }
            }
        }

        //We use dynamic instead of IntPtr as IntPtr causes issues when using the dispatch interface
        public void RemoveItem(dynamic handle, string name, string typeName, bool rendered)
        {
            IntPtr handleOfControlToRemove = new IntPtr(handle);
            lock (AxItemsLock)
            {
                int numberOfItems = Items.Count;
                for (int index = 0; index < numberOfItems; index++)
                {
                    if (Items[index].Handle == handleOfControlToRemove &&
                        Items[index].Name == name &&
                        Items[index].TypeName == typeName &&
                        Items[index].Rendered == rendered)
                    {
                        Marshal.ReleaseComObject(Items[index].Control);
                        Items.RemoveAt(index);
                        break;
                    }
                }
            }
        }
    }
}
