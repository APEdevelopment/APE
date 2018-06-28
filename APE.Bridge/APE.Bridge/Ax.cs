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
using System.IO;
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
            public short ContainerScaleMode;

            public Item(string containerUniqueId, short containerScaleMode, string objectUniqueId, IntPtr windowHandle, object control, string objectName, bool rendered)
            {
                ContainerUniqueId = containerUniqueId;
                UniqueId = objectUniqueId;
                ParentHandle = GetAncestor(windowHandle, GetAncestorFlags.GetRoot);
                Handle = windowHandle;
                ContainerScaleMode = containerScaleMode;
                Name = objectName;
                TypeName = null;        //lazy populate this as it can be relatively slow to do
                TypeNameSpace = null;   //lazy populate this as it can be relatively slow to do
                Control = control;
                Rendered = rendered;
            }
        }

        public static List<Item> Items = new List<Item>();
        public static readonly object AxItemsLock = new object();
        public static Form InvokeForm;
        private static bool Abort = false;

        //IntPtr can causes issues with late binding so we use int / long and convert manually
        public void AddItem(int containerPointer, short containerScaleMode, int objectPointer, int objectHandle, string objectName, bool objectRendered)
        {
            IntPtr containerIntPtr = new IntPtr(containerPointer);
            IntPtr objectIntPtr = new IntPtr(objectPointer);
            IntPtr objectHandleIntPtr = new IntPtr(objectHandle);
            AddItem(containerIntPtr, containerScaleMode, objectIntPtr, objectHandleIntPtr, objectName, objectRendered);
        }

        public void AddItem(long containerPointer, short containerScaleMode, long objectPointer, long objectHandle, string objectName, bool objectRendered)
        {
            IntPtr containerIntPtr = new IntPtr(containerPointer);
            IntPtr objectIntPtr = new IntPtr(objectPointer);
            IntPtr objectHandleIntPtr = new IntPtr(objectHandle);
            AddItem(containerIntPtr, containerScaleMode, objectIntPtr, objectHandleIntPtr, objectName, objectRendered);
        }

        private void AddItem(IntPtr containerIntPtr, short containerScaleMode, IntPtr objectIntPtr, IntPtr objectHandleIntPtr, string objectName, bool objectRendered)
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

                string containerUniqueId = "A" + containerIntPtr.ToString();
                string objectUniqueId = "A" + objectIntPtr.ToString();

                lock (AxItemsLock)
                {
                    int numberOfItems = Items.Count;
                    for (int index = 0; index < numberOfItems; index++)
                    {
                        if (Items[index].UniqueId == objectUniqueId)
                        {
                            //Found the item in the list
                            if (containerIntPtr == objectIntPtr)
                            {
                                //The newly found item references its self so make sure the scale mode is updated
                                Items[index].ContainerScaleMode = containerScaleMode;
                                //Debug
                                //File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ActiveX.debug.log", "Adding container: " + objectUniqueId + " Name: " + objectName + Environment.NewLine);
                            }
                            else
                            {
                                //The newly found item references its parents container so update the 
                                //container unique id and name
                                Items[index].ContainerUniqueId = containerUniqueId;
                                Items[index].Name = objectName;
                            }
                            return;
                        }
                    }

                    //Create a unique rcw for the com object (unique so we can call finalrelease without causing issues)
                    object control = Marshal.GetUniqueObjectForIUnknown(objectIntPtr);

                    //Workout the window handle
                    IntPtr windowHandle = IntPtr.Zero;
                    if (objectHandleIntPtr == IntPtr.Zero)
                    {
                        //If we didn't pass in a handle then we have a custom activex control which should support the IOleWindow interface if it has a window
                        IOleWindow controlAsOleWindow = control as IOleWindow;
                        if (controlAsOleWindow == null)
                        {
                            Marshal.FinalReleaseComObject(control);
                            control = null;
                            return;
                        }

                        try
                        {
                            controlAsOleWindow.GetWindow(out windowHandle);
                        }
                        catch
                        {
                            //Some controls support the IOleWindow interface but don't actually have windows
                            Marshal.FinalReleaseComObject(control);
                            controlAsOleWindow = null;
                            control = null;
                            return;
                        }
                        controlAsOleWindow = null;
                    }
                    else
                    {
                        //We passed in a handle so we have a VB Intrinsic control (which doesn't support the IOleWindow interface)
                        windowHandle = objectHandleIntPtr;
                    }

                    //Create the item
                    Item item = new Item(containerUniqueId, containerScaleMode, objectUniqueId, windowHandle, control, objectName, objectRendered);

                    //Add it
                    Items.Add(item);
                    control = null;
                    //Debug
                    //if (containerIntPtr == objectIntPtr)
                    //{
                    //    File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ActiveX.debug.log", "Adding container: " + item.UniqueId + " Name: " + item.Name + Environment.NewLine);
                    //}
                }
            }
            catch (Exception ex)
            {
                //Debug
                //File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ActiveX.debug.log", "AddItem: Name: " + objectName + " Message: " + ex.Message + Environment.NewLine + ex.StackTrace + Environment.NewLine);
                //If any errors occur in the code in this class then don't do any more processing so we don't leak
                Abort = true;
            }
        }

        //IntPtr can causes issues with late binding so we use int / long and convert manually
        public void RemoveAllItemsFromContainer(int containerToRemovePointer, string containerName)
        {
            IntPtr containerToRemoveIntPtr = new IntPtr(containerToRemovePointer);
            RemoveAllItemsFromContainer(containerToRemoveIntPtr, containerName);
        }

        public void RemoveAllItemsFromContainer(long containerToRemovePointer, string containerName)
        {
            IntPtr containerToRemoveIntPtr = new IntPtr(containerToRemovePointer);
            RemoveAllItemsFromContainer(containerToRemoveIntPtr, containerName);
        }

        private void RemoveAllItemsFromContainer(IntPtr containerToRemoveIntPtr, string containerName)
        {
            if (Abort)
            {
                return;
            }

            try
            {
                string containerToRemoveUniqueId = "A" + containerToRemoveIntPtr.ToString();

                //Debug
                //File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ActiveX.debug.log", "Removing container: " + containerToRemoveUniqueId + " Name: " + containerName + Environment.NewLine);

                lock (AxItemsLock)
                {
                    int numberOfItems = Items.Count;
                    for (int index = numberOfItems - 1; index > -1; index--)
                    {
                        if (Items[index].ContainerUniqueId == containerToRemoveUniqueId)
                        {
                            Marshal.FinalReleaseComObject(Items[index].Control);
                            Items[index].Control = null;
                            Items.RemoveAt(index);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //Debug
                //File.AppendAllText(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + @"\ActiveX.debug.log", "RemoveAllItemsFromContainer: Name: " + containerName + " Message: " + ex.Message + Environment.NewLine + ex.StackTrace + Environment.NewLine);
                //If any errors occur in the code in this class then don't do any more processing so we don't leak
                Abort = true;
            }
        }
    }
}
