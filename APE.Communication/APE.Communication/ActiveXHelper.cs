using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using APE.Bridge;
using NM = APE.Native.NativeMethods;
using WF = System.Windows.Forms;

namespace APE.Communication
{
    public partial class APEIPC
    {
        private object FindByHandleActiveX(IntPtr handle, out string name, out string typeName, out string uniqueId)
        {
            if (Ax.Items.Count > 0)
            {
                lock (Ax.AxItemsLock)
                {
                    int items = Ax.Items.Count;
                    for (int item = 0; item < items; item++)
                    {
                        if (Ax.Items[item].Handle == handle)
                        {
                            name = Ax.Items[item].Name;
                            typeName = Ax.Items[item].TypeName;
                            uniqueId = Ax.Items[item].UniqueId;
                            return Ax.Items[item].Control;
                        }
                    }
                }
            }
            name = null;
            typeName = null;
            uniqueId = null;
            return null;
        }

        private object FindByUniqueIdActiveX(string uniqueId, out string name, out string typeName, out IntPtr handle)
        {
            if (Ax.Items.Count > 0)
            {
                lock (Ax.AxItemsLock)
                {
                    int items = Ax.Items.Count;
                    for (int item = 0; item < items; item++)
                    {
                        if (Ax.Items[item].UniqueId == uniqueId)
                        {
                            name = Ax.Items[item].Name;
                            typeName = Ax.Items[item].TypeName;
                            handle = Ax.Items[item].Handle;
                            return Ax.Items[item].Control;
                        }
                    }
                }
            }
            name = null;
            typeName = null;
            handle = IntPtr.Zero;
            return null;
        }

        private void FindByIdentifierActiveX(ControlIdentifier identifier, ref IntPtr handle, ref string name, ref string theText, ref string typeName, ref int currentIndex, ref string uniqueId, ref bool foundControl)
        {
            bool found = false;
            IntPtr parentHandle = IntPtr.Zero;

            if (Ax.Items.Count > 0)
            {
                lock (Ax.AxItemsLock)
                {
                    int items = Ax.Items.Count;
                    for (int item = 0; item < items; item++)
                    {
                        if (Ax.Items[item].Handle == handle)
                        {
                            parentHandle = Ax.Items[item].ParentHandle;
                            name = Ax.Items[item].Name;
                            typeName = Ax.Items[item].TypeName;
                            uniqueId = Ax.Items[item].UniqueId;
                            found = true;
                            break;
                        }
                    }
                }
            }

            if (!found)
            {
                return;
            }

            if (identifier.ParentHandle == parentHandle || (identifier.ParentHandle == IntPtr.Zero && parentHandle == handle))
            {
            }
            else
            {
                return;
            }

            if (identifier.Name != null)
            {
                if (name != identifier.Name)
                {
                    return;
                }
            }

            if (identifier.TechnologyType != null)
            {
                if ("Windows ActiveX" != identifier.TechnologyType)
                {
                    return;
                }
            }

            if (identifier.TypeNameSpace != null)
            {
                return;
            }

            if (identifier.TypeName != null)
            {
                if (typeName != identifier.TypeName)
                {
                    return;
                }
            }

            if (identifier.ModuleName != null)
            {
                if (Path.GetFileName(NM.GetWindowModuleFileName(handle)) != identifier.ModuleName)
                {
                    return;
                }
            }

            if (identifier.AssemblyName != null)
            {
                return;
            }

            if (identifier.ChildOf != IntPtr.Zero)
            {
                if (!NM.IsChild(identifier.ChildOf, handle))
                {
                    return;
                }
            }

            if (identifier.SiblingOf != IntPtr.Zero)
            {
                if (!NM.IsSibling(identifier.SiblingOf, handle))
                {
                    return;
                }
            }

            if (identifier.ParentOf != IntPtr.Zero)
            {
                if (!NM.IsChild(handle, identifier.ParentOf))
                {
                    return;
                }
            }

            theText = GetWindowTextViaWindowMessage(handle);
            if (identifier.Text != null)
            {
                if (theText == null)
                {
                    return;
                }
                else
                {
                    if (!Regex.IsMatch(theText, identifier.Text))
                    {
                        return;
                    }
                }
            }

            if (identifier.AccessibilityObjectName != null)
            {
                return;
            }

            currentIndex++;
            if (identifier.Index > 0)
            {
                if (currentIndex != identifier.Index)
                {
                    return;
                }
            }
            
            //we have a match
            foundControl = true;
        }

        private void FindByIdentifierRenderedActiveX(ControlIdentifier identifier, ref IntPtr handle, ref string name, ref string theText, ref string typeName, ref string technologyType, ref string uniqueId, ref bool foundControl)
        {
            IntPtr parentHandle = IntPtr.Zero;
            int currentIndex = 0;

            if (Ax.Items.Count > 0)
            {
                lock (Ax.AxItemsLock)
                {
                    int items = Ax.Items.Count;
                    for (int item = 0; item < items; item++)
                    {
                        handle = Ax.Items[item].Handle;
                        parentHandle = Ax.Items[item].ParentHandle;
                        name = Ax.Items[item].Name;
                        typeName = Ax.Items[item].TypeName;
                        uniqueId = Ax.Items[item].UniqueId;

                        if (identifier.ParentHandle == parentHandle || (identifier.ParentHandle == IntPtr.Zero && parentHandle == handle))
                        {
                        }
                        else
                        {
                            continue;
                        }

                        //if (!Ax.Items[item].Rendered)
                        //{
                        //    continue;
                        //}

                        if (identifier.ChildOf != IntPtr.Zero)
                        {
                            if (identifier.ChildOf == handle || NM.IsChild(identifier.ChildOf, handle))
                            {
                            }
                            else
                            {
                                continue;
                            }
                        }

                        //if (identifier.SiblingOf != IntPtr.Zero)
                        //{
                        //    if (!NM.IsSibling(identifier.SiblingOf, handle))
                        //    {
                        //        continue;
                        //    }
                        //}

                        //if (identifier.ParentOf != IntPtr.Zero)
                        //{
                        //    if (!NM.IsChild(handle, identifier.ParentOf))
                        //    {
                        //        continue;
                        //    }
                        //}

                        if (identifier.ModuleName != null)
                        {
                            if (Path.GetFileName(NM.GetWindowModuleFileName(handle)) != identifier.ModuleName)
                            {
                                continue;
                            }
                        }

                        if (identifier.TypeName != null)
                        {
                            if (typeName != identifier.TypeName)
                            {
                                continue;
                            }
                        }

                        if (identifier.Name != null)
                        {
                            if (name != identifier.Name)
                            {
                                continue;
                            }
                        }

                        //object[] comParameters = new object[3];
                        //comParameters[0] = Name;
                        //comParameters[1] = SourceObject;
                        //comParameters[2] = ParametersObject;
                        //DestinationObject = ((WF.Control)tempStore0).Invoke(m_ComReflectDelegater, comParameters);
                        theText = (string)m_ComReflectDelegater.Invoke("Caption", Ax.Items[item].Control, null);
                        if (identifier.Text != null)
                        {
                            if (theText == null)
                            {
                                continue;
                            }
                            else
                            {
                                if (!Regex.IsMatch(theText, identifier.Text))
                                {
                                    continue;
                                }
                            }
                        }

                        //accessibilityObjectName = (string)form.Invoke(m_GetAccessibilityObjectNameDelegater, parameters);
                        //if (Identifier.AccessibilityObjectName != null)
                        //{
                        //    if (accessibilityObjectName != Identifier.AccessibilityObjectName)
                        //    {
                        //        continue;
                        //    }
                        //}

                        currentIndex++;

                        if (identifier.Index > 0)
                        {
                            if (currentIndex != identifier.Index)
                            {
                                continue;
                            }
                        }

                        technologyType = "Windows ActiveX";
                        //WF.MessageBox.Show("Found handle: " + name + " " + handle.ToString());
                        foundControl = true;
                        return;
                    }
                }
            }
        }
    }
}
