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
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using System.Diagnostics;
using APE.Communication;
using APE.Capture;
using System.Reflection;
using System.Threading;
using NM = APE.Native.NativeMethods;
using System.Xml;
using System.Text.RegularExpressions;
using System.Text;
using APE.Language;

namespace APE.Spy
{
    public partial class APESpy : Form
    {
        Dictionary<Process, String> ListOfProcesses;
        KeyValuePair<Process, string> m_CurrentAttached = new KeyValuePair<Process, string>(Process.GetCurrentProcess(), "");
        Dictionary<IntPtr, String> ListOfTopLevelWindows;
        APEIPC m_APE;
        Dictionary<uint, uint> m_Pid;
        KeyValuePair<Process, string> m_SelectedItem = new KeyValuePair<Process, string>(Process.GetCurrentProcess(), "");
        string m_ObjectCodeLocator = "";

        public APESpy()
        {
            InitializeComponent();
            this.Icon = Properties.Resources.APESPY256x256;
            WindowTree.Height = PropertyListbox.Height;
            RefreshProcesses();
        }

        private void ObjectSpy_FormClosing(Object sender, FormClosingEventArgs e)
        {
            m_CurrentAttached = new KeyValuePair<Process, string>(Process.GetCurrentProcess(), "");
            m_ControlKey = true;
            m_Closing = true;
            if (m_APE != null)
            {
                m_APE.RemoveFileMapping();
            }
        }

        private void RefreshProcesses()
        {
            WinformsProcessesCombobox.Enabled = false;

            m_Pid = new Dictionary<uint, uint>();
            NM.EnumWindowsProc WindowsGetProcessesCallback = new NM.EnumWindowsProc(EnumProcToGetProcesses);
            NM.EnumWindows(WindowsGetProcessesCallback, IntPtr.Zero);

            Process[] pc = Process.GetProcesses();

            ListOfProcesses = new Dictionary<Process, string>();
            ListOfProcesses.Add(Process.GetCurrentProcess(), "");

            foreach (Process p in pc)
            {
                if (p.Id != Process.GetCurrentProcess().Id)
                {
                    if (m_Pid.ContainsKey((uint)p.Id))
                    {
                        Dictionary<string, string> ProcessModules = Modules.Get(p);

                        foreach (KeyValuePair<string, string> Module in ProcessModules)
                        {
                            // Does the process reference winforms
                            if (Module.Value == "System.Windows.Forms" || Module.Value == "System.Windows.Forms.ni")
                            {
                                // Does the process use .NET 4.x
                                if (Module.Key.Contains(@"\v4.") || Module.Key.Contains(@"\NativeImages_v4."))
                                {
                                    ListOfProcesses.Add(p, p.ProcessName);
                                    break;
                                }
                            }

                            if (Module.Value == "APE.Bridge")
                            {
                                ListOfProcesses.Add(p, p.ProcessName);
                                break;
                            }
                        }
                    }
                }
            }

            WinformsProcessesCombobox.DataSource = new BindingSource(ListOfProcesses, null);
            WinformsProcessesCombobox.DisplayMember = "Value";
            WinformsProcessesCombobox.ValueMember = "Key";

            foreach (KeyValuePair<Process, string> cbi in WinformsProcessesCombobox.Items)
            {
                if (cbi.Key.Id == m_CurrentAttached.Key.Id)
                {
                    WinformsProcessesCombobox.SelectedItem = cbi;
                    break;
                }
            }
            WinformsProcessesCombobox.Enabled = true;
        }

        private void BuildTree()
        {
            LocateButton.Enabled = false;

            //store a temp copy of m_Identity
            ControlIdentifier temp = m_Identity;

            WindowTree.Nodes.Clear();
            PropertyListbox.Items.Clear();
            m_CurrentAttached = (KeyValuePair<Process, string>)WinformsProcessesCombobox.SelectedItem;

            if (m_CurrentAttached.Key.HasExited)
            {
                WinformsProcessesCombobox.SelectedIndex = 0;
                Populate();
            }
            else
            {
                if (m_CurrentAttached.Key.Id != Process.GetCurrentProcess().Id)
                {
                    ListOfTopLevelWindows = new Dictionary<IntPtr, string>();

                    NM.EnumWindowsProc WindowsCallback = new NM.EnumWindowsProc(EnumProc);

                    if (m_APE != null)
                    {
                        m_APE.RemoveFileMapping();
                    }

                    if (AppDomainComboBox.Enabled)
                    {
                        string NewDomain = AppDomainComboBox.SelectedItem.ToString();
                        m_APE = new APEIPC(m_CurrentAttached.Key, NewDomain);
                    }
                    else
                    {
                        m_APE = new APEIPC(m_CurrentAttached.Key);
                    }
                    m_APE.TimeOut = 0;

                    NM.EnumWindows(WindowsCallback, new IntPtr(m_CurrentAttached.Key.Id));

                    // Set up APE without calling AttachToProcess as we are already setup
                    typeof(GUI).GetField("m_APE", BindingFlags.NonPublic | BindingFlags.Static).SetValue(null, m_APE);
                    typeof(GUI).GetField("m_APESpy", BindingFlags.NonPublic | BindingFlags.Static).SetValue(null, true);
                }
            }
            //restore m_Identity
            m_Identity = temp;
        }

        private bool EnumProcToGetProcesses(IntPtr hWnd, IntPtr lParam)
        {
            uint Pid;
            NM.GetWindowThreadProcessId(hWnd, out Pid);

            if (NM.IsWindowVisible(hWnd))
            {
                NM.tagRect WindowSize;
                NM.GetClientRect(hWnd, out WindowSize);
                
                if (WindowSize.right > 0)   //If the window has 0 width then ignore it
                {
                    if (!m_Pid.ContainsKey(Pid))
                    {
                        m_Pid.Add(Pid, Pid);
                    }
                }
            }
            return true;
        }

        private bool EnumProc(IntPtr hWnd, IntPtr lParam)
        {
            uint Pid;
            NM.GetWindowThreadProcessId(hWnd, out Pid);

            if (new IntPtr(Pid) == lParam)
            {
                if (NM.IsWindowVisible(hWnd))
                {
                    NM.tagRect WindowSize;
                    NM.GetClientRect(hWnd, out WindowSize);

                    if (WindowSize.right > 0)   //If the window has 0 width then ignore it
                    {
                        TreeNode ParentNode;
                        GetIdentity(IntPtr.Zero, hWnd);
                        string APEType = GetAPEType(m_Identity);
                        ListOfTopLevelWindows.Add(hWnd, m_Identity.Name);

                        NM.tagRect WindowLoc;
                        NM.GetWindowRect(hWnd, out WindowLoc);

                        if (m_Identity.TechnologyType == "Windows Native")
                        {
                            ParentNode = WindowTree.Nodes.Add("0:" + hWnd.ToString(), APEType + " (Windows Native) [" + hWnd.ToString() + "]");
                        }
                        else
                        {
                            ParentNode = WindowTree.Nodes.Add("0:" + hWnd.ToString(), APEType + " " + m_Identity.Name + " [" + hWnd.ToString() + "]");
                        }
                        AddChildNode(hWnd, ParentNode, hWnd);
                    }
                }
            }
            return true;
        }

        private void AddChildNode(IntPtr RootParent, TreeNode ParentNode, IntPtr ParentWindow)
        {
            //string Name;
            IntPtr Child = NM.GetWindow(ParentWindow, NM.GetWindowEnum.CHILD);

            if (Child != IntPtr.Zero)
            {
                if (NM.IsWindowVisible(Child))
                {
                    GetIdentity(RootParent, Child);
                    string APEType = GetAPEType(m_Identity);

                    //Name = GetName(RootParent, Child);
                    TreeNode ChildNode = ParentNode.Nodes.Add(RootParent.ToString() + ":" + Child.ToString(), APEType + " " + m_Identity.Name + " [" + Child.ToString() + "]");
                    AddChildNode(RootParent, ChildNode, Child);
                }

                IntPtr Sibling = Child;
                do
                {
                    Sibling = NM.GetWindow(Sibling, NM.GetWindowEnum.HWNDNEXT);
                    if (Sibling != IntPtr.Zero)
                    {
                        if (NM.IsWindowVisible(Sibling))
                        {
                            GetIdentity(RootParent, Sibling);
                            string APEType = GetAPEType(m_Identity);

                            //Name = GetName(RootParent, Sibling);
                            TreeNode SiblingNode = ParentNode.Nodes.Add(RootParent.ToString() + ":" + Sibling.ToString(), APEType + " " + m_Identity.Name + " [" + Sibling.ToString() + "]");
                            AddChildNode(RootParent, SiblingNode, Sibling);
                        }
                    }
                }
                while (Sibling != IntPtr.Zero);
            }
        }

        private void WindowTree_AfterSelect(object sender, TreeViewEventArgs e)
        {
            m_ControlKey = true;
            LocateButton.Enabled = true;

            string[] SplitCharacters = { ":" };
            string[] Handles = WindowTree.SelectedNode.Name.Split(SplitCharacters, StringSplitOptions.None);

            IntPtr Parent = new IntPtr(int.Parse(Handles[0]));
            IntPtr Handle = new IntPtr(int.Parse(Handles[1]));
            if (NM.IsWindow(Handle))
            {
                PopulatePropertyListbox(Parent, Handle);
            }
            else
            {
                PropertyListbox.Items.Clear();
                if (m_CurrentAttached.Key.HasExited)
                {
                    WinformsProcessesCombobox.SelectedIndex = 0;
                    Populate();
                }
                else
                {
                    BuildTree();
                }
            }
        }

        private void LocateButton_Click(object sender, EventArgs e)
        {
            if (WindowTree.SelectedNode != null)
            {
                string[] SplitCharacters = { ":" };
                string[] Handles = WindowTree.SelectedNode.Name.Split(SplitCharacters, StringSplitOptions.None);

                IntPtr ParentHandle = new IntPtr(int.Parse(Handles[0]));
                IntPtr Handle = new IntPtr(int.Parse(Handles[1]));

                if (NM.IsWindow(Handle) && NM.IsWindowVisible(Handle))
                {
                    if (ParentHandle == IntPtr.Zero)
                    {
                        NM.BringWindowToTop(Handle);
                    }
                    else
                    {
                        NM.BringWindowToTop(ParentHandle);
                    }

                    NM.tagRect area = Display.GetWindowRectangleDIP(Handle);

                    IntPtr hWindowDC = NM.GetDC(IntPtr.Zero);
                    IntPtr hRectanglePen = NM.CreatePen(NM.PenStyle.PS_SOLID, 3, (uint)ColorTranslator.ToWin32(Color.Red));
                    IntPtr hPrevPen = NM.SelectObject(hWindowDC, hRectanglePen);
                    IntPtr hPrevBrush = NM.SelectObject(hWindowDC, NM.GetStockObject(NM.StockObjects.HOLLOW_BRUSH));

                    for (int i = 0; i < 3; i++)
                    {
                        NM.Rectangle(hWindowDC, area.left, area.top, area.right, area.bottom);
                        Thread.Sleep(300);
                        ClearHighlight(area);
                        if (i < 2)
                        {
                            Thread.Sleep(300);
                        }
                    }

                    NM.SelectObject(hWindowDC, hPrevPen);
                    NM.SelectObject(hWindowDC, hPrevBrush);
                    NM.ReleaseDC(Handle, hWindowDC);
                }
                else
                {
                    BuildTree();
                }
            }

            WindowTree.Focus();
        }

        private void ClearHighlight(NM.tagRect area)
        {
            NM.tagRect tmpArea = new NM.tagRect();

            tmpArea.left = area.left - 3;
            tmpArea.top = area.top - 3;
            tmpArea.right = area.right + 3;
            tmpArea.bottom = area.bottom + 3;

            NM.RedrawWindow(IntPtr.Zero, ref tmpArea, IntPtr.Zero, NM.RedrawWindowFlags.UpdateNow | NM.RedrawWindowFlags.Invalidate | NM.RedrawWindowFlags.AllChildren);
            Thread.Sleep(50);
        }
        
        private bool m_ControlKey = true;
        private bool m_ShiftKey = false;
        private bool m_Closing = false;
        private ControlIdentifier m_Identity = new ControlIdentifier();
        private NM.tagRect m_Area = new NM.tagRect();

        private void IdentifyButton_Click(object sender, EventArgs e)
        {
            m_CurrentAttached = (KeyValuePair<Process, string>)WinformsProcessesCombobox.SelectedItem;

            if (m_CurrentAttached.Key.HasExited)
            {
                WinformsProcessesCombobox.SelectedIndex = 0;
                Populate();
                return;
            }

            IdentifyButton.Enabled = false;
            WinformsProcessesCombobox.Enabled = false;
            buttonLeaks.Enabled = false;
            label1.Text = "Move the mouse cursor over the desired window and then press";
            label4.Text = "the control key or press shift to refresh the list of windows";

            //install hotkey hook for control key
            if (!NM.RegisterHotKey(this.Handle, 1, NM.MOD_CONTROL, NM.VK_CONTROL))
            {
                throw new Exception("Failed to register control hotkey");
            }
            if (!NM.RegisterHotKey(this.Handle, 2, NM.MOD_SHIFT, NM.VK_SHIFT))
            {
                throw new Exception("Failed to register shift hotkey");
            }
            m_ControlKey = false;

            IntPtr OldHandle = IntPtr.Zero;

            while (m_ControlKey == false)
            {
                if (m_ShiftKey)
                {
                    BuildTree();
                    m_ShiftKey = false;
                }

                Point cursorPosition = Cursor.Position;
                NM.tagPoint screenLocation;
                NM.tagPoint Location;
                IntPtr Handle;
                IntPtr ChildWindow;
                IntPtr parent = NM.GetDesktopWindow();
                screenLocation.x = cursorPosition.X;
                screenLocation.y = cursorPosition.Y;

                while (true)
                {
                    Location = screenLocation;
                    NM.ScreenToClient(parent, ref Location);
                    ChildWindow = NM.RealChildWindowFromPoint(parent, Location);

                    if (ChildWindow == IntPtr.Zero || parent == ChildWindow)
                    {
                        Handle = parent;
                        break;
                    }

                    parent = ChildWindow;
                }

                uint Pid;
                NM.GetWindowThreadProcessId(Handle, out Pid);

                if (Handle != OldHandle)
                {
                    if (OldHandle != IntPtr.Zero)
                    {
                        ClearHighlight(m_Area);
                    }

                    if (m_CurrentAttached.Key.Id == Pid)
                    {
                        PopulatePropertyListbox(NM.GetAncestor(Handle, NM.GetAncestorFlags.GetRoot), Handle);

                        Highlight(Handle);

                        OldHandle = Handle;
                    }
                    else
                    {
                        OldHandle = IntPtr.Zero;
                    }
                }

                Application.DoEvents();
                Thread.Sleep(100);
            }

            if (OldHandle != IntPtr.Zero)
            {
                ClearHighlight(m_Area);

                TreeNode[] Nodes = WindowTree.Nodes.Find(m_Identity.ParentHandle.ToString() + ":" + m_Identity.Handle.ToString(), true);
                if (Nodes.GetLength(0) == 0)
                {
                    BuildTree();
                    Nodes = WindowTree.Nodes.Find(m_Identity.ParentHandle.ToString() + ":" + m_Identity.Handle.ToString(), true);
                }
                if (Nodes.GetLength(0) > 0)
                {
                    WindowTree.SelectedNode = Nodes[0];
                    WindowTree.Focus();
                }
            }

            if (!m_Closing)
            {
                NM.UnregisterHotKey(this.Handle, 1);
                NM.UnregisterHotKey(this.Handle, 2);
                label1.Text = "";
                label4.Text = "";
                IdentifyButton.Enabled = true;
                WinformsProcessesCombobox.Enabled = true;
                buttonLeaks.Enabled = true;
            }
        }

        private void Highlight(IntPtr hWnd)
        {
            bool doHighlight = false;
            try
            {
                m_Area = Display.GetWindowRectangleDIP(hWnd);
                doHighlight = true;
            }
            catch (OutOfMemoryException)
            {
                // Ignore
            }

            if (doHighlight)
            {
                IntPtr hWindowDC = NM.GetDC(IntPtr.Zero);
                IntPtr hRectanglePen = NM.CreatePen(NM.PenStyle.PS_SOLID, 3, (uint)ColorTranslator.ToWin32(Color.Red));
                IntPtr hPrevPen = NM.SelectObject(hWindowDC, hRectanglePen);
                IntPtr hPrevBrush = NM.SelectObject(hWindowDC, NM.GetStockObject(NM.StockObjects.HOLLOW_BRUSH));

                NM.Rectangle(hWindowDC, m_Area.left, m_Area.top, m_Area.right, m_Area.bottom);

                NM.SelectObject(hWindowDC, hPrevPen);
                NM.SelectObject(hWindowDC, hPrevBrush);
                NM.ReleaseDC(Handle, hWindowDC);
            }
        }
        private void ObjectSpy_Activate(object sender, System.EventArgs e)
        {
            WindowTree.Focus();
        }

        private void TreeView_Resize(object sender, System.EventArgs e)
        {
            WindowTree.Top = PropertyListbox.Top;
            WindowTree.Height = PropertyListbox.Height;
        }

        private void GetIdentity(IntPtr Parent, IntPtr Handle)
        {
            m_Identity = new ControlIdentifier();

            try
            {
                if (Parent != Handle)
                {
                    m_Identity.ParentHandle = Parent;
                }
                m_Identity.Handle = Handle;
                m_APE.AddFirstMessageFindByProperty(m_Identity);
                m_APE.SendMessages(EventSet.APE);
                m_APE.WaitForMessages(EventSet.APE);
                m_APE.DecodeControl(1, out m_Identity);
            }
            catch (Exception ex)
            {
                Debug.WriteLine("GetIdentity by handle failed " + Parent.ToString() + " " + Handle.ToString() + ": " + ex.Message + "\r\n" + ex.StackTrace.ToArray());
            }

            try
            {
                //If the type name starts with Ax then use that as a hint that there might be a native ActiveX version of the
                //control which can be used instead
                if (m_Identity.TypeName.StartsWith("Ax"))
                {
                    ControlIdentifier identity = new ControlIdentifier();
                    if (Parent != Handle)
                    {
                        identity.ParentHandle = Parent;
                    }
                    identity.Handle = Handle;
                    identity.TechnologyType = "Windows ActiveX";
                    m_APE.AddFirstMessageFindByProperty(identity);
                    m_APE.SendMessages(EventSet.APE);
                    m_APE.WaitForMessages(EventSet.APE);
                    m_APE.DecodeControl(1, out identity);
                    m_Identity = identity;
                }
            }
            catch
            {
            }
        }

        private ControlIdentifier GetIdentity(ControlIdentifier identity)
        {
            try
            {
                m_APE.AddFirstMessageFindByProperty(identity);
                m_APE.SendMessages(EventSet.APE);
                m_APE.WaitForMessages(EventSet.APE);
                m_APE.DecodeControl(1, out identity);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message + "\r\n" + ex.StackTrace);
                identity.Handle = IntPtr.Zero;
            }

            return identity;
        }

        private void PopulatePropertyListbox(IntPtr Parent, IntPtr Handle)
        {
            GetIdentity(Parent, Handle);
            string APEType = GetAPEType(m_Identity);

            PropertyListbox.Items.Clear();
            PropertyListbox.Items.Add("Control Handle\t: " + m_Identity.Handle);
            PropertyListbox.Items.Add("Form Handle\t: " + m_Identity.ParentHandle);
            PropertyListbox.Items.Add("Name\t\t: " + m_Identity.Name);
            PropertyListbox.Items.Add("Accessibility Name\t: " + m_Identity.AccessibilityObjectName);
            if (m_Identity.Text == null)
            {
                m_Identity.Text = "";
            }
            PropertyListbox.Items.Add("Text\t\t: " + m_Identity.Text.Replace("\r", @"\r").Replace("\n", @"\n").Replace("\t", @"\t"));
            PropertyListbox.Items.Add("TypeName\t: " + m_Identity.TypeName);
            PropertyListbox.Items.Add("TypeNameSpace\t: " + m_Identity.TypeNameSpace);
            PropertyListbox.Items.Add("TechnologyType\t: " + m_Identity.TechnologyType);
            PropertyListbox.Items.Add("ModuleName\t: " + m_Identity.ModuleName);

            //Workout the index
            Guid notUsed;
            m_Identity.Index = 0;
            while (true)
            {
                m_Identity.Index++;

                ControlIdentifier identity = new ControlIdentifier();
                identity.ParentHandle = m_Identity.ParentHandle;
                identity.TechnologyType = m_Identity.TechnologyType;
                if (m_Identity.Name != "" && m_Identity.Name != null && !Guid.TryParse(m_Identity.Name, out notUsed) && m_Identity.Name != "DinamicallyLoadedControl")
                {
                    //Use Name as the index key
                    identity.Name = m_Identity.Name;
                }
                else if (m_Identity.Text != "" && m_Identity.Text != null)
                {
                    //Use Typename and Text as the index key
                    identity.Text = "^" + Regex.Escape(m_Identity.Text) + "$";
                    identity.TypeName = "^" + m_Identity.TypeName + "$";
                }
                else
                {
                    //Use type as the index key
                    identity.TypeNameSpace = "^" + m_Identity.TypeNameSpace + "$";
                    identity.TypeName = "^" + m_Identity.TypeName + "$";
                }
                identity.Index = m_Identity.Index;
                identity = GetIdentity(identity);

                if (identity.Handle == m_Identity.Handle)
                {
                    break;
                }

                if (identity.Handle == IntPtr.Zero)
                {
                    m_Identity.Index = -1;
                    break;
                }
            }

            string mainLocatorText = "";
            string indexLocatorText = "";
            string parentLocatorText = "";

            if (m_Identity.Name != "" && m_Identity.Name != null && !Guid.TryParse(m_Identity.Name, out notUsed) && m_Identity.Name != "DinamicallyLoadedControl")
            {
                PropertyListbox.Items.Add("Index (by Name)\t: " + m_Identity.Index);
                mainLocatorText = ", new Identifier(Identifiers.Name, \"" + m_Identity.Name + "\")";
            }
            else if (m_Identity.Text != "" && m_Identity.Text != null)
            {
                PropertyListbox.Items.Add("Index (by Text and\t: " + m_Identity.Index);
                PropertyListbox.Items.Add(" TypeName)");
                mainLocatorText = ", new Identifier(Identifiers.Text, @\"^" + Regex.Escape(m_Identity.Text) + "$\"), new Identifier(Identifiers.TypeName, \"" + m_Identity.TypeName + "\")";
            }
            else
            {
                PropertyListbox.Items.Add("Index (by Type)\t: " + m_Identity.Index);
                mainLocatorText = ", new Identifier(Identifiers.TypeName, \"" + m_Identity.TypeName + "\")";
            }

            if (m_Identity.Index > 1)
            {
                indexLocatorText = ", new Identifier(Identifiers.Index, " + m_Identity.Index + ")";
            }

            if (m_Identity.ParentHandle != IntPtr.Zero)
            {
                parentLocatorText = "$parentForm$, ";
            }

            PropertyListbox.Items.Add("APEType\t\t: " + APEType);

            if (APEType == "")
            {
                m_ObjectCodeLocator = "";
            }
            else
            {
                m_ObjectCodeLocator = APEType.TrimEnd(new char[] { '*' }) + " $name$ = new " + APEType.TrimEnd(new char[] { '*' }) + "(" + parentLocatorText + "$friendlyName$" + mainLocatorText + indexLocatorText + ");";
            }

            PropertyListbox.Items.Add("");

            AddCommonPropertiesToPropertyListbox();

            //Add ActiveX labels
            if (m_Identity.TechnologyType == "Windows ActiveX")
            {
                bool first = true;
                List<string> labelNamesAdded = new List<string>();
                ControlIdentifier identity;

                int index = 0;
                while (true)
                {
                    index++;
                    identity = FindByTypeNameAndIndex("Label", index);
                    if (!string.IsNullOrEmpty(identity.Name))
                    {
                        if (identity.Handle == m_Identity.Handle)
                        {
                            if (!labelNamesAdded.Contains(identity.Name))
                            {
                                int nameIndex = 0;
                                while (true)
                                {
                                    nameIndex++;
                                    identity = FindByNameAndIndex(identity.Name, nameIndex);
                                    if (identity.Handle == m_Identity.Handle)
                                    {
                                        AddRenderedActiveXToPropertyListbox(identity, "Name", ref first);
                                        if (!labelNamesAdded.Contains(identity.Name))
                                        {
                                            labelNamesAdded.Add(identity.Name);
                                        }
                                    }

                                    if (identity.Handle == IntPtr.Zero)
                                    {
                                        identity.Handle = new IntPtr(1);    //So we don't exit the loop on the check below
                                        break;
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        if (identity.Handle == m_Identity.Handle)
                        {
                            AddRenderedActiveXToPropertyListbox(identity, "Type", ref first);
                        }
                    }

                    if (identity.Handle == IntPtr.Zero)
                    {
                        break;
                    }
                }
            }

            switch (APEType.TrimEnd(new char[] { '*' }))
            {
                case "GUIToolStrip":
                case "GUIMenuStrip":
                case "GUIStatusStrip":
                    AddGUIToolStripToPropertyListbox();
                    break;
                case "GUIStatusBar":
                    AddGUIStatusBarToPropertyListbox();
                    break;
                case "GUITitleFrame":
                    AddGUITitleFrameToPropertyListbox();
                    break;
                case "GUIElementStripGrid":
                    AddGUIElementStripGridToPropertyListbox();
                    break;
                case "GUIFlexgrid":
                    AddGUIFlexGridToPropertyListbox();
                    break;
                case "GUIStretchyCombo":
                    AddGUIStretchyComboToPropertyListbox();
                    break;
                case "GUIDockContainer":
                    AddGUIDockContainerToPropertyListbox();
                    break;
                default:
                    switch (m_Identity.TypeName)
                    {
                        case "ContextMenuStrip":
                        case "ToolStripDropDownMenu":
                            AddToolStripDropDownMenuToPropertyListbox();
                            break;
                    }
                    break;
            }
        }

        private void AddRenderedActiveXToPropertyListbox(ControlIdentifier identity, string indexBy, ref bool first)
        {
            PropertyListbox.Items.Add("");
            if (first)
            {
                first = false;
                PropertyListbox.Items.Add("Rendered Controls");
            }

            PropertyListbox.Items.Add("Name\t\t: " + identity.Name);
            if (identity.Text == null)
            {
                identity.Text = "";
            }
            PropertyListbox.Items.Add("Text\t\t: " + identity.Text.Replace("\r", @"\r").Replace("\n", @"\n").Replace("\t", @"\t"));
            PropertyListbox.Items.Add("TypeName\t: " + identity.TypeName);
            PropertyListbox.Items.Add("TechnologyType\t: " + identity.TechnologyType);
            PropertyListbox.Items.Add("ModuleName\t: " + identity.ModuleName);
            PropertyListbox.Items.Add("Index (by " + indexBy + ")\t: " + identity.Index);

            string APEType = GetAPEType(identity);
            PropertyListbox.Items.Add("APEType\t\t: " + APEType);

            // Locate the form and the rendered control (label although should also work for any others)
            GUIForm form = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUILabel rendered = new GUILabel(form, "rendered", new Identifier(Identifiers.UniqueId, identity.UniqueId));

            PropertyListbox.Items.Add("Visible\t\t: " + rendered.IsVisible.ToString());
            PropertyListbox.Items.Add("Enabled\t\t: " + rendered.IsEnabled.ToString());
            PropertyListbox.Items.Add("Location\t\t: " + rendered.Left.ToString() + ", " + rendered.Top.ToString());
            PropertyListbox.Items.Add("Size\t\t: " + rendered.Width.ToString() + ", " + rendered.Height.ToString());
        }

        private ControlIdentifier FindByTypeNameAndIndex(string typeName, int index)
        {
            ControlIdentifier identity;

            identity = new ControlIdentifier();
            identity.ParentHandle = m_Identity.ParentHandle;

            identity.TechnologyType = m_Identity.TechnologyType;
            identity.TypeName = "^" + typeName + "$";

            identity.Index = index;
            identity = GetIdentity(identity);

            return identity;
        }

        private ControlIdentifier FindByNameAndIndex(string name, int index)
        {
            ControlIdentifier identity;
            
            identity = new ControlIdentifier();
            identity.ParentHandle = m_Identity.ParentHandle;

            identity.TechnologyType = m_Identity.TechnologyType;
            identity.Name = name;

            identity.Index = index;
            identity = GetIdentity(identity);

            return identity;
        }

        private void AddToolStripDropDownMenuToPropertyListbox()
        {
            int Items;
            IntPtr temp;
            temp = m_Identity.ParentHandle;
            m_Identity.ParentHandle = NM.GetAncestor(m_Identity.Handle, NM.GetAncestorFlags.GetRootOwner);

            //Get the number of items on the menustrip
            m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);
            //get the values returned
            Items = m_APE.GetValueFromMessage();

            //Loop through all items
            for (int item = 0; item < Items; item++)
            {
                m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
                m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(m_APE, item));
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Text", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "Name", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store5, "AccessibilityObject", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store5, DataStores.Store6, "Name", MemberTypes.Property);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store4);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                m_APE.SendMessages(EventSet.APE);
                m_APE.WaitForMessages(EventSet.APE);
                //get the values returned
                string itemText = m_APE.GetValueFromMessage();
                string itemName = m_APE.GetValueFromMessage();
                string itemAccessibilityObjectName = m_APE.GetValueFromMessage();

                PropertyListbox.Items.Add("Item " + item.ToString() + " Name\t: " + itemName);
                PropertyListbox.Items.Add("Item " + item.ToString() + " Text\t: " + itemText);
                PropertyListbox.Items.Add("Item " + item.ToString() + " Accessibility");
                PropertyListbox.Items.Add("    Object Name\t: " + itemAccessibilityObjectName);
            }

            m_Identity.ParentHandle = temp;
        }

        private void AddCommonPropertiesToPropertyListbox()
        {
            PropertyListbox.Items.Add("Visible\t\t: " + NM.IsWindowVisible(m_Identity.Handle).ToString());
            PropertyListbox.Items.Add("Enabled\t\t: " + NM.IsWindowEnabled(m_Identity.Handle).ToString());

            NM.tagRect windowPosition;
            NM.tagRect windowSize;
            try
            {
                NM.GetWindowRect(m_Identity.Handle, out windowPosition);
                windowSize = NM.GetClipBox(m_Identity.Handle);
                //NM.GetClientRect(m_Identity.Handle, out windowSize);

                PropertyListbox.Items.Add("Location\t\t: " + windowPosition.left.ToString() + ", " + windowPosition.top.ToString());
                PropertyListbox.Items.Add("Size\t\t: " + windowSize.right.ToString() + ", " + windowSize.bottom.ToString());
            }
            catch
            {
            }
        }

        private void AddGUIElementStripGridToPropertyListbox()
        {
            // Locate the form and the grid
            GUIForm elementStripGridForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUIElementStripGrid elementStripGrid = new GUIElementStripGrid(elementStripGridForm, "grid", new Identifier(Identifiers.Handle, m_Identity.Handle));

            PropertyListbox.Items.Add("Title Rows\t: " + elementStripGrid.TitleRows().ToString());
            PropertyListbox.Items.Add("Fixed Rows\t: " + elementStripGrid.FixedRows().ToString());
            PropertyListbox.Items.Add("Rows\t\t: " + elementStripGrid.Rows().ToString());
            PropertyListbox.Items.Add("First Visible Row\t: " + elementStripGrid.FirstVisibleRow().ToString());
            PropertyListbox.Items.Add("Fixed Columns\t: " + elementStripGrid.FixedColumns().ToString());
            PropertyListbox.Items.Add("Columns\t\t: " + elementStripGrid.Columns().ToString());
            PropertyListbox.Items.Add("First Visible Column\t: " + elementStripGrid.FirstVisibleColumn().ToString());
            PropertyListbox.Items.Add("");

            PropertyListbox.Items.Add("Columns:");
            int columns = elementStripGrid.Columns();
            int titleRows = elementStripGrid.TitleRows();
            string[] title = new string[titleRows];

            for (int column = 0; column < columns; column++)
            {
                if (!elementStripGrid.IsColumnHidden(column))
                {
                    for (int titleRow = 0; titleRow < titleRows; titleRow++)
                    {
                        string cell = elementStripGrid.GetCell(titleRow, column);
                        if (cell != "")
                        {
                            title[titleRow] = cell;
                        }
                    }

                    PropertyListbox.Items.Add(string.Join("->", title));
                }
            }
        }

        private void AddGUIFlexGridToPropertyListbox()
        {
            // Locate the form and the grid
            GUIForm flexGridForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUIFlexgrid flexGrid = new GUIFlexgrid(flexGridForm, "grid", new Identifier(Identifiers.Handle, m_Identity.Handle));
            
            PropertyListbox.Items.Add("Title Rows\t: " + flexGrid.TitleRows().ToString());
            PropertyListbox.Items.Add("Fixed Rows\t: " + flexGrid.FixedRows().ToString());
            PropertyListbox.Items.Add("Rows\t\t: " + flexGrid.Rows().ToString());
            PropertyListbox.Items.Add("First Visible Row\t: " + flexGrid.FirstVisibleRow().ToString());
            PropertyListbox.Items.Add("Fixed Columns\t: " + flexGrid.FixedColumns().ToString());
            PropertyListbox.Items.Add("Columns\t\t: " + flexGrid.Columns().ToString());
            PropertyListbox.Items.Add("First Visible Column\t: " + flexGrid.FirstVisibleColumn().ToString());
        }

        private void AddGUIStretchyComboToPropertyListbox()
        {
            // Locate the form and the stretchy combo
            GUIForm stretchyComboForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUIStretchyCombo stretchyCombo = new GUIStretchyCombo(stretchyComboForm, "stretchy combo", new Identifier(Identifiers.Handle, m_Identity.Handle));

            PropertyListbox.Items.Add("Selected Item Text\t: " + stretchyCombo.SelectedItemText());
        }

        private void AddGUIDockContainerToPropertyListbox()
        {
            // Locate the form and the dock container
            GUIForm dockContainerForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUIDockContainer dockContainer = new GUIDockContainer(dockContainerForm, "dock container", new Identifier(Identifiers.Handle, m_Identity.Handle));

            int items = dockContainer.ItemsCount();

            PropertyListbox.Items.Add("Active Item\t: " + dockContainer.ActiveItem());
            PropertyListbox.Items.Add("Item Count\t: " + items);

            for (int item = 0; item < items; item++)
            {
                PropertyListbox.Items.Add("Item " + item.ToString() + "\t\t: " + dockContainer.ItemText(item));
            }
        }

        private void AddGUIStatusBarToPropertyListbox()
        {
            // Locate the form and the status bar
            GUIForm statusBarForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUIStatusBar statusBar = new GUIStatusBar(statusBarForm, "status bar", new Identifier(Identifiers.Handle, m_Identity.Handle));

            int panels = statusBar.PanelCount();

            for (int panel = 0; panel < panels; panel++)
            {
                string panelName = statusBar.PanelName(panel);

                PropertyListbox.Items.Add("Panel " + panel.ToString() + " Name\t: " + panelName);
            }
        }

        private void AddGUITitleFrameToPropertyListbox()
        {
            // Locate the form and the title frame
            GUIForm titleFrameForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUITitleFrame titleFrame = new GUITitleFrame(titleFrameForm, "title frame", new Identifier(Identifiers.Handle, m_Identity.Handle));

            m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_toolButtons", MemberTypes.Field);
            m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int items = m_APE.GetValueFromMessage();

            for (int item = 0; item < items; item++)
            {
                m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
                m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "m_toolButtons", MemberTypes.Field);
                m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "<Indexer>", MemberTypes.Property, new Parameter(m_APE, item));
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Name", MemberTypes.Property);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                m_APE.SendMessages(EventSet.APE);
                m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemName = m_APE.GetValueFromMessage();
                string itemType = m_APE.GetValueFromMessage();

                string APESubType = "";

                switch (itemType)
                {
                    case "TitleFrameButton":
                    case "TitleFrameIconButton":
                        APESubType = "GUITitleFrameButton";
                        break;
                }

                PropertyListbox.Items.Add("Item " + item.ToString() + " Name\t: " + itemName);
                PropertyListbox.Items.Add("APESubType\t: " + APESubType);
            }
        }

        private void AddGUIToolStripToPropertyListbox()
        {
            // Locate the form and the tool strip
            GUIForm toolStripForm = new GUIForm("form", new Identifier(Identifiers.Handle, m_Identity.ParentHandle));
            GUIToolStrip toolStrip = new GUIToolStrip(toolStripForm, "toolstrip", new Identifier(Identifiers.Handle, m_Identity.Handle));
            
            m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
            m_APE.SendMessages(EventSet.APE);
            m_APE.WaitForMessages(EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int items = m_APE.GetValueFromMessage();

            for (int item = 0; item < items; item++)
            {
                m_APE.AddFirstMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
                m_APE.AddQueryMessageReflect(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(m_APE, item));
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                m_APE.AddQueryMessageReflect(DataStores.Store4, DataStores.Store5, "Name", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store6, "Text", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store2, DataStores.Store7, "AccessibilityObject", MemberTypes.Property);
                m_APE.AddQueryMessageReflect(DataStores.Store7, DataStores.Store8, "Name", MemberTypes.Property);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store5);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store6);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store8);
                m_APE.SendMessages(EventSet.APE);
                m_APE.WaitForMessages(EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemName = m_APE.GetValueFromMessage();
                string itemType = m_APE.GetValueFromMessage();
                string itemText = m_APE.GetValueFromMessage();
                string itemAccessibilityObjectName = m_APE.GetValueFromMessage();

                string APESubType = "";

                switch (itemType)
                {
                    case "ToolStripButton":
                        APESubType = "GUIToolStripButton";
                        break;
                    case "ToolStripLabel":
                    case "ToolStripStatusLabel":
                        APESubType = "GUIToolStripLabel";
                        break;
                    case "ToolStripSplitButton":
                        APESubType = "GUIToolStripSplitButton";
                        break;
                    case "ToolStripDropDownButton":
                        APESubType = "GUIToolStripDropDownButton";
                        break;
                    case "ToolStripSeparator":
                        APESubType = "GUIToolStripSeparator";
                        break;
                    case "ToolStripComboBox":
                        APESubType = "GUIToolStripComboBox";
                        break;
                    case "ToolStripTextBox":
                        APESubType = "GUIToolStripTextBox";
                        break;
                    case "ToolStripProgressBar":
                        APESubType = "GUIToolStripProgressBar";
                        break;
                    case "ToolStripMenuItem":
                        APESubType = "GUIToolStripMenu";
                        break;
                }

                PropertyListbox.Items.Add("Item " + item.ToString() + " Name\t: " + itemName);
                PropertyListbox.Items.Add("Item " + item.ToString() + " Text\t: " + itemText);
                PropertyListbox.Items.Add("Item " + item.ToString() + " Accessibility");
                PropertyListbox.Items.Add("    Object Name\t: " + itemAccessibilityObjectName);
                PropertyListbox.Items.Add("APESubType\t: " + APESubType);
            }
        }

        private string GetAPEType(ControlIdentifier Identity)
        {
            string APEType = "";

            if (Identity.ParentHandle == IntPtr.Zero)
            {
                if (m_Identity.TypeName == "ContextMenuStrip")
                {
                    // TODO add support for ContextMenu and native menus
                    APEType = "GUIContextMenu";
                }
                else
                {
                    APEType = "GUIForm";
                }
            }

            if (Identity.TechnologyType == "Windows ActiveX")
            {
                switch (Identity.TypeNameSpace)
                {
                    case null:
                    case "":
                        switch (Identity.TypeName)
                        {
                            case "CommandButton":
                                APEType = "GUIButton";
                                break;
                            case "CheckBox":
                                APEType = "GUICheckBox";
                                break;
                            case "ComboBox":
                                APEType = "GUIComboBox";
                                break;
                            case "Label":
                                APEType = "GUILabel";
                                break;
                            case "ListBox":
                                APEType = "GUIListBox";
                                break;
                            case "PictureBox":
                                APEType = "GUIPictureBox";
                                break;
                            case "OptionButton":
                                APEType = "GUIRadioButton";
                                break;
                            case "TextBox":
                                APEType = "GUITextBox";
                                break;
                        }
                        break;
                    case "VSFlex7L":
                    case "VSFlex7N":
                    case "VSFlex8L":
                    case "VSFlex8N":
                        switch (Identity.TypeName)
                        {
                            case "VSFlexGrid":
                                APEType = "GUIFlexgrid";
                                break;
                        }
                        break;
                    case "LzArrowButton":
                        switch (Identity.TypeName)
                        {
                            case "ArrowButton":
                                APEType = "GUIButton";
                                break;
                        }
                        break;
                    case "LZFLEXGRIDLib":
                        switch (Identity.TypeName)
                        {
                            case "LzFlexGridCtrl":
                                APEType = "GUIFlexgrid";
                                break;
                        }
                        break;
                    case "LzGenericWalker":
                        switch (Identity.TypeName)
                        {
                            case "GenericWalker":
                                APEType = "GUIGenericWalker";
                                break;
                        }
                        break;
                    case "LzNavBarControls":
                        switch (Identity.TypeName)
                        {
                            case "LzNavBarGridControl":
                                APEType = "GUILzNavBarGridControl";
                                break;
                        }
                        break;
                    case "DRILLDOWNLib":
                        switch (Identity.TypeName)
                        {
                            case "LZResultsGrid":
                                APEType = "GUIAxLZResultsGrid";
                                break;
                        }
                        break;
                    case "LZCASHVIEWERLib":
                        switch (Identity.TypeName)
                        {
                            case "LzCashViewerCtrl":
                                APEType = "GUIAxLzCashViewerCtrl";
                                break;
                        }
                        break;
                    case "MSComctlLib":
                        switch (Identity.TypeName)
                        {
                            case "ImageCombo":
                                APEType = "GUIComboBox";
                                break;
                            case "TabStrip":
                                APEType = "GUITabControl";
                                break;
                        }
                        break;
                    case "SftTabsLib":
                        switch (Identity.TypeName)
                        {
                            case "SftTabs":
                                APEType = "GUITabControl";
                                break;
                        }
                        break;
                    case "LzStretchyCombo":
                        switch (Identity.TypeName)
                        {
                            case "StretchyCombo":
                                APEType = "GUIStretchyCombo";
                                break;
                        }
                        break;
                    case "lzcDealingResults":
                        {
                            switch (Identity.TypeName)
                            {
                                case "lzDealingResults":
                                    APEType = "GUILzDealingResults";
                                    break;
                            }
                        }
                        break;
                }
            }
            else if (Identity.TechnologyType == "Windows Forms (WinForms)")
            {
                m_APE.AddFirstMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                m_APE.AddMessageGetRecognisedType(DataStores.Store0, DataStores.Store1);
                m_APE.AddMessageGetApeTypeFromObject(DataStores.Store0, DataStores.Store2);
                m_APE.AddMessageGetApeTypeFromType(DataStores.Store1, DataStores.Store3);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store2);
                m_APE.AddRetrieveMessageGetValue(DataStores.Store3);
                m_APE.SendMessages(EventSet.APE);
                m_APE.WaitForMessages(EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call;
                string APEDirectType = m_APE.GetValueFromMessage();
                string APEBaseType = m_APE.GetValueFromMessage();

                if (APEDirectType != "")
                {
                    APEType = APEDirectType;
                }
                else if (APEBaseType != "")
                {
                    APEType = APEBaseType + "*";     //Make it clear that its the base class which is supported rather than the extended control
                }
            }
            else if (Identity.TechnologyType == "Windows Native")
            {
                //TODO fill this in with all supported types
                if (Identity.TypeName.StartsWith("ImageCombo"))
                {
                    APEType = "GUIComboBox";
                }
                else
                {
                    switch (Identity.TypeName)
                    {
                        case "Button":
                            APEType = "GUIButton";
                            break;
                        case "Static":
                            APEType = "GUILabel";
                            break;
                    }
                }
            }

            return APEType;
        }

        protected override void WndProc(ref Message msg)
        {
            if (msg.Msg == NM.WM_HOTKEY)
            {
                if ((msg.LParam.ToInt32() & NM.MOD_CONTROL) > 0)
                {
                    m_ControlKey = true;
                }
                else if ((msg.LParam.ToInt32() & NM.MOD_SHIFT) > 0)
                {
                    m_ShiftKey = true;
                }
            }

            base.WndProc(ref msg);
        }

        private void Copy_Click(object sender, EventArgs e)
        {
            string SelectedItem = (string)PropertyListbox.SelectedItem;
            if (SelectedItem != null)
            {
                string[] Separators = { "\t: " };
                string[] Item = SelectedItem.Split(Separators, StringSplitOptions.None);
                if (Item[1] != "")
                {
                    Clipboard.SetText(Item[1]);
                }
            }
        }

        private void generateObjectCodeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (m_ObjectCodeLocator != "")
            {
                Clipboard.SetText(m_ObjectCodeLocator);
            }
        }

        private void AppDomainComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (AppDomainComboBox.Enabled)
            {
                BuildTree();
            }
        }

        private void aboutButton_Click(object sender, EventArgs e)
        {
            //TODO
        }

        private void PropertyListbox_MouseDown(object sender, MouseEventArgs args)
        {
            if (args.Button == MouseButtons.Right)
            {
                int item = PropertyListbox.IndexFromPoint(args.Location);
                if (item > -1 && PropertyListbox.SelectedIndices.Contains(item) == false)
                {
                    PropertyListbox.SelectedIndex = item;
                }
            }
        }

        private void WinformsProcessesCombobox_DropDown(object sender, EventArgs e)
        {
            RefreshProcesses();
        }

        private void WinformsProcessesCombobox_DropDownClosed(object sender, EventArgs e)
        {
            Populate();
        }

        private void Populate()
        {
            KeyValuePair<Process, string> selectedItem = (KeyValuePair<Process, string>)WinformsProcessesCombobox.SelectedItem;
            if (selectedItem.Key.Id != m_SelectedItem.Key.Id)
            {
                m_SelectedItem = selectedItem;
                BuildTree();
                m_ControlKey = true;

                if (m_CurrentAttached.Key.HasExited)
                {
                    m_CurrentAttached = new KeyValuePair<Process, string>(Process.GetCurrentProcess(), "");
                }

                if (m_CurrentAttached.Key.Id == Process.GetCurrentProcess().Id)
                {
                    AppDomainComboBox.Items.Clear();
                    AppDomainComboBox.Enabled = false;
                    IdentifyButton.Enabled = false;
                    buttonLeaks.Enabled = false;
                }
                else
                {
                    m_APE.AddFirstMessageGetAppDomains();
                    m_APE.SendMessages(EventSet.APE);
                    m_APE.WaitForMessages(EventSet.APE);
                    //Get the value(s) returned MUST be done straight after the WaitForMessages call
                    string AppDomains = m_APE.GetValueFromMessage();
                    string[] Separator = { "\t" };
                    string[] AppDomainArray = AppDomains.Split(Separator, StringSplitOptions.None);

                    AppDomainComboBox.Items.Clear();
                    for (int i = 0; i < AppDomainArray.GetLength(0); i++)
                    {
                        AppDomainComboBox.Items.Add(AppDomainArray[i]);
                    }
                    AppDomainComboBox.SelectedItem = AppDomainComboBox.Items[0];

                    AppDomainComboBox.Enabled = true;
                    IdentifyButton.Enabled = true;
                    buttonLeaks.Enabled = true;
                }
            }
        }

        private void ListBoxContextMenuStrip_Opening(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (PropertyListbox.Items.Count == 0)
            {
                copyToolStripMenuItem.Enabled = false;
                generateLocatorCodeToolStripMenuItem.Enabled = false;
            }
            else
            {
                copyToolStripMenuItem.Enabled = true;
                if (m_ObjectCodeLocator == "")
                {
                    generateLocatorCodeToolStripMenuItem.Enabled = false;
                }
                else
                {
                    generateLocatorCodeToolStripMenuItem.Enabled = true;
                }
            }
        }

        private void buttonLeaks_Click(object sender, EventArgs e)
        {
            Form leaks = new FormLeaks(m_APE);
            leaks.ShowDialog();
        }
    }
}
