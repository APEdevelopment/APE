using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.IO;
using APE.Communication;
using APE.Capture;
using System.Reflection;
using System.Threading;
using NM = APE.Native.NativeMethods;
using System.Xml;

namespace APE.Spy
{
    public partial class APESpy : Form
    {
        Dictionary<Process, String> ListOfProcesses;
        KeyValuePair<Process, string> m_CurrentAttached = new KeyValuePair<Process, string>(Process.GetCurrentProcess(), "");
        Dictionary<IntPtr, String> ListOfTopLevelWindows;
        APEIPC m_APE;

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
        }

        private void RefreshProcesses()
        {
            WinformsProcessesCombobox.Enabled = false;

            Process[] pc = Process.GetProcesses();

            ListOfProcesses = new Dictionary<Process, string>();
            ListOfProcesses.Add(Process.GetCurrentProcess(), "");

            foreach (Process p in pc)
            {
                if (p.Id != Process.GetCurrentProcess().Id)
                {
                    Dictionary<string, string> ProcessModules = Modules.Get(p);

                    foreach(KeyValuePair<string, string> Module in ProcessModules)
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

            BuildTree();
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            RefreshProcesses();   
        }

        private void BuildTree()
        {
            LocateButton.Enabled = false;

            if (WinformsProcessesCombobox.Enabled)
            {
                WindowTree.Nodes.Clear();
                PropertyListbox.Items.Clear();
                m_CurrentAttached = (KeyValuePair<Process, string>)WinformsProcessesCombobox.SelectedItem;

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

                    //m_APE.AddMessageGarbageCollect();
                    //m_APE.SendMessages(APEIPC.EventSet.APE);
                    //m_APE.WaitForMessages(APEIPC.EventSet.APE);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();

                    NM.EnumWindows(WindowsCallback, new IntPtr(m_CurrentAttached.Key.Id));
                }
            }
        }

        private void WinformsProcessesCombobox_SelectedIndexChanged(object sender, EventArgs e)
        {
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
            }
            else
            {
                m_APE.AddMessageGetAppDomains();
                m_APE.SendMessages(APEIPC.EventSet.APE);
                m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call
                string AppDomains = m_APE.GetValueFromMessage(1);
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
            }
        }

        private bool EnumProc(IntPtr hWnd, IntPtr lParam)
        {
            uint Pid;
            NM.GetWindowThreadProcessId(hWnd, out Pid);

            if (new IntPtr(Pid) == lParam)
            {
                if (NM.IsWindowVisible(hWnd))
                {
                    GetIdentity(hWnd);
                    string APEType = GetAPEType(m_Identity);
                    ListOfTopLevelWindows.Add(hWnd, m_Identity.Name);
                    TreeNode ParentNode = WindowTree.Nodes.Add("0:" + hWnd.ToString(), APEType + " " + m_Identity.Name + " [" + hWnd.ToString() + "]");
                    AddChildNode(hWnd, ParentNode, hWnd);
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
                    GetIdentity(Child);
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
                            GetIdentity(Sibling);
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

            IntPtr Handle = new IntPtr(int.Parse(Handles[1]));
            if (NM.IsWindow(Handle))
            {
                PopulatePropertyListbox(Handle);
            }
            else
            {
                PropertyListbox.Items.Clear();
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
        private bool m_Closing = false;
        private ControlIdentifier m_Identity = new ControlIdentifier();
        private NM.tagRect m_Area = new NM.tagRect();

        private void IdentifyButton_Click(object sender, EventArgs e)
        {
            IdentifyButton.Enabled = false;
            label1.Text = "Move the mouse cursor over the desired window and then press the control key";

            //install hotkey hook for control key
            if (!NM.RegisterHotKey(this.Handle, 1, NM.MOD_CONTROL, NM.VK_CONTROL))
            {
                throw new Exception("Failed to register hotkey");
            }
            m_ControlKey = false;

            IntPtr OldHandle = IntPtr.Zero;

            while (m_ControlKey == false)
            {
                Point cursorPosition = Cursor.Position;
                NM.tagPoint Location;
                Location.x = cursorPosition.X;
                Location.y = cursorPosition.Y;

                IntPtr Handle = NM.WindowFromPoint(Location);
                NM.ScreenToClient(Handle, ref Location);
                IntPtr ChildWindow = NM.ChildWindowFromPointEx(Handle, Location, NM.ChildWindowFromPointExFlags.SKIPINVISIBLE);

                if (ChildWindow != IntPtr.Zero)
                {
                    Handle = ChildWindow;
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
                        PopulatePropertyListbox(Handle);

                        m_Area = Display.GetWindowRectangleDIP(m_Identity.Handle);

                        IntPtr hWindowDC = NM.GetDC(IntPtr.Zero);
                        IntPtr hRectanglePen = NM.CreatePen(NM.PenStyle.PS_SOLID, 3, (uint)ColorTranslator.ToWin32(Color.Red));
                        IntPtr hPrevPen = NM.SelectObject(hWindowDC, hRectanglePen);
                        IntPtr hPrevBrush = NM.SelectObject(hWindowDC, NM.GetStockObject(NM.StockObjects.HOLLOW_BRUSH));

                        NM.Rectangle(hWindowDC, m_Area.left, m_Area.top, m_Area.right, m_Area.bottom);

                        NM.SelectObject(hWindowDC, hPrevPen);
                        NM.SelectObject(hWindowDC, hPrevBrush);
                        NM.ReleaseDC(Handle, hWindowDC);

                        OldHandle = Handle;
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
                label1.Text = "";
                IdentifyButton.Enabled = true;
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

        private void GetIdentity(IntPtr Handle)
        {
            m_Identity = new ControlIdentifier();

            try
            {
                m_Identity.Handle = Handle;
                m_APE.AddMessageFindByProperty(m_Identity);
                m_APE.SendMessages(APEIPC.EventSet.APE);
                m_APE.WaitForMessages(APEIPC.EventSet.APE);
                m_APE.DecodeControl(1, out m_Identity);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.Message + "\r\n" + ex.StackTrace.ToArray());

                m_Identity.Handle = Handle;
                m_Identity.ParentHandle = NM.GetAncestor(Handle, NM.GetAncestorFlags.GetRoot);
                if (m_Identity.Handle == m_Identity.ParentHandle)
                {
                    m_Identity.ParentHandle = IntPtr.Zero;
                }
                //TODO fix this properly
                m_Identity.TechnologyType = "Windows Native";
                m_Identity.Index = 1;
            }
        }

        private void PopulatePropertyListbox(IntPtr Handle)
        {
            GetIdentity(Handle);
            string APEType = GetAPEType(m_Identity);

            PropertyListbox.Items.Clear();
            PropertyListbox.Items.Add("Control Handle\t: " + m_Identity.Handle);
            PropertyListbox.Items.Add("Form Handle\t: " + m_Identity.ParentHandle);
            PropertyListbox.Items.Add("Name\t\t: " + m_Identity.Name);
            PropertyListbox.Items.Add("Text\t\t: " + m_Identity.Text);
            PropertyListbox.Items.Add("TypeName\t: " + m_Identity.TypeName);
            PropertyListbox.Items.Add("TypeNameSpace\t: " + m_Identity.TypeNameSpace);
            PropertyListbox.Items.Add("TechnologyType\t: " + m_Identity.TechnologyType);
            PropertyListbox.Items.Add("ModuleName\t: " + m_Identity.ModuleName);
            PropertyListbox.Items.Add("Index\t\t: " + m_Identity.Index);
            PropertyListbox.Items.Add("APEType\t\t: " + APEType);

            PropertyListbox.Items.Add("");

            switch(APEType)
            {
                case "GUIToolStrip":
                case "GUIMenuStrip":
                case "GUIStatusStrip":
                    AddGUIToolStripToPropertyListbox();
                    break;
                case "GUIStatusBar":
                    AddGUIStatusBarToPropertyListbox();
                    break;
                case "GUIElementStripGrid":
                    AddGUIElementStripGridToPropertyListbox();
                    break;
            }
        }

        private void AddGUIElementStripGridToPropertyListbox()
        {
            m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "GetColumnInfoXML", MemberTypes.Method);
            m_APE.AddMessageGetValue(DataStores.Store1);
            m_APE.SendMessages(APEIPC.EventSet.APE);
            m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            string ColumnInfoXML = m_APE.GetValueFromMessage(1);

            XmlDocument columnDocument = new XmlDocument();
            columnDocument.LoadXml(ColumnInfoXML);

            int maxLevels = -1;
            int currentLevel = -1;
            foreach (XmlNode node in columnDocument.SelectSingleNode("Columns").ChildNodes)
            {
                if (int.TryParse(node.Attributes.GetNamedItem("ColumnHeaderLevel").Value, out currentLevel))
                {
                    if (currentLevel > maxLevels)
                    {
                        maxLevels = currentLevel;
                    }
                }
            }

            PropertyListbox.Items.Add("Title Rows\t: " + (maxLevels + 1).ToString());

            m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NumRows", MemberTypes.Property);
            m_APE.AddMessageGetValue(DataStores.Store1);
            m_APE.SendMessages(APEIPC.EventSet.APE);
            m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int Rows = m_APE.GetValueFromMessage(1);

            PropertyListbox.Items.Add("Rows\t\t: " + (Rows + maxLevels + 1).ToString());

            m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "NumFrozenRows", MemberTypes.Method);
            m_APE.AddMessageGetValue(DataStores.Store1);
            m_APE.SendMessages(APEIPC.EventSet.APE);
            m_APE.WaitForMessages(APEIPC.EventSet.APE);
            //Get the value(s) returned MUST be done straight after the WaitForMessages call
            int FixedRows = m_APE.GetValueFromMessage(1);

            PropertyListbox.Items.Add("Fixed Rows\t: " + (FixedRows + maxLevels + 1).ToString());
        }

        private void AddGUIStatusBarToPropertyListbox()
        {
            m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
            m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            m_APE.AddMessageGetValue(DataStores.Store2);
            m_APE.SendMessages(APEIPC.EventSet.APE);
            m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int items = m_APE.GetValueFromMessage(1);

            for (int item = 0; item < items; item++)
            {
                m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
                m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Panels", MemberTypes.Property);
                m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(m_APE, item));
                m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store5, "Name", MemberTypes.Property);
                m_APE.AddMessageGetValue(DataStores.Store3);
                m_APE.AddMessageGetValue(DataStores.Store5);
                m_APE.SendMessages(APEIPC.EventSet.APE);
                m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemName = m_APE.GetValueFromMessage(1);
                string itemType = m_APE.GetValueFromMessage(2);

                //string APESubType = "GUIStatusBarPanel";

                PropertyListbox.Items.Add("Item Name\t: " + itemName);
                //PropertyListbox.Items.Add("APESubType\t: " + APESubType);
                PropertyListbox.Items.Add("");
            }
        }

        private void AddGUIToolStripToPropertyListbox()
        {
            m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
            m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
            m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Count", MemberTypes.Property);
            m_APE.AddMessageGetValue(DataStores.Store2);
            m_APE.SendMessages(APEIPC.EventSet.APE);
            m_APE.WaitForMessages(APEIPC.EventSet.APE);
            // Get the value(s) returned MUST be done straight after the WaitForMessages call
            int items = m_APE.GetValueFromMessage(1);

            for (int item = 0; item < items; item++)
            {
                m_APE.AddMessageFindByHandle(DataStores.Store0, m_Identity.ParentHandle, m_Identity.Handle);
                m_APE.AddMessageQueryMember(DataStores.Store0, DataStores.Store1, "Items", MemberTypes.Property);
                m_APE.AddMessageQueryMember(DataStores.Store1, DataStores.Store2, "Item", MemberTypes.Property, new Parameter(m_APE, item));
                m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store3, "Name", MemberTypes.Property);
                m_APE.AddMessageQueryMember(DataStores.Store2, DataStores.Store4, "GetType", MemberTypes.Method);
                m_APE.AddMessageQueryMember(DataStores.Store4, DataStores.Store5, "Name", MemberTypes.Property);
                m_APE.AddMessageGetValue(DataStores.Store3);
                m_APE.AddMessageGetValue(DataStores.Store5);
                m_APE.SendMessages(APEIPC.EventSet.APE);
                m_APE.WaitForMessages(APEIPC.EventSet.APE);
                // Get the value(s) returned MUST be done straight after the WaitForMessages call
                string itemName = m_APE.GetValueFromMessage(1);
                string itemType = m_APE.GetValueFromMessage(2);

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

                PropertyListbox.Items.Add("Item Name\t: " + itemName);
                PropertyListbox.Items.Add("APESubType\t: " + APESubType);
                PropertyListbox.Items.Add("");
            }
        }

        private string GetAPEType(ControlIdentifier Identity)
        {           
            string APEType = "";

            // TODO will this always be the case?
            if (Identity.ParentHandle == IntPtr.Zero)
            {
                APEType = "GUIForm";
            }

            if (Identity.TechnologyType == "Windows Forms (WinForms)")
            {
                m_APE.AddMessageFindByHandle(DataStores.Store0, Identity.ParentHandle, Identity.Handle);
                m_APE.AddMessageGetRecognisedType(DataStores.Store0, DataStores.Store1);
                m_APE.AddMessageGetApeTypeFromObject(DataStores.Store0, DataStores.Store2);
                m_APE.AddMessageGetApeTypeFromType(DataStores.Store1, DataStores.Store3);
                m_APE.AddMessageGetValue(DataStores.Store2);
                m_APE.AddMessageGetValue(DataStores.Store3);
                m_APE.SendMessages(APEIPC.EventSet.APE);
                m_APE.WaitForMessages(APEIPC.EventSet.APE);
                //Get the value(s) returned MUST be done straight after the WaitForMessages call;
                string APEDirectType = m_APE.GetValueFromMessage(1);
                string APEBaseType = m_APE.GetValueFromMessage(2);

                if (APEDirectType != "")
                {
                    APEType = APEDirectType;
                }
                else if (APEBaseType != "")
                {
                    APEType = APEBaseType + "*";     //Make it clear that its the base class which is supported rather than the extended control
                }
            }

            return APEType;
        }

        protected override void WndProc(ref Message msg)
        {
            if (msg.Msg == NM.WM_HOTKEY)
            {
                m_ControlKey = true;
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
                Clipboard.SetText(Item[1]);
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
    }
}
