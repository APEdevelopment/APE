using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
//
using System.Diagnostics;
using APE.Language;
using APE.Capture;

using System.Threading;
using System.Drawing;
using System.IO;
//
using System.Net.Mail;
//Debug.Listeners[0].WriteLine("\t Time: " + timer.ElapsedMilliseconds.ToString());

namespace APE.Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            GUI.Log("Starting", APE.LogItemTypeEnum.ApeContext);
        }

        private void button1_Click(object sender, EventArgs e)  //Sentinel
        {
            button1.Enabled = false;

            Process p;

            //close any running sentinels
            while (Process.GetProcessesByName("LzSentinel").Length > 0)
            {
                // kill the process
                try
                {
                    p = Process.GetProcessesByName("LzSentinel")[0];
                    p.Kill();
                    p.WaitForExit();
                }
                catch
                {
                }
            }
            
            //start sentinel
            GUI.Log("Launch Sentinel", APE.LogItemTypeEnum.Action);
            ProcessStartInfo SentinelStartup = new ProcessStartInfo();
            SentinelStartup.WorkingDirectory = @"C:\Program Files (x86)\LatentZero\LZ Sentinel\Client\";
            SentinelStartup.FileName = @"LzSentinel.exe";
            SentinelStartup.Arguments = @".\lzSentinel.ini";
            Process.Start(SentinelStartup);

            //find the process
            p = Process.GetProcessesByName("LzSentinel")[0];

            //attach
            GUI.AttachToProcess(p);
            GUI.SetTimeOuts(4000);
            
            //find the login controls and login
            GUIForm Login = new GUIForm("login form", new Identifier(Identifiers.Name, "SimpleContainer"));
            GUITextBox UserId = new GUITextBox(Login, "user id textbox", new Identifier(Identifiers.Name, "txtUserId"));
            GUITextBox Password = new GUITextBox(Login, "password textbox", new Identifier(Identifiers.Name, "txtPassword"));

            GUIButton OK = new GUIButton(Login, "ok button", new Identifier(Identifiers.Text, "OK"));
            GUIButton Cancel = new GUIButton(Login, "cancel button", new Identifier(Identifiers.Text, "Cancel"));

            UserId.SetText("autosen");
            Password.SetText("quality");

            OK.MouseSingleClick(MouseButton.Left);

            //find main sentinel window and the status bar
            GUI.Log("Wait for Sentinel to load", APE.LogItemTypeEnum.Information);
            GUIForm Sentinel = new GUIForm("sentinel form", new Identifier(Identifiers.Name, "CapstoneContainer"));
            GUIStatusBar SentinelStatusBar = new GUIStatusBar(Sentinel, "sentiinel form status strip", new Identifier(Identifiers.Name, "statusBar"));

            //wait for the first panel in the status bar to be empty
            while (SentinelStatusBar.PanelText(0) != "")    //sbStatus
            {
                Thread.Sleep(100);
            }


            ////select the File\New Layout... menu item
            GUIMenuStrip MenuStrip = new GUIMenuStrip(Sentinel, "sentinel form menu strip", new Identifier(Identifiers.TypeName, "MenuStrip"));            
            //MenuStrip.Select(@"&File\New Layout...");
            
            ////Find the inputbox and close it
            //GUIForm NewLayout = new GUIForm(new Identifier(Identifiers.Name, "LzInputBox"));
            //NewLayout.Close();

            //// Test garbage collection (does both AUT and APE)
            //GUI.GarbageCollect();

            //GUIFlexgrid vfgStats = new GUIFlexgrid(Sentinel, new Identifier(Identifiers.Name, "vfgStats"));

            //string gridcontents = vfgStats.GetCellRangeClip(0, 0, vfgStats.Rows() - 1, vfgStats.Columns() - 1);
            
            //Debug.Listeners[0].WriteLine("\t Grid:\n" + gridcontents);

            ////MessageBox.Show("unhide thr row");
            //vfgStats.Select(3, 0, MouseButton.Left, CellClickLocation.ExpandCollapseIconOfCell);
            //MessageBox.Show(vfgStats.FindRow("Post-trade Incidents -> New Missing Data").ToString());

            
            //GUIFlexgrid vfgStats = new GUIFlexgrid(Sentinel, new Identifier(Identifiers.ModuleName, "DocumentContainer"));
            GUIDocumentContainer DocumentContainer = new GUIDocumentContainer(Sentinel, "panel tab", new Identifier(Identifiers.Name, ""), new Identifier(Identifiers.TypeNameSpace, "TD.SandDock"), new Identifier(Identifiers.TypeName, "DocumentContainer"));

            bool Exists = DocumentContainer.ItemExists("Sentinel Today");
            int Count = DocumentContainer.ItemCount();
            DocumentContainer.ItemSelect("Pre-trade");
            DocumentContainer.ItemSelect("Rule Maintenance");
            DocumentContainer.ItemSelect("Sentinel Today");
            DocumentContainer.ItemRemove("Pre-trade");

            //Add a panel
            int InitialItems = DocumentContainer.ItemCount();
            int CurrentItems = 0;
            MenuStrip.Select(@"&Panels\Pre-trade");

            do
            {
                Thread.Sleep(50);

                CurrentItems = DocumentContainer.ItemCount();
            } while (CurrentItems != InitialItems + 1);

            DocumentContainer.ItemSelect("Sentinel Today");
    
    

            //classname = ffdhgfd.fshdj.10.fsdhf (useless)
            //modulename = (whatever.dll)
            //typename = toolstripcontainer
            //TypeNameSpace = system.wtahtever.ook

            //MessageBox.Show(vfgStats.IsInEditMode().ToString());

            //close sentinel
            Sentinel.Close();



            //GUIForm myForm = new GUIForm(new Identifier(Identifiers.Name, "CapstoneContainer"));
            //MessageBox.Show(
            //    "Child = " + UserId.Handle.ToString() + Environment.NewLine +
            //    "Name = " + UserId.Name + Environment.NewLine +
            //    "TechnologyType = " + UserId.TechnologyType + Environment.NewLine +
            //    "TypeNameSpace = " + UserId.TypeNameSpace + Environment.NewLine +
            //    "TypeName = " + UserId.TypeName + Environment.NewLine +
            //    "ModuleName = " + UserId.ModuleName + Environment.NewLine +
            //    "AssemblyName = " + UserId.AssemblyName + Environment.NewLine +
            //    "Text = " + UserId.Text
            //);

            button1.Enabled = true;
        }

        private void SwitchLayout(GUIForm IMS, GUIFlexgrid layoutGrid, string group, string layout)
        {
            Stopwatch timer;

            GUIStatusBar statusBar = new GUIStatusBar(IMS, "IMS status bar", new Identifier(Identifiers.Name, "statusBar"));



            layoutGrid.Select(group + " -> " + layout, layoutGrid.FirstVisibleColumn(), MouseButton.Left, CellClickLocation.CentreOfCell);
            
            timer = Stopwatch.StartNew();
            do
            {
                if (timer.ElapsedMilliseconds > GUI.GetTimeOuts())
                {
                    throw new Exception("Failed to switch layouts");
                }

                if (statusBar.PanelText("sbStatus") == "")
                {
                    break;
                }

                Thread.Sleep(50);
            } while (true);
            timer.Stop();
            
            GUIElementStripGrid psGrid = new GUIElementStripGrid(IMS, "m_elementStripGrid", new Identifier(Identifiers.Name, "m_elementStripGrid"));

            switch(layout)
            {
                case "Order Viewer Only":
                case "test":
                    break;
                default:
                    timer = Stopwatch.StartNew();
                    do
                    {
                        if (timer.ElapsedMilliseconds > GUI.GetTimeOuts())
                        {
                            throw new Exception("Failed to switch layouts");
                        }
                        //psGrid.IsVisible
                        int rows = psGrid.Rows();
                        int fixedRow = psGrid.FixedRows();

                        if (rows > fixedRow)
                        {
                            break;
                        }

                        Thread.Sleep(50);
                    }
                    while (true);
                    timer.Stop();
                    break;
            }

            //whole loop maybe takes 100ms if product is idle
            for (int i = 0; i < 50; i++)
            {
                GUI.WaitForInputIdle(psGrid);
            }
                
        }

        private void button2_Click(object sender, EventArgs e)  //IMS
        {
            button2.Enabled = false;

            Stopwatch stepTimer;
            


            Process p = Process.GetProcessesByName("LzCapstone")[0];
            GUI.AttachToProcess(p);


            GUIForm IMS = new GUIForm("IMS form", new Identifier(Identifiers.Name, "CapstoneContainer"));
            GUIFlexgrid layout = new GUIFlexgrid(IMS, "layout treeview", new Identifier(Identifiers.Name, "treeView"));

            //https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx
            //Debug.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("Orders [Count: 3][Link:]", "Orders [[]Count: [0-9]+[]][[]Link:[]]").ToString());


            GUIDockableWindow OrderViewerDockableWindow = new GUIDockableWindow(IMS, "Order viewer dockable window", new Identifier(Identifiers.Text, "Orders [[]Count: [0-9]+[]][[]Link:[]]"), new Identifier(Identifiers.TypeName, "DockableWindow"));
            GUIElementStripGrid OrderViewerGrid = new GUIElementStripGrid(IMS, "Order viewer grid", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.ChildOf, OrderViewerDockableWindow));


            GUIDockableWindow ExecutionViewerDockableWindow = new GUIDockableWindow(IMS, "Execution viewer dockable window", new Identifier(Identifiers.Text, "Executions [[]Count: [0-9]+[]][[]Link:[]]"), new Identifier(Identifiers.TypeName, "DockableWindow"));
            GUIElementStripGrid ExecutionGrid = new GUIElementStripGrid(IMS, "Execution viewer grid", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.ChildOf, ExecutionViewerDockableWindow));





            GUIForm ASA = new GUIForm("ASA form", new Identifier(Identifiers.Name, "SimpleContainer"));

            GUIElementStripGrid entry = new GUIElementStripGrid(ASA, "entry strip", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.Index, 2));

            Debug.WriteLine(entry.FindRow("Order"));
            Debug.WriteLine(entry.FindColumn("Instrument"));

            //entry.SetCellValue(1, 0, "Execution", "Execution", null);
            entry.SetCellValue(1, 4, "AEGON 5 3/4 12/15/20", "AEGON 5 3/4 12/15/20", null);

            //Stopwatch foo = Stopwatch.StartNew();

            
            //foo.Stop();
            //Debug.WriteLine((foo.ElapsedMilliseconds / entry.Columns()).ToString());


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "test");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "test");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));

            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Standard");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Standard");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Maximum");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Maximum");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Portfolio Studio x 2");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Portfolio Studio x 2");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "PS and Matrix");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "PS and Matrix");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "PS and Matrix Max");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));


            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "PS and Matrix Max");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));
            //stepTimer = Stopwatch.StartNew();
            //SwitchLayout(IMS, layout, "Automated Screens", "Order Viewer Only");
            //stepTimer.Stop();
            //Debug.WriteLine(String.Format("{0:N2}", (Convert.ToDouble(stepTimer.ElapsedMilliseconds) / 1000)));

            //Process p = Process.GetProcessesByName("toolstrip")[0];
            //GUI.AttachToProcess(p);

            //GUIForm test = new GUIForm("splash screen form", new Identifier(Identifiers.Name, "Form1"));

            //GUIToolStrip tt = new GUIToolStrip(test, "toolstrip", new Identifier(Identifiers.Name, "toolStrip1"));

            //GUIToolStripButton daButton = tt.GetButton("da button", new Identifier(Identifiers.Name, "toolStripButton1"));

            //GUIToolStripComboBox daCombo = tt.GetComboBox("da combo", new Identifier(Identifiers.Name, "toolStripComboBox1"));

            //GUIToolStripDropDownButton daDropDownButton = tt.GetDropDownButton("da drop down button", new Identifier(Identifiers.Name, "toolStripDropDownButton1"));


            //GUIMenuStrip ms = new GUIMenuStrip(test, "toolstrip", new Identifier(Identifiers.Name, "menuStrip1"));
            //GUIToolStripMenu daMenu = ms.GetMenu("da menu", new Identifier(Identifiers.Name, "toolStripMenuItem1"));

            //daMenu.Select("2");

            //daDropDownButton.Select(@"1\C\y");

            //toolStripComboBox1

            //daButton.MouseSingleClick(MouseButton.Left);


            //Process p = Process.GetProcessesByName("LzSentinel")[0];
            //GUI.AttachToProcess(p);


            //GUIForm sentinel = new GUIForm("splash screen form", new Identifier(Identifiers.Name, "SplashScreenForm"));

            //GUIPictureBox splash = new GUIPictureBox(sentinel, "splash picturebox", new Identifier(Identifiers.Name, "picSplash"));
            //splash.SaveBackground(@"C:\Tools\splash.bmp");

            //GUIProgressBar prog = new GUIProgressBar(sentinel, "progress bar", new Identifier(Identifiers.Name, "progressBar"));
            //MessageBox.Show(prog.Value.ToString() + " " + prog.Maximum.ToString());
            //Process p = Process.GetProcessesByName("LzCapstone")[0];

            //GUI.AttachToProcess(p);

            ////Thread.Sleep(2000);

            //GUIForm IMS = new GUIForm("single quote form", new Identifier(Identifiers.Name, "frmSingleQuote"));

            //GUIFlexgrid fgGrid = new GUIFlexgrid(IMS, "add quote grid", new Identifier(Identifiers.Name, "fgGrid"));

            //Stopwatch xxx = Stopwatch.StartNew();
            //string text = "";










            //xxx = Stopwatch.StartNew();
            //fgGrid.Select(1, "Counterparty", MouseButton.Left, CellClickLocation.CentreOfCell);
            //GUI.WaitForInputIdle(fgGrid);
            //SendKeys.SendWait("HSBC{ENTER}");
            //GUI.WaitForInputIdle(fgGrid);
            //while (true)
            //{
            //    text = fgGrid.GetCellValue(1, "Counterparty", GUIFlexgrid.CellProperty.TextDisplay);
            //    if (text == "HSBC")
            //    {
            //        break;
            //    }
            //}
            //Debug.WriteLine("Counterparty: " + xxx.ElapsedMilliseconds.ToString());

            //xxx = Stopwatch.StartNew();
            //fgGrid.SetCellValue(1, "Source", "Counterparty");
            //Debug.WriteLine("Source: " + xxx.ElapsedMilliseconds.ToString());

            //xxx = Stopwatch.StartNew();
            //fgGrid.SetCellValue(1, "Bid Price", "106.80");
            //Debug.WriteLine("Bid Price: " + xxx.ElapsedMilliseconds.ToString());

            //xxx = Stopwatch.StartNew();
            //fgGrid.SetCellValue(1, "Bid Size", "10000", "10,000.00");
            //Debug.WriteLine("Bid Size: " + xxx.ElapsedMilliseconds.ToString());

            //xxx = Stopwatch.StartNew();
            //fgGrid.SetCellValue(1, "Offer Price", "106.99");
            //Debug.WriteLine("Offer Price: " + xxx.ElapsedMilliseconds.ToString());


            //xxx = Stopwatch.StartNew();
            //fgGrid.SetCellValue(1, "Offer Size", "200000", "200,000.00");
            //Debug.WriteLine("Offer Size: " + xxx.ElapsedMilliseconds.ToString());

            //fgGrid.SetCellValue(1, "Counterparty", "ook");

            button2.Enabled = true;
        }

        private static void WriteToFile(string Line)
        {
            TextWriter log = File.AppendText(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\TestApplication.log");
            log.WriteLine(DateTime.Now.ToString() + "\t" + Line);
            log.Close();
        }

        private void Log(string textToLog, LogItemTypeEnum type)
        {
            WriteToFile(textToLog);
        }

        private void button3_Click(object sender, EventArgs e)  //Test Application
        {
            button3.Enabled = false;

            SmtpClient mySMTP = new SmtpClient("smtp.fidessa.com", 25);
            int loops = int.Parse(textBox1.Text);
            Stopwatch testTimer = Stopwatch.StartNew();

            if (loops > 100)
            {
                Thread.Sleep(5000);
            }

            try
            {
                textBox1.Refresh();

                //Setup our custom logger
                File.Delete(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory) + @"\TestApplication.log");
                GUI.Logger = new GUI.LoggerDelegate(Log);

                for (int x = 0; x < loops; x++)
                {
                    textBox1.Text = (loops - x).ToString();
                    Process p;

                    // close any already running instances
                    while (Process.GetProcessesByName("TestApplication").Length > 0)
                    {
                        // kill the process
                        try
                        {
                            p = Process.GetProcessesByName("TestApplication")[0];
                            p.Kill();
                            p.WaitForExit();
                        }
                        catch
                        {
                        }
                    }

                    //start
                    GUI.Log("Launch TestApplication", APE.LogItemTypeEnum.Action);

                    ProcessStartInfo AppStartup = new ProcessStartInfo();
                    //AppStartup.WorkingDirectory = @"C:\Tools\TestApplication\TestApplication\bin\Debug\";
                    AppStartup.WorkingDirectory = @".\..\..\..\..\TestApplication\TestApplication\bin\Debug\";
                    AppStartup.FileName = @"TestApplication.exe";
                    Process.Start(AppStartup);

                    //find the process
                    p = Process.GetProcessesByName("TestApplication")[0];
                    //p = Process.GetProcessesByName("TestApplication.vshost")[0];

                    //attach
                    GUI.AttachToProcess(p);

                    GUI.SetTimeOuts(20000);

                    //find the form
                    GUIForm TestApplication = new GUIForm("test application main form", new Identifier(Identifiers.Name, "MainForm"));
                    GUITextBox tbxContext = new GUITextBox(TestApplication, "textbox", new Identifier(Identifiers.Name, "tbxContext"));
                    GUIButton btnStatusBar = new GUIButton(TestApplication, "increment statusbar button", new Identifier(Identifiers.Text, "Increment StatusBar"));

                    tbxContext.SetText("autosen");
                    tbxContext.SetText("whoop");
                    tbxContext.SetText("moo");

                    GUIMenuStrip MenuStrip = new GUIMenuStrip(TestApplication, "menu strip", new Identifier(Identifiers.TypeName, "MenuStrip"));
                    MenuStrip.Select(@"Second\OpensPopup");

                    GUIForm CustomMessageBox = new GUIForm("messagebox form", new Identifier(Identifiers.Name, "frmCustomMessageBox"));
                    CustomMessageBox.Maximise();
                    CustomMessageBox.Restore();

                    CustomMessageBox.Move(500, 500);

                    CustomMessageBox.Minimise();
                    CustomMessageBox.Restore();

                    CustomMessageBox.Close();

                    GUIStatusStrip ssStatusBar = new GUIStatusStrip(TestApplication, "status strip", new Identifier(Identifiers.Name, "StatusStrip"));

                    string PanelText;

                    GUIToolStripLabel statusBarFirstPanel = ssStatusBar.GetLabel("first panel", new Identifier(Identifiers.Name, "toolStripStatusLabel1"));

                    GUIToolStripLabel statusBarFirstPanelIdentTest = ssStatusBar.GetLabel("first panel for ident test", new Identifier(Identifiers.Name, "toolStripStatusLabel1"), new Identifier(Identifiers.TechnologyType, "Windows Forms (WinForms)"), new Identifier(Identifiers.TypeNameSpace, "System.Windows.Forms"), new Identifier(Identifiers.TypeName, "ToolStripStatusLabel"), new Identifier(Identifiers.ModuleName, "System.Windows.Forms.dll"), new Identifier(Identifiers.AssemblyName, "System.Windows.Forms"), new Identifier(Identifiers.Text, "toolStripStatusLabel1"));

                    PanelText = statusBarFirstPanel.Text;

                    //int PanelIndex = ssStatusBar.PanelIndex("toolStripStatusLabel1");
                    //PanelText = ssStatusBar.PanelText(PanelIndex);
                    //PanelText = ssStatusBar.PanelText("toolStripStatusLabel5");

                    btnStatusBar.MouseSingleClick(MouseButton.Left);
                    statusBarFirstPanel.PollForText("6");
                    statusBarFirstPanel.PollForText("");
                    Stopwatch doo = Stopwatch.StartNew();
                    statusBarFirstPanel.PollForText("");
                    Debug.WriteLine(doo.ElapsedMilliseconds.ToString());
//                    while (ssStatusBar.PanelText(0) != "")
//                    {
//                        Thread.Sleep(100);
//                    }

                    //bring up the context menu of the textbox and select the menu item
                    //TODO make this a single command??
                    tbxContext.MouseSingleClick(MouseButton.Right);
                    tbxContext.ContextMenuSelect(@"ConItem2\ConSubItem2");
                    
                    CustomMessageBox = new GUIForm("messagebox form", new Identifier(Identifiers.Name, "frmCustomMessageBox"));
                    CustomMessageBox.Close();

                    GUILabel EditBoxLabel = new GUILabel(TestApplication, "label1 label", new Identifier(Identifiers.Name, "lblTextBox"));
                    string labelText = EditBoxLabel.Text;

                    GUICheckBox MyCheckBox = new GUICheckBox(TestApplication, "checkbox1 checkbox", new Identifier(Identifiers.Name, "checkBox1"));
                    MyCheckBox.MouseSingleClick(MouseButton.Left);

                    GUIRadioButton MyRadioButton = new GUIRadioButton(TestApplication, "radioButton1 radio button", new Identifier(Identifiers.Name, "radioButton1"));
                    MyRadioButton.MouseSingleClick(MouseButton.Left);

                    GUIComboBox MyComboBox1 = new GUIComboBox(TestApplication, "top left combobox", new Identifier(Identifiers.Name, "comboBox1"));
                    MyComboBox1.SetText("foo");
                    MyComboBox1.SetText("bar");
                    MyComboBox1.ItemSelect("Test31");

                    GUIComboBox MyComboBox2 = new GUIComboBox(TestApplication, "bottom left combobox", new Identifier(Identifiers.Name, "comboBox2"));
                    MyComboBox2.ItemSelect("Test40");

                    GUIComboBox MyComboBox3 = new GUIComboBox(TestApplication, "top right combobox", new Identifier(Identifiers.Name, "comboBox3"));
                    MyComboBox3.ItemSelect("Test50");
                    
                    MyComboBox3 = new GUIComboBox(TestApplication, "top right combobox", new Identifier(Identifiers.Name, "comboBox3"));
                    MyComboBox3.ItemSelect("Test25");

                    GUIListBox MyListBox = new GUIListBox(TestApplication, "bottom right listbox", new Identifier(Identifiers.Name, "listBox1"));
                    MyListBox.ItemSelect("Test25");

                    GUITreeView MyTreeView1 = new GUITreeView(TestApplication, "top left treeview", new Identifier(Identifiers.Name, "treeView1"));
                    MyTreeView1.Select(@"Node6\Node9\Node10");

                    GUITreeView MyTreeView2 = new GUITreeView(TestApplication, "bottom left treeview", new Identifier(Identifiers.Name, "treeView2"));
                    MyTreeView2.Select(@"Node1\Node4\Node9");

                    GUITreeView MyTreeView3 = new GUITreeView(TestApplication, "top right treeview", new Identifier(Identifiers.Name, "treeView3"));
                    MyTreeView3.Select(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.Select(@"Node1\Node2\Node4444444444444444444444444444444444444444444444");
                    MyTreeView3.Select(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.Select(@"Node1\Node2\Node4444444444444444444444444444444444444444444444");
                    MyTreeView3.Select(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.Check(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.Check(@"Node1\Node3");
                    MyTreeView3.Check(@"Node1\Node2\Node4444444444444444444444444444444444444444444444");

                    GUIListView MyListView1 = new GUIListView(TestApplication, "bottom left listview", new Identifier(Identifiers.Name, "listView1"));
                    MyListView1.SelectGroup("Test Group");
                    MyListView1.SelectGroup("<Default>");
                    MyListView1.Select("ook5");
                    MyListView1.Select("Test Group", "B9");
                    MyListView1.Select("<Default>", "A6");

                    GUIListView MyListView2 = new GUIListView(TestApplication, "bottom right listview", new Identifier(Identifiers.Name, "listView2"));
                    MyListView2.Select("9");

                    //GUI.SetTimeOuts(0);
                    //try
                    //{
                    //    GUIComboBox Fail = new GUIComboBox(TestApplication.Handle, new Identifier(Identifiers.Name, "fail"));
                    //}
                    //catch (Exception ex)
                    //{
                    //    Debug.Listeners[0].WriteLine("\t " + ex.Message + "\r\n" + ex.StackTrace);
                    //}
                    //GUI.SetTimeOuts(30000);

                    TestApplication.Close();
                }

                testTimer.Stop();

                Debug.Listeners[0].WriteLine("\t Time: " + testTimer.ElapsedMilliseconds.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("Failed", ex.Message + "\r\n" + ex.StackTrace);
                mySMTP.Send("david.beales@fidessa.com", "david.beales@fidessa.com", "Failed", ex.Message + "\r\n" + ex.StackTrace);
            }

            textBox1.Text = loops.ToString();

            MessageBox.Show(testTimer.ElapsedMilliseconds.ToString());

            button3.Enabled = true;
        }

        private void btnSpamLog_Click(object sender, EventArgs e)
        {
            GUI.Log("This is a Test!", APE.LogItemTypeEnum.ApeContext);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            button4.Enabled = false;

            for (int x = 0; x < int.Parse(textBox1.Text); x++)
            {

                Process p;

                // close any running sentinels
                while (Process.GetProcessesByName("TestApplication").Length > 0)
                {
                    // kill the process
                    try
                    {
                        p = Process.GetProcessesByName("TestApplication")[0];
                        p.Kill();
                        p.WaitForExit();
                    }
                    catch
                    {
                    }
                }

                //start
                GUI.Log("Launch TestApplication", APE.LogItemTypeEnum.Action);
                
                ProcessStartInfo AppStartup = new ProcessStartInfo();
                AppStartup.WorkingDirectory = @"C:\Tools\TestApplication\TestApplication\bin\Debug\";
                AppStartup.FileName = @"TestApplication.exe";
                Process.Start(AppStartup);

                //find the process
                p = Process.GetProcessesByName("TestApplication")[0];

                //attach
                GUI.AttachToProcess(p);

                //find the form
                GUIForm TestApplication = new GUIForm("test application form", new Identifier(Identifiers.Name, "MainForm"));
                TestApplication.Close();
            }

            button4.Enabled = true;
        }
    }
}


