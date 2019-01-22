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
using System.Windows.Forms;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Net.Mail;
using System.Drawing;
using APE.Language;
using APE.Capture;
using System.Text.RegularExpressions;
using System.Reflection;

namespace APE.Test
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            GUI.Log("Starting", LogItemType.Information);
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
            GUI.Log("Launch Sentinel", LogItemType.Action);
            ProcessStartInfo SentinelStartup = new ProcessStartInfo();
            SentinelStartup.WorkingDirectory = @"C:\Program Files (x86)\LatentZero\LZ Sentinel\Client\";
            SentinelStartup.FileName = @"LzSentinel.exe";
            SentinelStartup.Arguments = @".\lzSentinel.ini";
            Process.Start(SentinelStartup);

            //find the process
            p = Process.GetProcessesByName("LzSentinel")[0];

            //attach
            GUI.AttachToProcess(p);
            GUI.SetTimeOut(4000);
            
            //find the login controls and login
            GUIForm Login = new GUIForm("login form", new Identifier(Identifiers.Name, "SimpleContainer"));
            GUITextBox UserId = new GUITextBox(Login, "user id textbox", new Identifier(Identifiers.Name, "txtUserId"));
            GUITextBox Password = new GUITextBox(Login, "password textbox", new Identifier(Identifiers.Name, "txtPassword"));

            GUIButton OK = new GUIButton(Login, "ok button", new Identifier(Identifiers.Text, "OK"));
            GUIButton Cancel = new GUIButton(Login, "cancel button", new Identifier(Identifiers.Text, "Cancel"));

            UserId.SetText("autosen");
            Password.SetText("quality");

            OK.SingleClick(MouseButton.Left);

            //find main sentinel window and the status bar
            GUI.Log("Wait for Sentinel to load", LogItemType.Information);
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
            int Count = DocumentContainer.ItemsCount();
            DocumentContainer.SingleClickItem("Pre-trade");
            DocumentContainer.SingleClickItem("Rule Maintenance");
            DocumentContainer.SingleClickItem("Sentinel Today");
            DocumentContainer.RemoveItem("Pre-trade");

            //Add a panel
            int InitialItems = DocumentContainer.ItemsCount();
            int CurrentItems = 0;
            GUIToolStripMenu panelsMenu = MenuStrip.GetMenu("panels menu", new Identifier(Identifiers.Name, "PanelsMenu"));
            panelsMenu.SingleClickItem(@"Pre-trade");

            do
            {
                Thread.Sleep(50);

                CurrentItems = DocumentContainer.ItemsCount();
            } while (CurrentItems != InitialItems + 1);

            DocumentContainer.SingleClickItem("Sentinel Today");
    
    

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

            

            layoutGrid.SingleClickCell(group + " -> " + layout, layoutGrid.FirstVisibleColumn(), MouseButton.Left, CellClickLocation.CentreOfCell);
            
            timer = Stopwatch.StartNew();
            do
            {
                if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
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
                        if (timer.ElapsedMilliseconds > GUI.GetTimeOut())
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
            stepTimer = Stopwatch.StartNew();

            //string text = "";
            //Image image = null;

            //////IMS
            Process p = Process.GetProcessesByName("LzCapstone")[0];
            GUI.AttachToProcess(p);
            GUI.SetTimeOut(5000);

            GUIForm quote = new GUIForm("q", new Identifier(Identifiers.Name, "frmSingleQuote"));

            GUIFlexgrid grid = new GUIFlexgrid(quote, "grid", new Identifier(Identifiers.Name, "fgGrid"));
            grid.SetCellValue(1, "Counterparty", "ABN Amro");



            bool exists = GUI.Exists(new Identifier(Identifiers.Name, "CapstoneContainer"));
            

            GUIForm im = new GUIForm("im", new Identifier(Identifiers.Name, "SimpleContainer"));

            GUIElementStripGrid amend = new GUIElementStripGrid(im, "amend", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.Index, 3));
            amend.SetCellValue(amend.Rows() - 1, "Chinese CORP BOND 2 -> Order -> Qty", "3,000.000");



            //GUIDockableWindow dw = new GUIDockableWindow(im, "dw", new Identifier(Identifiers.Text, @"^Orders\ \[Count:\ 673]\[Link:]$"), new Identifier(Identifiers.TypeName, "DockableWindow"));

            //GUIElementStripGrid og = new GUIElementStripGrid(im, "og", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.ChildOf, dw));

            //Thread.Sleep(3000);

            //int sel = og.SelectedRow();
            //og.SelectedRowPollForIndex(sel - 1);


            //Thread.Sleep(5000);

            GUIForm Form = new GUIForm(im, "quick execution form", new Identifier(Identifiers.Name, "frmQuickExecute"));
            GUIFlexgrid AddGrid = new GUIFlexgrid(Form, "quick execution grid", new Identifier(Identifiers.Name, "fgAddGrid"));
            //AddGrid.GetCell


            AddGrid.ShowCell(1, 36);
            AddGrid.ShowCell(1, 36);


            AddGrid.SetCellValue(1, "Clearing Counterparty", "Cazenove");
            AddGrid.SetCellValue(1, "Price", "123");

            GUIForm rel = new GUIForm("rel", new Identifier(Identifiers.Name, "frmSingleRelease"));
            ////GUILzNavBarGridControl c = new GUILzNavBarGridControl(rel, "c", new Identifier(Identifiers.Name, "Counterparties"));
            //GUIFlexgrid grid = new GUIFlexgrid(rel, "grdi", new Identifier(Identifiers.Name, "fgData"));

            ////Rectangle r = grid.GetCellRectangle(grid.FindRow("Global Balanced Fund 7", "Account Name"), grid.FindColumn("Account Name"));

            //this.Focus()

            GUIComboBox cp = new GUIComboBox(rel, "cp", new Identifier(Identifiers.Name, "cboCpties"));
            cp.SingleClickItem("dodo");

            //string ook = grid.GetCellValue(5, 0, CellProperty.RowUserDataNodeCaption);
            Stopwatch ook = Stopwatch.StartNew();
            //c.Select("JP Morgan", MouseButton.Left, CellClickLocation.CentreOfCell);
            Debug.WriteLine(ook.ElapsedMilliseconds.ToString());
            GUIForm merge = new GUIForm("merge", new Identifier(Identifiers.Name, "frmMerge"));

            
            GUIButton mergeandclose = new GUIButton(merge, "m & c", new Identifier(Identifiers.Text, @"Merge\ Orders\ and\ Close"), new Identifier(Identifiers.TypeName, "Button"));
            GUIStretchyCombo type = new GUIStretchyCombo(merge, "xxx", new Identifier(Identifiers.Name, "cmbShow"));
            type.SingleClickItem("Executions");

            Debug.WriteLine("");
            //GUIForm exe = new GUIForm("rel", new Identifier(Identifiers.Name, "frmSingleExecution"));


            //GUITabControl tab = new GUITabControl(exe, "tab", new Identifier(Identifiers.Name, "sftViews"));

            //tab.Select("Summary");
            //tab.Select("Carnegie 2,500 @ 128");
            //tab.Select("ABN Amro 2,500 @ 128");

            //GUILabel nativeLabel = new GUILabel(exe, "message box label", new Identifier(Identifiers.Text, @"No unfilled release(s) found, do you want to continue\?"), new Identifier(Identifiers.TechnologyType, "Windows Native"));

            //GUIFlexgrid grid = new GUIFlexgrid(exe, "grid", new Identifier(Identifiers.Name, "fgAmendGrid"));

            //string text = grid.GetAllVisibleCells(CellProperty.TextDisplay);

            //grid.WaitForControlToNotBeVisible();

            //GUIForm asa = new GUIForm("aas", new Identifier(Identifiers.Name, "SimpleContainer"));
            //GUIElementStripGrid grid = new GUIElementStripGrid(asa, "GridColumnStylesCollection", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.Index, 3));

            ////GUIForm exec = new GUIForm("exec", new Identifier(Identifiers.Name, "frmSingleExecution"));
            ////GUIFlexgrid execGrid = new GUIFlexgrid(exec, "grid", new Identifier(Identifiers.Name, "fgAllocations"));

            //Stopwatch timerx = Stopwatch.StartNew();
            //string ook = grid.GetAllVisibleCells();
            //Debug.WriteLine(timerx.ElapsedMilliseconds.ToString());
            //Clipboard.SetText(ook);

            //GUIForm splash = new GUIForm("splash screen", new Identifier(Identifiers.Name, "SplashScreenForm"));
            //GUIPictureBox splashPicturebox = new GUIPictureBox(splash, "picture box", new Identifier(Identifiers.Name, "picSplash"));

            //Image ook1 = splashPicturebox.Image();
            //Image ook2 = splashPicturebox.BackgroundImage();
            //ook2.Save(@"C:\splash.bmp");


            //Sentinel
            //Process p = Process.GetProcessesByName("LzSentinel")[0];
            //GUI.AttachToProcess(p);
            //GUIForm Sentinel = new GUIForm("sentinel form", new Identifier(Identifiers.Name, "CapstoneContainer"));

            //GUIFlexgrid today = new GUIFlexgrid(Sentinel, "sentinel today", new Identifier(Identifiers.Name, "vfgStats"));

            //for (int x = 0; x < 1000; x++)
            //{
            //    today.Select("Overrides -> New", 0, MouseButton.Left, CellClickLocation.CentreOfCell);
            //}


            //GUIForm aaa = new GUIForm("aaa", new Identifier(Identifiers.Name, "SimpleContainer"));

            //GUIElementStripGrid xxx = new GUIElementStripGrid(aaa, "grid", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.Index, 3));

            //Debug.WriteLine(xxx.Rows().ToString());

            //GUIAxLZResultsGrid resgrid = new GUIAxLZResultsGrid(aaa, "resgrid", new Identifier(Identifiers.Name, "lzcResultsGrid"));

            //Debug.WriteLine(resgrid.Rows().ToString());
            //Debug.WriteLine(xxx.IsEnabled.ToString());

            //
            //
            //
            //GUIForm CheckSecuritiesForm = new GUIForm("check securities form", new Identifier(Identifiers.Text, "Check Securities"));
            //GUIAxLZResultsGrid resultsGrid = new GUIAxLZResultsGrid(CheckSecuritiesForm, "results grid", new Identifier(Identifiers.Name, ""), new Identifier(Identifiers.TypeNameSpace, "LatentZero.Capstone.ComSupport.ResultsGrid"), new Identifier(Identifiers.TypeName, "AxLZResultsGrid"));


            //GUIFlexgrid layout = new GUIFlexgrid(IMS, "treeView", new Identifier(Identifiers.Name, "treeView"));
            //GUIForm ASA = new GUIForm("ASA form", new Identifier(Identifiers.Name, "SimpleContainer"));
            //GUIAxLZResultsGrid resultsGrid = new GUIAxLZResultsGrid(ASA, "lzcResultsGrid", new Identifier(Identifiers.Name, "lzcResultsGrid"));


            //GUIForm me = new GUIForm("IMS form", new Identifier(Identifiers.Name, "frmMultiExecution"));
            //GUIFlexgrid ms = new GUIFlexgrid(me, "fgEditor", new Identifier(Identifiers.Name, "fgEditor"));

            //int cols = ms.Columns();
            //int rows = ms.Rows();

            //Stopwatch timer = Stopwatch.StartNew();
            //for (int x = 0; x < rows; x++)
            //{
            //    for (int y = 0; y < cols; y++)
            //    {
            //        text = ms.GetCellCheck(x, y);
            //    }
            //}
            //Debug.WriteLine((rows * cols).ToString() + " = " + timer.ElapsedMilliseconds.ToString());
            //Debug.WriteLine(text);

            //timer = Stopwatch.StartNew();
            //text = ms.GetCellRangeClip(0, 0, rows - 1, cols - 1);
            //Debug.WriteLine((rows * cols).ToString() + " = " + timer.ElapsedMilliseconds.ToString());
            ////Debug.WriteLine(text);

            //GUIForm mr = new GUIForm("IMS form", new Identifier(Identifiers.Name, "frmMultiRelease"));
            //GUIFlexgrid ml = new GUIFlexgrid(mr, "treeView", new Identifier(Identifiers.Name, "fgMarketListInfo"));
            //text = ml.GetCellValue(0, 0, GUIFlexgrid.CellProperty.TextDisplay);
            //Debug.WriteLine("= " + text);
            //text = ml.GetCellValue(1, 0, GUIFlexgrid.CellProperty.TextDisplay);
            //Debug.WriteLine("= " + text);
            //string ook = ml.GetCellValue(1, 0, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("= " + ook.ToString());
            //string typ = ml.GetColumnType(0);
            //Debug.WriteLine("= " + typ);
            //string format = ml.GetColumnFormat(0);
            //Debug.WriteLine("= " + format);
            //GUIForm sr = new GUIForm("sr", new Identifier(Identifiers.Name, "frmSingleRelease"));
            //GUIFlexgrid st = new GUIFlexgrid(sr, "fgGrid", new Identifier(Identifiers.Name, "fgGrid"));
            //int columns = st.Columns();
            //text = st.GetCellValue(0, columns - 1, GUIFlexgrid.CellProperty.TextDisplay);
            //Debug.WriteLine("= " + text);
            //text = st.GetCellValue(1, columns - 1, GUIFlexgrid.CellProperty.TextDisplay);
            //Debug.WriteLine("= " + text);

            //int col;
            ////GUIForm se = new GUIForm("se", new Identifier(Identifiers.Name, "frmSingleExecution"));
            ////GUIFlexgrid st = new GUIFlexgrid(se, "fgAddGrid", new Identifier(Identifiers.Name, "fgAddGrid"));
            ////GUIFlexgrid ag = new GUIFlexgrid(se, "fgAmendGrid", new Identifier(Identifiers.Name, "fgAmendGrid"));
            ////int columns = st.Columns();
            //Stopwatch timer = Stopwatch.StartNew();
            //GUIForm IMS = new GUIForm("IMS form", new Identifier(Identifiers.Name, "CapstoneContainer"));
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());


            //GUIForm ASA = new GUIForm("ASA", new Identifier(Identifiers.Name, "SimpleContainer"));
            //GUIAxLZResultsGrid sentinelGrid = new GUIAxLZResultsGrid(ASA, "sentinel grid", new Identifier(Identifiers.Name, "lzcResultsGrid"));


            //timer = Stopwatch.StartNew();
            //Debug.WriteLine(sentinelGrid.Columns().ToString());
            //Debug.WriteLine("columns time: " + timer.ElapsedMilliseconds.ToString());

            //timer = Stopwatch.StartNew();
            //GUIForm sr = new GUIForm("single release", new Identifier(Identifiers.Name, "frmSingleRelease"));
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());

            //timer = Stopwatch.StartNew();
            //GUIFlexgrid addamend = new GUIFlexgrid(sr, "add / amend grid", new Identifier(Identifiers.Name, "fgGrid"));
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());

            //timer = Stopwatch.StartNew();
            //GUIFlexgrid al = new GUIFlexgrid(sr, "alloc control grid", new Identifier(Identifiers.Name, "fgData"));
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());

            //timer = Stopwatch.StartNew();
            //string backColours1 = addamend.GetCellRange(0, 0, addamend.Rows() - 1, addamend.Columns() - 1, CellProperty.BackColourName);
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());
            ////Clipboard.SetDataObject(backColours, true, 2,100);

            //timer = Stopwatch.StartNew();
            //string DataDisplay1 = addamend.GetCellRange(0, 0, addamend.Rows() - 1, addamend.Columns() - 1, CellProperty.TextDisplay);
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());


            //timer = Stopwatch.StartNew();
            //string backColours = al.GetCellRange(0, 0, al.Rows() - 1, al.Columns() - 1, CellProperty.BackColourName);
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());
            ////Clipboard.SetDataObject(backColours, true, 2,100);

            //timer = Stopwatch.StartNew();
            //string DataDisplay = al.GetCellRange(0, 0, al.Rows() - 1, al.Columns() - 1, CellProperty.TextDisplay);
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());

            //timer = Stopwatch.StartNew();
            //col = st.FindColumn("Counterparty");
            //col = st.FindColumn("Stmp");
            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("add row 0 TD = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("add row 1 TD = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("ame row 0 TD = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("add row 0 CB = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("add row 1 CB = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("ame row 0 CB = " + text);

            //Debug.WriteLine("");

            //image = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.Image);
            //if (image == null)
            //{
            //    Debug.WriteLine("add row 0 IM = ");
            //}
            //else
            //{
            //    Debug.WriteLine("add row 0 IM = " + image.ToString());
            //    image.Save(@"C:\expand.bmp");
            //}

            //image = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.Image);
            //if (image == null)
            //{
            //    Debug.WriteLine("add row 1 IM = ");
            //}
            //else
            //{
            //    Debug.WriteLine("add row 1 IM = " + text.ToString());
            //}

            //image = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.Image);
            //if (image == null)
            //{
            //    Debug.WriteLine("ame row 0 IM = ");
            //}
            //else
            //{
            //    Debug.WriteLine("ame row 0 IM = " + text.ToString());
            //    image.Save(@"C:\counterparty.bmp");
            //}

            //Debug.WriteLine("");

            //image = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackgroundImage);
            //if (image == null)
            //{
            //    Debug.WriteLine("add row 0 BI = ");
            //}
            //else
            //{
            //    Debug.WriteLine("add row 0 BI = " + image.ToString());
            //}
            //image = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.BackgroundImage);
            //if (image == null)
            //{
            //    Debug.WriteLine("add row 1 BI = ");
            //}
            //else
            //{
            //    Debug.WriteLine("add row 1 BI = " + image.ToString());
            //}
            //image = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackgroundImage);
            //if (image == null)
            //{
            //    Debug.WriteLine("ame row 0 BI = ");
            //}
            //else
            //{
            //    Debug.WriteLine("ame row 0 BI = " + image.ToString());
            //    image.Save(@"C:\oooooooook.bmp");
            //}

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("add row 0 CB = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("add row 1 CB = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("ame row 0 CB = " + text);

            //Debug.WriteLine("");


            //col = st.FindColumn("Whs");
            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("add row 0 TD = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("add row 1 TD = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("ame row 0 TD = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("add row 0 CB = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("add row 1 CB = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("ame row 0 CB = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.Image).ToString();
            //Debug.WriteLine("add row 0 Im = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.Image).ToString();
            //Debug.WriteLine("add row 1 Im = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.Image).ToString();
            //Debug.WriteLine("ame row 0 Im = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackgroundImage).ToString();
            //Debug.WriteLine("add row 0 BI = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.BackgroundImage).ToString();
            //Debug.WriteLine("add row 1 BI = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackgroundImage).ToString();
            //Debug.WriteLine("ame row 0 BI = " + text);

            //Debug.WriteLine("");

            //col = st.FindColumn("Counterparty");
            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("add row 0 TD = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("add row 1 TD = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.DataDisplay);
            //Debug.WriteLine("ame row 0 TD = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("add row 0 CB = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("add row 1 CB = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.CheckBox);
            //Debug.WriteLine("ame row 0 CB = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.Image).ToString();
            //Debug.WriteLine("add row 0 Im = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.Image).ToString();
            //Debug.WriteLine("add row 1 Im = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.Image).ToString();
            //Debug.WriteLine("ame row 0 Im = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackgroundImage).ToString();
            //Debug.WriteLine("add row 0 BI = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.BackgroundImage).ToString();
            //Debug.WriteLine("add row 1 BI = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackgroundImage).ToString();
            //Debug.WriteLine("ame row 0 BI = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.ForeColor).Name;
            //Debug.WriteLine("add row 0 FC = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.ForeColor).Name;
            //Debug.WriteLine("add row 1 FC = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.ForeColor).Name;
            //Debug.WriteLine("ame row 0 FC = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackColor).Name;
            //Debug.WriteLine("add row 0 BC = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.BackColor).Name;
            //Debug.WriteLine("add row 1 BC = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.BackColor).Name;
            //Debug.WriteLine("ame row 0 BC = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellValue(0, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("add row 0 CB = " + text);
            //text = st.GetCellValue(1, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("add row 1 CB = " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("ame row 0 CB = " + text);

            //Debug.WriteLine("");

            //text = st.GetCellCheck(0, "Stmp");
            //Debug.WriteLine("= " + text);
            //text = st.GetCellCheck(1, "Stmp");
            //Debug.WriteLine("= " + text);
            //text = ag.GetCellCheck(0, col);
            //Debug.WriteLine("= " + text);

            //text = st.GetCellValue(0, "Stmp", GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("= " + text);
            //text = st.GetCellValue(1, "Stmp", GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("= " + text);
            //text = ag.GetCellValue(0, col, GUIFlexgrid.CellProperty.Clip);
            //Debug.WriteLine("= " + text);

            //resultsGrid.CopyToClipboard();
            //Debug.WriteLine(resultsGrid.Rows().ToString());
            //Debug.WriteLine(resultsGrid.Columns().ToString());
            //Debug.WriteLine(resultsGrid.FirstVisibleColumn().ToString());

            ////for (int i = resultsGrid.FirstVisibleColumn(); i < resultsGrid.Columns(); i++)
            //for (int i = 0; i < resultsGrid.Columns(); i++)
            //{
            //    //if (!resultsGrid.IsColumnHidden(i))
            //    //{
            //        Debug.WriteLine(resultsGrid.GetCellValue(0, i, GUIAxLZResultsGrid.VSFlexgridCellPropertySettings.flexcpTextDisplay));
            //        //Debug.WriteLine(resultsGrid.GetCellRangeClip(0, i, 0, i) + " " + resultsGrid.ColumnWidth(i).ToString());
            //    //}
            //}

            //Debug.WriteLine(resultsGrid.GetCellRangeClip(0, 0, 5, resultsGrid.Columns() - 1));



            //GUIAxLZResultsGrid res = new GUIAxLZResultsGrid(ASA, "lzcResultsGrid", new Identifier(Identifiers.Name, "faddcdb9-0d32-46b6-941a-57a7784bc2ff"));
            //GUIAxLZResultsGrid res = new GUIAxLZResultsGrid(ASA, "lzcResultsGrid", new Identifier(Identifiers.Name, "lzcResultsGrid"));
            //Stopwatch timer = Stopwatch.StartNew();
            //res.foo2();
            //Debug.WriteLine(timer.ElapsedMilliseconds.ToString());



            //Debug.WriteLine("done");


            ////GUIFlexgrid grid = new GUIFlexgrid(Sentinel, "treeView", new Identifier(Identifiers.Name, "vfgTree"));
            ////GUIFlexgrid grid = new GUIFlexgrid(Sentinel, "treeView", new Identifier(Identifiers.Name, "vfgRuleGrid"));

            //GUIDocumentContainer dc = new GUIDocumentContainer(Sentinel, "dc", new Identifier(Identifiers.TypeName, "DocumentContainer"));

            //dc.ItemSelect("monitor dooooooooooooooooooooooooooooooooo csdddddddddddddddc");
            //dc.ItemSelect("Sentinel Today");
            //dc.ItemSelect("monitor dooooooooooooooooooooooooooooooooo csdddddddddddddddc");
            //dc.ItemSelect("Sentinel Today");

            //GUITitleFrame tf = new GUITitleFrame(Sentinel, "TitleFrame", new Identifier(Identifiers.Name, "TitleFrame"));
            //Debug.WriteLine(tf.CanMaximise().ToString());
            //Debug.WriteLine(tf.IsMaximised().ToString());

            //GUITitleFrameButton max = tf.GetButton(" titleframe min max button", new Identifier(Identifiers.Name, "MinMax"));
            //Debug.WriteLine(max.ToolTipText);
            //max.MouseSingleClick(MouseButton.Left);
            //Debug.WriteLine(tf.IsMaximised().ToString());
            //Thread.Sleep(1000);
            //max.MouseSingleClick(MouseButton.Left);
            //Thread.Sleep(1000);
            //Debug.WriteLine("st: " + tf.Subtitle());
            //max.MouseMove();
            //Debug.WriteLine("st: " + tf.Subtitle());

            //GUIExpando xp = new GUIExpando(Sentinel, "Expando", new Identifier(Identifiers.Name, "ScreenSwitcherControl"));
            //xp.Collapse();
            //xp.Expand();

            //grid.ExpandTreeView();
            //grid.CollapseTreeView();

            //GUIForm ASA = new GUIForm("ASA form", new Identifier(Identifiers.Name, "SimpleContainer"));

            //GUIElementStripGrid entry = new GUIElementStripGrid(ASA, "entry strip", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.Index, 2));

            //entry.SetCellValue(1, "Instrument", "AEGON 5 3/4 12/15/20");

            //Debug.WriteLine("");


            //for (int i = 0; i < layout.Rows(); i++)
            //{
            //    Debug.WriteLine(i.ToString() + " " + layout.IsRowHidden(i).ToString());
            //}

            //layout.CollapseTreeView();
            //layout.ExpandNodes("Manual -> Level 1 -> Level 2b");
            //layout.CollapseNodes("Manual -> Level 1 -> Level 2b");
            //stepTimer = Stopwatch.StartNew();
            //layout.ExpandTreeView();
            //Debug.WriteLine("expand: " + stepTimer.ElapsedMilliseconds.ToString());
            //stepTimer = Stopwatch.StartNew();
            //layout.CollapseTreeView();
            //Debug.WriteLine("collapse: " + stepTimer.ElapsedMilliseconds.ToString());




            //GUIForm execution = new GUIForm("execution form", new Identifier(Identifiers.Name, "frmSingleExecution"));
            //GUIFlexgrid add = new GUIFlexgrid(execution, "add grid", new Identifier(Identifiers.Name, "fgAddGrid"));
            //add.SetCellValue(1, "Counterparty", "ABN Amro");
            //add.SetCellValue(1, "Price", "128.3", "128.3000");
            //add.SetCellValue(1, "Stmp", "True");
            //add.SetCellValue(1, "Trade Date", "16 Jun 2016", "16 Jun 2016");

            //GUIForm IMS = new GUIForm("IMS form", new Identifier(Identifiers.Name, "CapstoneContainer"));
            //GUIFlexgrid layout = new GUIFlexgrid(IMS, "layout treeview", new Identifier(Identifiers.Name, "treeView"));

            ////https://msdn.microsoft.com/en-us/library/az24scfc(v=vs.110).aspx
            ////Debug.WriteLine(System.Text.RegularExpressions.Regex.IsMatch("Orders [Count: 3][Link:]", "Orders [[]Count: [0-9]+[]][[]Link:[]]").ToString());


            //GUIDockableWindow OrderViewerDockableWindow = new GUIDockableWindow(IMS, "Order viewer dockable window", new Identifier(Identifiers.Text, "Orders [[]Count: [0-9]+[]][[]Link:[]]"), new Identifier(Identifiers.TypeName, "DockableWindow"));
            //GUIElementStripGrid OrderViewerGrid = new GUIElementStripGrid(IMS, "Order viewer grid", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.ChildOf, OrderViewerDockableWindow));


            //GUIDockableWindow ExecutionViewerDockableWindow = new GUIDockableWindow(IMS, "Execution viewer dockable window", new Identifier(Identifiers.Text, "Executions [[]Count: [0-9]+[]][[]Link:[]]"), new Identifier(Identifiers.TypeName, "DockableWindow"));
            //GUIElementStripGrid ExecutionGrid = new GUIElementStripGrid(IMS, "Execution viewer grid", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.ChildOf, ExecutionViewerDockableWindow));





            //GUIForm ASA = new GUIForm("ASA form", new Identifier(Identifiers.Name, "SimpleContainer"));

            //GUIElementStripGrid entry = new GUIElementStripGrid(ASA, "entry strip", new Identifier(Identifiers.Name, "m_elementStripGrid"), new Identifier(Identifiers.Index, 2));

            //Debug.WriteLine(entry.FindRow("Order"));
            //Debug.WriteLine(entry.FindColumn("Instrument"));

            ////entry.SetCellValue(1, 0, "Execution", "Execution", null);
            //entry.SetCellValue(1, 4, "AEGON 5 3/4 12/15/20", "AEGON 5 3/4 12/15/20", null);

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

        private void Log(string textToLog, LogItemType type)
        {
            WriteToFile(textToLog);
        }

        private void button3_Click(object sender, EventArgs e)  //Test Application
        {
            button3.Enabled = false;

            SmtpClient mySMTP = new SmtpClient("smtp.fidessa.com", 25);
            int loops = int.Parse(textBox1.Text);
            Stopwatch testTimer = Stopwatch.StartNew();

            GUI.TreeViewDelimiter = @"\";
            GUI.MenuDelimiter = @"\";
            GUI.GridDelimiter = @" -> ";

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
                    GUI.Log("Launch TestApplication", LogItemType.Action);

                    ProcessStartInfo AppStartup = new ProcessStartInfo();
                    AppStartup.WorkingDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + @"\..\..\..\..\TestApplication\TestApplication\bin\Debug\";
                    AppStartup.FileName = @"TestApplication.exe";
                    Process.Start(AppStartup);

                    //find the process
                    p = Process.GetProcessesByName("TestApplication")[0];

                    //attach
                    GUI.AttachToProcess(p);
                    
                    GUI.SetTimeOut(20000);

                    //find the form
                    GUIForm TestApplication = new GUIForm("test application main form", new Identifier(Identifiers.Name, "MainForm"));
                    GUITextBox tbxContext = new GUITextBox(TestApplication, "textbox", new Identifier(Identifiers.Name, "tbxContext"));
                    GUIButton btnStatusBar = new GUIButton(TestApplication, "increment statusbar button", new Identifier(Identifiers.Text, "Increment StatusBar"));

                    tbxContext.SetText("autosen");
                    tbxContext.SetText("whoop");
                    tbxContext.SetText("moo");

                    GUIMenuStrip MenuStrip = new GUIMenuStrip(TestApplication, "menu strip", new Identifier(Identifiers.TypeName, "MenuStrip"));
                    GUIToolStripMenu secondMenu = MenuStrip.GetMenu("second", new Identifier(Identifiers.Name, "secondToolStripMenuItem"));
                    secondMenu.SingleClickItem(@"OpensPopup", ItemIdentifier.AccessibilityObjectName);

                    GUIForm CustomMessageBox = new GUIForm("messagebox form", new Identifier(Identifiers.Name, "frmCustomMessageBox"));
                    CustomMessageBox.Maximise();
                    CustomMessageBox.Restore();

                    CustomMessageBox.Move(500, 500);

                    CustomMessageBox.Minimise();
                    CustomMessageBox.Restore();

                    CustomMessageBox.Close();

                    GUITabControl tab1 = new GUITabControl(TestApplication, "tab control", new Identifier(Identifiers.Name, "tabControl1"));
                    tab1.SingleClickTab("tabPage3");
                    tab1.SingleClickTab("tabPage1");

                    GUITabControl tab2 = new GUITabControl(TestApplication, "tab control", new Identifier(Identifiers.Name, "tabControl2"));
                    tab2.SingleClickTab("tabPage4");
                    tab2.SingleClickTab("tabPage7");

                    GUIStatusStrip ssStatusBar = new GUIStatusStrip(TestApplication, "status strip", new Identifier(Identifiers.Name, "StatusStrip"));

                    string PanelText;

                    GUIToolStripLabel statusBarFirstPanel = ssStatusBar.GetLabel("first panel", new Identifier(Identifiers.Name, "toolStripStatusLabel1"));

                    GUIToolStripLabel statusBarFirstPanelIdentTest = ssStatusBar.GetLabel("first panel for ident test", new Identifier(Identifiers.Name, "toolStripStatusLabel1"), new Identifier(Identifiers.TechnologyType, "Windows Forms (WinForms)"), new Identifier(Identifiers.TypeNameSpace, "System.Windows.Forms"), new Identifier(Identifiers.TypeName, "ToolStripStatusLabel"), new Identifier(Identifiers.ModuleName, "System.Windows.Forms.dll"), new Identifier(Identifiers.AssemblyName, "System.Windows.Forms"), new Identifier(Identifiers.Text, "toolStripStatusLabel1"));

                    PanelText = statusBarFirstPanel.Text;

                    //int PanelIndex = ssStatusBar.PanelIndex("toolStripStatusLabel1");
                    //PanelText = ssStatusBar.PanelText(PanelIndex);
                    //PanelText = ssStatusBar.PanelText("toolStripStatusLabel5");

                    btnStatusBar.SingleClick(MouseButton.Left);
                    //statusBarFirstPanel.PollForText("6");
                    //statusBarFirstPanel.PollForText("");
                    //Stopwatch doo = Stopwatch.StartNew();
                    statusBarFirstPanel.PollForText("");
                    //Debug.WriteLine(doo.ElapsedMilliseconds.ToString());
//                    while (ssStatusBar.PanelText(0) != "")
//                    {
//                        Thread.Sleep(100);
//                    }

                    //bring up the context menu of the textbox and select the menu item
                    //TODO make this a single command??
                    tbxContext.SingleClick(MouseButton.Right);
                    tbxContext.ContextMenuSelect(@"ConItem2\ConSubItem2");
                    
                    CustomMessageBox = new GUIForm("messagebox form", new Identifier(Identifiers.Name, "frmCustomMessageBox"));
                    CustomMessageBox.Close();

                    GUILabel EditBoxLabel = new GUILabel(TestApplication, "label1 label", new Identifier(Identifiers.Name, "lblTextBox"));
                    string labelText = EditBoxLabel.Text;

                    GUICheckBox MyCheckBox = new GUICheckBox(TestApplication, "checkbox1 checkbox", new Identifier(Identifiers.Name, "checkBox1"));
                    MyCheckBox.SingleClick(MouseButton.Left);

                    GUIRadioButton MyRadioButton = new GUIRadioButton(TestApplication, "radioButton1 radio button", new Identifier(Identifiers.Name, "radioButton1"));
                    MyRadioButton.SingleClick(MouseButton.Left);

                    GUIComboBox MyComboBox1 = new GUIComboBox(TestApplication, "top left combobox", new Identifier(Identifiers.Name, "comboBox1"));
                    MyComboBox1.SetText("foo");
                    MyComboBox1.SetText("bar");
                    MyComboBox1.SingleClickItem("Test31");

                    GUIComboBox MyComboBox2 = new GUIComboBox(TestApplication, "bottom left combobox", new Identifier(Identifiers.Name, "comboBox2"));
                    MyComboBox2.SingleClickItem("Test40");

                    GUIComboBox MyComboBox3 = new GUIComboBox(TestApplication, "top right combobox", new Identifier(Identifiers.Name, "comboBox3"));
                    MyComboBox3.SingleClickItem("Test50");
                    
                    MyComboBox3 = new GUIComboBox(TestApplication, "top right combobox", new Identifier(Identifiers.Name, "comboBox3"));
                    MyComboBox3.SingleClickItem("Test25");

                    GUIListBox MyListBox = new GUIListBox(TestApplication, "bottom right listbox", new Identifier(Identifiers.Name, "listBox1"));
                    MyListBox.SingleClickItem("Test25");

                    GUITreeView MyTreeView1 = new GUITreeView(TestApplication, "top left treeview", new Identifier(Identifiers.Name, "treeView1"));
                    MyTreeView1.SingleClickItem(@"Node6\Node9\Node10");

                    GUITreeView MyTreeView2 = new GUITreeView(TestApplication, "bottom left treeview", new Identifier(Identifiers.Name, "treeView2"));
                    MyTreeView2.SingleClickItem(@"Node1\Node4\Node9");

                    GUITreeView MyTreeView3 = new GUITreeView(TestApplication, "top right treeview", new Identifier(Identifiers.Name, "treeView3"));
                    MyTreeView3.SingleClickItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.SingleClickItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444");
                    MyTreeView3.SingleClickItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.SingleClickItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444");
                    MyTreeView3.SingleClickItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.CheckItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444\Node5");
                    MyTreeView3.CheckItem(@"Node1\Node3");
                    MyTreeView3.CheckItem(@"Node1\Node2\Node4444444444444444444444444444444444444444444444");

                    GUIListView MyListView1 = new GUIListView(TestApplication, "bottom left listview", new Identifier(Identifiers.Name, "listView1"));
                    MyListView1.SingleClickGroup("Test Group");
                    MyListView1.SingleClickGroup("<Default>");
                    MyListView1.SingleClickItem("ook5");
                    MyListView1.SingleClickItem("Test Group", "B9");
                    MyListView1.SingleClickItem("<Default>", "A6");
                    
                    GUIListView MyListView2 = new GUIListView(TestApplication, "bottom right listview", new Identifier(Identifiers.Name, "listView2"));
                    MyListView2.SingleClickItem("9");

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
            GUI.Log("This is a Test Action!", LogItemType.Action);
            GUI.Log("This is a Test Debug!", LogItemType.Debug);
            GUI.Log("This is a Test Error!", LogItemType.Error);
            GUI.Log("This is a Test Fail!", LogItemType.Fail);
            GUI.Log("This is a Test Finish!", LogItemType.Finish);
            GUI.Log("This is a Test Information!", LogItemType.Information);
            GUI.Log("This is a Test Pass!", LogItemType.Pass);
            GUI.Log("This is a Test Start!", LogItemType.Start);
            GUI.Log("This is a Test Warning!", LogItemType.Warning);
            GUI.Log("This is a Test Disabled!", LogItemType.Disabled);
            GUI.Log("This is a Test NA!", LogItemType.NA);
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
                GUI.Log("Launch TestApplication", LogItemType.Action);
                
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


