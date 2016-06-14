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
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Diagnostics;

namespace TestApplication
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            GC.Collect();
            this.HelpButton = true;// only pres
        }

        private void opensPopupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmCustomMessageBox myCustomMessageBox = new frmCustomMessageBox();

            myCustomMessageBox.ShowDialog(this);
        }

        private void btnStatusBar_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < 10; i++)
            {
                StatusStrip.Items[0].Text = i.ToString();
                StatusStrip.Update();
                for (int x = 0; x < 3; x++)
                {
                    Thread.Sleep(25);
                    Application.DoEvents();
                }
            }
            StatusStrip.Items[0].Text = "";
        }

        private void lblTextBox_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(tbxContext.ContextMenu.Handle.ToString());
            //MessageBox.Show(tbxContext.ContextMenuStrip.Handle.ToString());
        }

        private void conSubItem2ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmCustomMessageBox myCustomMessageBox = new frmCustomMessageBox();

            myCustomMessageBox.ShowDialog(this);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            bool cbstate = checkBox1.Checked;

            //MessageBox.Show(comboBox1.AccessibilityObject.GetChildCount().ToString());

            //comboBox1.GetItemHeight;
            //comboBox1.GetItemText;
            //comboBox1.Items;
            for (int i = 0; i < comboBox1.Items.Count; i++)
            {
                //Debug.Listeners[0].WriteLine(comboBox1.Items[i].ToString());
            }

            //comboBox1.DroppedDown
            //comboBox1.FindStringExact
            //MessageBox.Show(comboBox1.DropDownStyle.ToString("G"));
            //MessageBox.Show((int)(ComboBoxStyle.DropDownList).ToString());
            //MessageBox.Show((int)(ComboBoxStyle.Simple).ToString());
        }


        public delegate void ProcessBookDelegate();
        ProcessBookDelegate ps1 = new ProcessBookDelegate(foo);

        private static void foo()
        {
            
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
//            MessageBox.Show(typeof(ComboBoxStyle).IsEnum.ToString());
            //comboBox1.DropDownStyle.ToString

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            //treeView1.CheckBoxes = false;
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            //listView1.items
            //ListViewItem foo;
            //listView1.SelectedItems.Clear();
            //listView1.Items.count
            //MessageBox.Show(ItemBoundsPortion.Label.)
            //foo.GetBounds(2);
            ////MessageBox.Show(listView1.Groups.Count.ToString());
            ////listView1.Groups.item;
            //ListViewGroup ook;
            //MessageBox.Show(listView1.Items.Count.ToString());
            //MessageBox.Show(listView1.Items[2].Group.Header);
            //ook.Header;
            //ook.Items.item;
            //foo.Text;
            //MessageBox.Show(listView1.Groups[0].Header);
            //foo.
            //for (int x = 0; x < listView1.Items.Count; x++)
            //{
            //    Debug.Listeners[0].WriteLine("\t text: " + listView1.Items[x].Text);
            //}
            //foo.Index
            //ook.a
        }
    }
}
