namespace APE.Spy
{
    partial class APESpy
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.WinformsProcessesCombobox = new System.Windows.Forms.ComboBox();
            this.LocateButton = new System.Windows.Forms.Button();
            this.WindowTree = new System.Windows.Forms.TreeView();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.PropertyListbox = new System.Windows.Forms.ListBox();
            this.ListBoxContextMenuStrip = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.generateLocatorCodeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.IdentifyButton = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.AppDomainComboBox = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.aboutButton = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.ListBoxContextMenuStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // WinformsProcessesCombobox
            // 
            this.WinformsProcessesCombobox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.WinformsProcessesCombobox.FormattingEnabled = true;
            this.WinformsProcessesCombobox.Location = new System.Drawing.Point(63, 11);
            this.WinformsProcessesCombobox.Name = "WinformsProcessesCombobox";
            this.WinformsProcessesCombobox.Size = new System.Drawing.Size(190, 21);
            this.WinformsProcessesCombobox.TabIndex = 1;
            this.WinformsProcessesCombobox.DropDown += new System.EventHandler(this.WinformsProcessesCombobox_DropDown);
            this.WinformsProcessesCombobox.DropDownClosed += new System.EventHandler(this.WinformsProcessesCombobox_DropDownClosed);
            // 
            // LocateButton
            // 
            this.LocateButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.LocateButton.Enabled = false;
            this.LocateButton.Location = new System.Drawing.Point(645, 501);
            this.LocateButton.Name = "LocateButton";
            this.LocateButton.Size = new System.Drawing.Size(52, 21);
            this.LocateButton.TabIndex = 3;
            this.LocateButton.Text = "Locate";
            this.LocateButton.UseVisualStyleBackColor = true;
            this.LocateButton.Click += new System.EventHandler(this.LocateButton_Click);
            // 
            // WindowTree
            // 
            this.WindowTree.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Left | System.Windows.Forms.AnchorStyles.Right)));
            this.WindowTree.Location = new System.Drawing.Point(0, 3);
            this.WindowTree.Name = "WindowTree";
            this.WindowTree.Size = new System.Drawing.Size(429, 455);
            this.WindowTree.TabIndex = 7;
            this.WindowTree.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.WindowTree_AfterSelect);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.splitContainer1.Location = new System.Drawing.Point(13, 40);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.WindowTree);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.PropertyListbox);
            this.splitContainer1.Size = new System.Drawing.Size(684, 455);
            this.splitContainer1.SplitterDistance = 429;
            this.splitContainer1.TabIndex = 8;
            // 
            // PropertyListbox
            // 
            this.PropertyListbox.ContextMenuStrip = this.ListBoxContextMenuStrip;
            this.PropertyListbox.Dock = System.Windows.Forms.DockStyle.Fill;
            this.PropertyListbox.FormattingEnabled = true;
            this.PropertyListbox.Location = new System.Drawing.Point(0, 0);
            this.PropertyListbox.Name = "PropertyListbox";
            this.PropertyListbox.Size = new System.Drawing.Size(251, 455);
            this.PropertyListbox.TabIndex = 0;
            this.PropertyListbox.MouseDown += new System.Windows.Forms.MouseEventHandler(this.PropertyListbox_MouseDown);
            this.PropertyListbox.Resize += new System.EventHandler(this.TreeView_Resize);
            // 
            // ListBoxContextMenuStrip
            // 
            this.ListBoxContextMenuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.copyToolStripMenuItem,
            this.generateLocatorCodeToolStripMenuItem});
            this.ListBoxContextMenuStrip.Name = "ListBoxContextMenuStrip";
            this.ListBoxContextMenuStrip.Size = new System.Drawing.Size(196, 70);
            this.ListBoxContextMenuStrip.Opening += new System.ComponentModel.CancelEventHandler(this.ListBoxContextMenuStrip_Opening);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(195, 22);
            this.copyToolStripMenuItem.Text = "Copy";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.Copy_Click);
            // 
            // generateLocatorCodeToolStripMenuItem
            // 
            this.generateLocatorCodeToolStripMenuItem.Name = "generateLocatorCodeToolStripMenuItem";
            this.generateLocatorCodeToolStripMenuItem.Size = new System.Drawing.Size(195, 22);
            this.generateLocatorCodeToolStripMenuItem.Text = "Generate Locator Code";
            this.generateLocatorCodeToolStripMenuItem.Click += new System.EventHandler(this.generateObjectCodeToolStripMenuItem_Click);
            // 
            // IdentifyButton
            // 
            this.IdentifyButton.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.IdentifyButton.Enabled = false;
            this.IdentifyButton.Location = new System.Drawing.Point(12, 501);
            this.IdentifyButton.Name = "IdentifyButton";
            this.IdentifyButton.Size = new System.Drawing.Size(52, 21);
            this.IdentifyButton.TabIndex = 9;
            this.IdentifyButton.Text = "Identify";
            this.IdentifyButton.UseVisualStyleBackColor = true;
            this.IdentifyButton.Click += new System.EventHandler(this.IdentifyButton_Click);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(70, 505);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(0, 13);
            this.label1.TabIndex = 10;
            // 
            // AppDomainComboBox
            // 
            this.AppDomainComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.AppDomainComboBox.Enabled = false;
            this.AppDomainComboBox.FormattingEnabled = true;
            this.AppDomainComboBox.Location = new System.Drawing.Point(336, 11);
            this.AppDomainComboBox.Name = "AppDomainComboBox";
            this.AppDomainComboBox.Size = new System.Drawing.Size(190, 21);
            this.AppDomainComboBox.TabIndex = 11;
            this.AppDomainComboBox.SelectedIndexChanged += new System.EventHandler(this.AppDomainComboBox_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(10, 15);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(48, 13);
            this.label2.TabIndex = 12;
            this.label2.Text = "Process:";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(266, 15);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(65, 13);
            this.label3.TabIndex = 13;
            this.label3.Text = "AppDomain:";
            // 
            // aboutButton
            // 
            this.aboutButton.Location = new System.Drawing.Point(541, 11);
            this.aboutButton.Name = "aboutButton";
            this.aboutButton.Size = new System.Drawing.Size(52, 21);
            this.aboutButton.TabIndex = 14;
            this.aboutButton.Text = "About";
            this.aboutButton.UseVisualStyleBackColor = true;
            this.aboutButton.Click += new System.EventHandler(this.aboutButton_Click);
            // 
            // APESpy
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(709, 532);
            this.Controls.Add(this.aboutButton);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.AppDomainComboBox);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.IdentifyButton);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.LocateButton);
            this.Controls.Add(this.WinformsProcessesCombobox);
            this.Name = "APESpy";
            this.Text = "APE Spy";
            this.Activated += new System.EventHandler(this.ObjectSpy_Activate);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.ObjectSpy_FormClosing);
            this.Resize += new System.EventHandler(this.TreeView_Resize);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.ListBoxContextMenuStrip.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.ComboBox WinformsProcessesCombobox;
        private System.Windows.Forms.Button LocateButton;
        private System.Windows.Forms.TreeView WindowTree;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ListBox PropertyListbox;
        private System.Windows.Forms.Button IdentifyButton;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ContextMenuStrip ListBoxContextMenuStrip;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ComboBox AppDomainComboBox;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button aboutButton;
        private System.Windows.Forms.ToolStripMenuItem generateLocatorCodeToolStripMenuItem;
    }
}

