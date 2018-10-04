namespace APE.Spy
{
    partial class FormLeaks
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
            this.radioButtonDotNET = new System.Windows.Forms.RadioButton();
            this.radioButtonActiveX = new System.Windows.Forms.RadioButton();
            this.buttonBaseline = new System.Windows.Forms.Button();
            this.buttonLeaks = new System.Windows.Forms.Button();
            this.buttonGCFull = new System.Windows.Forms.Button();
            this.textBoxLeaks = new System.Windows.Forms.TextBox();
            this.labelBaselineCount = new System.Windows.Forms.Label();
            this.labelCompareCount = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // radioButtonDotNET
            // 
            this.radioButtonDotNET.AutoSize = true;
            this.radioButtonDotNET.Checked = true;
            this.radioButtonDotNET.Location = new System.Drawing.Point(12, 12);
            this.radioButtonDotNET.Name = "radioButtonDotNET";
            this.radioButtonDotNET.Size = new System.Drawing.Size(50, 17);
            this.radioButtonDotNET.TabIndex = 0;
            this.radioButtonDotNET.TabStop = true;
            this.radioButtonDotNET.Text = ".NET";
            this.radioButtonDotNET.UseVisualStyleBackColor = true;
            // 
            // radioButtonActiveX
            // 
            this.radioButtonActiveX.AutoSize = true;
            this.radioButtonActiveX.Location = new System.Drawing.Point(68, 12);
            this.radioButtonActiveX.Name = "radioButtonActiveX";
            this.radioButtonActiveX.Size = new System.Drawing.Size(62, 17);
            this.radioButtonActiveX.TabIndex = 1;
            this.radioButtonActiveX.Text = "ActiveX";
            this.radioButtonActiveX.UseVisualStyleBackColor = true;
            // 
            // buttonBaseline
            // 
            this.buttonBaseline.Location = new System.Drawing.Point(12, 35);
            this.buttonBaseline.Name = "buttonBaseline";
            this.buttonBaseline.Size = new System.Drawing.Size(115, 20);
            this.buttonBaseline.TabIndex = 2;
            this.buttonBaseline.Text = "Capture Baseline";
            this.buttonBaseline.UseVisualStyleBackColor = true;
            this.buttonBaseline.Click += new System.EventHandler(this.buttonBaseline_Click);
            // 
            // buttonLeaks
            // 
            this.buttonLeaks.Location = new System.Drawing.Point(12, 61);
            this.buttonLeaks.Name = "buttonLeaks";
            this.buttonLeaks.Size = new System.Drawing.Size(115, 20);
            this.buttonLeaks.TabIndex = 3;
            this.buttonLeaks.Text = "Compare to Baseline";
            this.buttonLeaks.UseVisualStyleBackColor = true;
            this.buttonLeaks.Click += new System.EventHandler(this.buttonLeaks_Click);
            // 
            // buttonGCFull
            // 
            this.buttonGCFull.Location = new System.Drawing.Point(12, 87);
            this.buttonGCFull.Name = "buttonGCFull";
            this.buttonGCFull.Size = new System.Drawing.Size(115, 20);
            this.buttonGCFull.TabIndex = 4;
            this.buttonGCFull.Text = "GC Full";
            this.buttonGCFull.UseVisualStyleBackColor = true;
            this.buttonGCFull.Click += new System.EventHandler(this.buttonGCFull_Click);
            // 
            // textBoxLeaks
            // 
            this.textBoxLeaks.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxLeaks.Location = new System.Drawing.Point(170, 12);
            this.textBoxLeaks.Multiline = true;
            this.textBoxLeaks.Name = "textBoxLeaks";
            this.textBoxLeaks.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxLeaks.Size = new System.Drawing.Size(456, 277);
            this.textBoxLeaks.TabIndex = 5;
            // 
            // labelBaselineCount
            // 
            this.labelBaselineCount.AutoSize = true;
            this.labelBaselineCount.Location = new System.Drawing.Point(133, 39);
            this.labelBaselineCount.Name = "labelBaselineCount";
            this.labelBaselineCount.Size = new System.Drawing.Size(13, 13);
            this.labelBaselineCount.TabIndex = 6;
            this.labelBaselineCount.Text = "0";
            this.labelBaselineCount.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // labelCompareCount
            // 
            this.labelCompareCount.AutoSize = true;
            this.labelCompareCount.Location = new System.Drawing.Point(133, 65);
            this.labelCompareCount.Name = "labelCompareCount";
            this.labelCompareCount.Size = new System.Drawing.Size(13, 13);
            this.labelCompareCount.TabIndex = 7;
            this.labelCompareCount.Text = "0";
            this.labelCompareCount.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // FormLeaks
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(638, 301);
            this.Controls.Add(this.labelCompareCount);
            this.Controls.Add(this.labelBaselineCount);
            this.Controls.Add(this.textBoxLeaks);
            this.Controls.Add(this.buttonGCFull);
            this.Controls.Add(this.buttonLeaks);
            this.Controls.Add(this.buttonBaseline);
            this.Controls.Add(this.radioButtonActiveX);
            this.Controls.Add(this.radioButtonDotNET);
            this.Name = "FormLeaks";
            this.Text = "Leaks";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.RadioButton radioButtonDotNET;
        private System.Windows.Forms.RadioButton radioButtonActiveX;
        private System.Windows.Forms.Button buttonBaseline;
        private System.Windows.Forms.Button buttonLeaks;
        private System.Windows.Forms.Button buttonGCFull;
        private System.Windows.Forms.TextBox textBoxLeaks;
        private System.Windows.Forms.Label labelBaselineCount;
        private System.Windows.Forms.Label labelCompareCount;
    }
}