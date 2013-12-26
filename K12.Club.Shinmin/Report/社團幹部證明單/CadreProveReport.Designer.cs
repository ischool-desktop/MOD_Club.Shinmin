namespace K12.Club.Shinmin
{
    partial class CadreProveReport
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
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.cbPrintByPresident = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.cbPrintAll = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.SuspendLayout();
            // 
            // linkLabel1
            // 
            this.linkLabel1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.BackColor = System.Drawing.Color.Transparent;
            this.linkLabel1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.linkLabel1.Location = new System.Drawing.Point(9, 112);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(60, 17);
            this.linkLabel1.TabIndex = 10;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "設定樣板";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnExit.Location = new System.Drawing.Point(275, 106);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.TabIndex = 9;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.labelX1.Location = new System.Drawing.Point(12, 14);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(114, 21);
            this.labelX1.TabIndex = 8;
            this.labelX1.Text = "請選擇列印內容：";
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnPrint.Location = new System.Drawing.Point(194, 106);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.TabIndex = 7;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // cbPrintByPresident
            // 
            this.cbPrintByPresident.AutoSize = true;
            this.cbPrintByPresident.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbPrintByPresident.BackgroundStyle.Class = "";
            this.cbPrintByPresident.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbPrintByPresident.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.cbPrintByPresident.Location = new System.Drawing.Point(194, 60);
            this.cbPrintByPresident.Name = "cbPrintByPresident";
            this.cbPrintByPresident.Size = new System.Drawing.Size(140, 21);
            this.cbPrintByPresident.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbPrintByPresident.TabIndex = 11;
            this.cbPrintByPresident.TabStop = false;
            this.cbPrintByPresident.Text = "僅列印社長/副社長";
            // 
            // cbPrintAll
            // 
            this.cbPrintAll.AutoSize = true;
            this.cbPrintAll.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.cbPrintAll.BackgroundStyle.Class = "";
            this.cbPrintAll.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.cbPrintAll.CheckBoxStyle = DevComponents.DotNetBar.eCheckBoxStyle.RadioButton;
            this.cbPrintAll.Checked = true;
            this.cbPrintAll.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbPrintAll.CheckValue = "Y";
            this.cbPrintAll.Location = new System.Drawing.Point(40, 60);
            this.cbPrintAll.Name = "cbPrintAll";
            this.cbPrintAll.Size = new System.Drawing.Size(134, 21);
            this.cbPrintAll.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbPrintAll.TabIndex = 12;
            this.cbPrintAll.Text = "列印所有社團幹部";
            // 
            // CadreProveReport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(362, 141);
            this.Controls.Add(this.cbPrintAll);
            this.Controls.Add(this.cbPrintByPresident);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnPrint);
            this.Name = "CadreProveReport";
            this.Text = "社團幹部證明單";
            this.Load += new System.EventHandler(this.CadreProveReport_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.LinkLabel linkLabel1;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbPrintByPresident;
        private DevComponents.DotNetBar.Controls.CheckBoxX cbPrintAll;
    }
}