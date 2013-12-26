namespace K12.Club.Shinmin
{
    partial class CheckStudentIsNotInClub
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle47 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle37 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle48 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle49 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle50 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle51 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle42 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle43 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle44 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle45 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle46 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridViewX1 = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.清除指定社團ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.加入待處理學生ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.btnOutPut = new DevComponents.DotNetBar.ButtonX();
            this.itmPnlTimeName = new DevComponents.DotNetBar.ItemPanel();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn4 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn5 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn6 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGrade_year = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSeat_no = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStudentNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSelectClub = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).BeginInit();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // dataGridViewX1
            // 
            this.dataGridViewX1.AllowUserToAddRows = false;
            this.dataGridViewX1.AllowUserToDeleteRows = false;
            this.dataGridViewX1.AllowUserToResizeRows = false;
            this.dataGridViewX1.BackgroundColor = System.Drawing.Color.White;
            this.dataGridViewX1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewX1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colGrade_year,
            this.colClass,
            this.colSeat_no,
            this.colName,
            this.colStudentNumber,
            this.colSelectClub});
            this.dataGridViewX1.ContextMenuStrip = this.contextMenuStrip1;
            dataGridViewCellStyle47.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle47.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle47.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle47.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle47.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle47.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle47.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewX1.DefaultCellStyle = dataGridViewCellStyle47;
            this.dataGridViewX1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dataGridViewX1.Location = new System.Drawing.Point(12, 40);
            this.dataGridViewX1.Name = "dataGridViewX1";
            this.dataGridViewX1.RowHeadersVisible = false;
            this.dataGridViewX1.RowTemplate.Height = 24;
            this.dataGridViewX1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dataGridViewX1.Size = new System.Drawing.Size(502, 396);
            this.dataGridViewX1.TabIndex = 1;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.清除指定社團ToolStripMenuItem,
            this.加入待處理學生ToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(155, 48);
            // 
            // 清除指定社團ToolStripMenuItem
            // 
            this.清除指定社團ToolStripMenuItem.Name = "清除指定社團ToolStripMenuItem";
            this.清除指定社團ToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.清除指定社團ToolStripMenuItem.Text = "清除指定社團";
            this.清除指定社團ToolStripMenuItem.Click += new System.EventHandler(this.清除指定社團ToolStripMenuItem_Click);
            // 
            // 加入待處理學生ToolStripMenuItem
            // 
            this.加入待處理學生ToolStripMenuItem.Name = "加入待處理學生ToolStripMenuItem";
            this.加入待處理學生ToolStripMenuItem.Size = new System.Drawing.Size(154, 22);
            this.加入待處理學生ToolStripMenuItem.Text = "加入待處理學生";
            this.加入待處理學生ToolStripMenuItem.Click += new System.EventHandler(this.加入待處理學生ToolStripMenuItem_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.AutoSize = true;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(585, 445);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(666, 445);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 6;
            this.btnExit.Text = "關閉";
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
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(87, 21);
            this.labelX1.TabIndex = 0;
            this.labelX1.Text = "未選社清單：";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(520, 13);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(74, 21);
            this.labelX2.TabIndex = 2;
            this.labelX2.Text = "社團清單：";
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(116, 447);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(108, 21);
            this.labelX3.TabIndex = 4;
            this.labelX3.Text = "待處理學生：0人";
            // 
            // btnOutPut
            // 
            this.btnOutPut.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnOutPut.AutoSize = true;
            this.btnOutPut.BackColor = System.Drawing.Color.Transparent;
            this.btnOutPut.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnOutPut.Location = new System.Drawing.Point(13, 445);
            this.btnOutPut.Name = "btnOutPut";
            this.btnOutPut.Size = new System.Drawing.Size(75, 25);
            this.btnOutPut.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnOutPut.TabIndex = 7;
            this.btnOutPut.Text = "匯出";
            this.btnOutPut.Click += new System.EventHandler(this.btnOutPut_Click);
            // 
            // itmPnlTimeName
            // 
            this.itmPnlTimeName.AutoScroll = true;
            this.itmPnlTimeName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.itmPnlTimeName.BackgroundStyle.Class = "ItemPanel";
            this.itmPnlTimeName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.itmPnlTimeName.ContainerControlProcessDialogKey = true;
            this.itmPnlTimeName.LayoutOrientation = DevComponents.DotNetBar.eOrientation.Vertical;
            this.itmPnlTimeName.LicenseKey = "F962CEC7-CD8F-4911-A9E9-CAB39962FC1F";
            this.itmPnlTimeName.Location = new System.Drawing.Point(520, 40);
            this.itmPnlTimeName.Name = "itmPnlTimeName";
            this.itmPnlTimeName.Size = new System.Drawing.Size(221, 396);
            this.itmPnlTimeName.TabIndex = 8;
            this.itmPnlTimeName.Text = "itemPanel1";
            // 
            // dataGridViewTextBoxColumn1
            // 
            dataGridViewCellStyle37.BackColor = System.Drawing.Color.LightCyan;
            this.dataGridViewTextBoxColumn1.DefaultCellStyle = dataGridViewCellStyle37;
            this.dataGridViewTextBoxColumn1.HeaderText = "班級";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 80;
            // 
            // dataGridViewTextBoxColumn2
            // 
            dataGridViewCellStyle48.BackColor = System.Drawing.Color.LightCyan;
            this.dataGridViewTextBoxColumn2.DefaultCellStyle = dataGridViewCellStyle48;
            this.dataGridViewTextBoxColumn2.HeaderText = "座號";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.ReadOnly = true;
            this.dataGridViewTextBoxColumn2.Width = 60;
            // 
            // dataGridViewTextBoxColumn3
            // 
            dataGridViewCellStyle49.BackColor = System.Drawing.Color.LightCyan;
            this.dataGridViewTextBoxColumn3.DefaultCellStyle = dataGridViewCellStyle49;
            this.dataGridViewTextBoxColumn3.HeaderText = "姓名";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            this.dataGridViewTextBoxColumn3.ReadOnly = true;
            this.dataGridViewTextBoxColumn3.Width = 60;
            // 
            // dataGridViewTextBoxColumn4
            // 
            dataGridViewCellStyle50.BackColor = System.Drawing.Color.LightCyan;
            this.dataGridViewTextBoxColumn4.DefaultCellStyle = dataGridViewCellStyle50;
            this.dataGridViewTextBoxColumn4.HeaderText = "學號";
            this.dataGridViewTextBoxColumn4.Name = "dataGridViewTextBoxColumn4";
            this.dataGridViewTextBoxColumn4.ReadOnly = true;
            this.dataGridViewTextBoxColumn4.Width = 80;
            // 
            // dataGridViewTextBoxColumn5
            // 
            dataGridViewCellStyle51.BackColor = System.Drawing.Color.LightCyan;
            this.dataGridViewTextBoxColumn5.DefaultCellStyle = dataGridViewCellStyle51;
            this.dataGridViewTextBoxColumn5.HeaderText = "指定社團";
            this.dataGridViewTextBoxColumn5.Name = "dataGridViewTextBoxColumn5";
            this.dataGridViewTextBoxColumn5.ReadOnly = true;
            this.dataGridViewTextBoxColumn5.Width = 120;
            // 
            // dataGridViewTextBoxColumn6
            // 
            this.dataGridViewTextBoxColumn6.HeaderText = "指定社團";
            this.dataGridViewTextBoxColumn6.Name = "dataGridViewTextBoxColumn6";
            this.dataGridViewTextBoxColumn6.ReadOnly = true;
            this.dataGridViewTextBoxColumn6.Width = 120;
            // 
            // colGrade_year
            // 
            dataGridViewCellStyle42.BackColor = System.Drawing.Color.LightCyan;
            this.colGrade_year.DefaultCellStyle = dataGridViewCellStyle42;
            this.colGrade_year.HeaderText = "年級";
            this.colGrade_year.Name = "colGrade_year";
            this.colGrade_year.ReadOnly = true;
            this.colGrade_year.Width = 60;
            // 
            // colClass
            // 
            dataGridViewCellStyle43.BackColor = System.Drawing.Color.LightCyan;
            this.colClass.DefaultCellStyle = dataGridViewCellStyle43;
            this.colClass.HeaderText = "班級";
            this.colClass.Name = "colClass";
            this.colClass.ReadOnly = true;
            this.colClass.Width = 80;
            // 
            // colSeat_no
            // 
            dataGridViewCellStyle44.BackColor = System.Drawing.Color.LightCyan;
            this.colSeat_no.DefaultCellStyle = dataGridViewCellStyle44;
            this.colSeat_no.HeaderText = "座號";
            this.colSeat_no.Name = "colSeat_no";
            this.colSeat_no.ReadOnly = true;
            this.colSeat_no.Width = 60;
            // 
            // colName
            // 
            dataGridViewCellStyle45.BackColor = System.Drawing.Color.LightCyan;
            this.colName.DefaultCellStyle = dataGridViewCellStyle45;
            this.colName.HeaderText = "姓名";
            this.colName.Name = "colName";
            this.colName.ReadOnly = true;
            this.colName.Width = 80;
            // 
            // colStudentNumber
            // 
            dataGridViewCellStyle46.BackColor = System.Drawing.Color.LightCyan;
            this.colStudentNumber.DefaultCellStyle = dataGridViewCellStyle46;
            this.colStudentNumber.HeaderText = "學號";
            this.colStudentNumber.Name = "colStudentNumber";
            this.colStudentNumber.ReadOnly = true;
            this.colStudentNumber.Width = 80;
            // 
            // colSelectClub
            // 
            this.colSelectClub.HeaderText = "指定社團";
            this.colSelectClub.Name = "colSelectClub";
            this.colSelectClub.ReadOnly = true;
            this.colSelectClub.Width = 120;
            // 
            // CheckStudentIsNotInClub
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(753, 479);
            this.Controls.Add(this.itmPnlTimeName);
            this.Controls.Add(this.btnOutPut);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dataGridViewX1);
            this.Name = "CheckStudentIsNotInClub";
            this.Text = "未選社團學生";
            this.Load += new System.EventHandler(this.CheckStudentIsNotInClub_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).EndInit();
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dataGridViewX1;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn4;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn5;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGrade_year;
        private System.Windows.Forms.DataGridViewTextBoxColumn colClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSeat_no;
        private System.Windows.Forms.DataGridViewTextBoxColumn colName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStudentNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSelectClub;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn6;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 清除指定社團ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 加入待處理學生ToolStripMenuItem;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.ButtonX btnOutPut;
        private DevComponents.DotNetBar.ItemPanel itmPnlTimeName;
    }
}