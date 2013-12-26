namespace K12.Club.Shinmin
{
    partial class OpenClubJoinDateTime
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgvTimes = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.chGradeYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chStartTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.chEndTime = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblSemester = new DevComponents.DotNetBar.LabelX();
            this.btnClose = new DevComponents.DotNetBar.ButtonX();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewTextBoxColumn3 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgvTimes)).BeginInit();
            this.SuspendLayout();
            // 
            // dgvTimes
            // 
            this.dgvTimes.AllowUserToAddRows = false;
            this.dgvTimes.AllowUserToDeleteRows = false;
            this.dgvTimes.BackgroundColor = System.Drawing.SystemColors.Window;
            this.dgvTimes.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvTimes.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.chGradeYear,
            this.chStartTime,
            this.chEndTime});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvTimes.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvTimes.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvTimes.Location = new System.Drawing.Point(13, 38);
            this.dgvTimes.Name = "dgvTimes";
            this.dgvTimes.RowTemplate.Height = 24;
            this.dgvTimes.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvTimes.Size = new System.Drawing.Size(462, 148);
            this.dgvTimes.TabIndex = 1;
            this.dgvTimes.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTimes_CellClick);
            this.dgvTimes.CellEndEdit += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTimes_CellEndEdit);
            this.dgvTimes.CellValidating += new System.Windows.Forms.DataGridViewCellValidatingEventHandler(this.dgvTimes_CellValidating);
            this.dgvTimes.RowValidated += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvTimes_RowValidated);
            this.dgvTimes.RowValidating += new System.Windows.Forms.DataGridViewCellCancelEventHandler(this.dgvTimes_RowValidating);
            // 
            // chGradeYear
            // 
            dataGridViewCellStyle1.BackColor = System.Drawing.Color.LightCyan;
            this.chGradeYear.DefaultCellStyle = dataGridViewCellStyle1;
            this.chGradeYear.HeaderText = "年級";
            this.chGradeYear.Name = "chGradeYear";
            this.chGradeYear.ReadOnly = true;
            // 
            // chStartTime
            // 
            this.chStartTime.HeaderText = "開始時間";
            this.chStartTime.Name = "chStartTime";
            this.chStartTime.Width = 150;
            // 
            // chEndTime
            // 
            this.chEndTime.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.chEndTime.HeaderText = "截止時間";
            this.chEndTime.Name = "chEndTime";
            // 
            // lblSemester
            // 
            this.lblSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblSemester.BackgroundStyle.Class = "";
            this.lblSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblSemester.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.lblSemester.Location = new System.Drawing.Point(13, 9);
            this.lblSemester.Name = "lblSemester";
            this.lblSemester.Size = new System.Drawing.Size(142, 23);
            this.lblSemester.TabIndex = 0;
            this.lblSemester.Text = "97學年度　第１學期";
            // 
            // btnClose
            // 
            this.btnClose.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClose.BackColor = System.Drawing.Color.Transparent;
            this.btnClose.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClose.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnClose.Location = new System.Drawing.Point(400, 208);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(75, 23);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "離開";
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnSave.Font = new System.Drawing.Font("微軟正黑體", 9.75F);
            this.btnSave.Location = new System.Drawing.Point(319, 208);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.TabIndex = 3;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
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
            this.labelX1.Location = new System.Drawing.Point(12, 192);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(299, 39);
            this.labelX1.TabIndex = 2;
            this.labelX1.Text = "說明：時間格式範例（2012/7/23 13:00）\r\n輸入（7/23）會自動替換為（2012/7/23 00:00）";
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "年級";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            // 
            // dataGridViewTextBoxColumn2
            // 
            this.dataGridViewTextBoxColumn2.HeaderText = "開始時間";
            this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
            this.dataGridViewTextBoxColumn2.Width = 150;
            // 
            // dataGridViewTextBoxColumn3
            // 
            this.dataGridViewTextBoxColumn3.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.dataGridViewTextBoxColumn3.HeaderText = "截止時間";
            this.dataGridViewTextBoxColumn3.Name = "dataGridViewTextBoxColumn3";
            // 
            // OpenClubJoinDateTime
            // 
            this.AcceptButton = this.btnClose;
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnSave;
            this.ClientSize = new System.Drawing.Size(489, 243);
            this.Controls.Add(this.dgvTimes);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.lblSemester);
            this.Name = "OpenClubJoinDateTime";
            this.Text = "開放選社時間";
            this.Load += new System.EventHandler(this.DailyLifeInputControl_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgvTimes)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgvTimes;
        private DevComponents.DotNetBar.LabelX lblSemester;
        private DevComponents.DotNetBar.ButtonX btnClose;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn3;
        private DevComponents.DotNetBar.LabelX labelX1;
        private System.Windows.Forms.DataGridViewTextBoxColumn chGradeYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn chStartTime;
        private System.Windows.Forms.DataGridViewTextBoxColumn chEndTime;
    }
}