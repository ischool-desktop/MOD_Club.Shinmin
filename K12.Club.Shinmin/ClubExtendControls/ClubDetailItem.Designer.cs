namespace K12.Club.Shinmin
{
    partial class ClubDetailItem
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器
        /// 修改這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.txtClubName = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.txtAbout = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.cbLocation = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.cbTeacher1 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbClubName = new DevComponents.DotNetBar.LabelX();
            this.lbTeacher1 = new DevComponents.DotNetBar.LabelX();
            this.lbLocation = new DevComponents.DotNetBar.LabelX();
            this.lbAbout = new DevComponents.DotNetBar.LabelX();
            this.lbSchoolYear = new DevComponents.DotNetBar.LabelX();
            this.lbCategory = new DevComponents.DotNetBar.LabelX();
            this.lbCLUBNumber = new DevComponents.DotNetBar.LabelX();
            this.tbCLUBNumber = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.cbCategory = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbTeacher2 = new DevComponents.DotNetBar.LabelX();
            this.cbTeacher2 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbTeacher3 = new DevComponents.DotNetBar.LabelX();
            this.cbTeacher3 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // txtClubName
            // 
            // 
            // 
            // 
            this.txtClubName.Border.Class = "TextBoxBorder";
            this.txtClubName.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtClubName.Location = new System.Drawing.Point(67, 52);
            this.txtClubName.Margin = new System.Windows.Forms.Padding(4);
            this.txtClubName.Name = "txtClubName";
            this.txtClubName.Size = new System.Drawing.Size(187, 25);
            this.txtClubName.TabIndex = 2;
            // 
            // txtAbout
            // 
            // 
            // 
            // 
            this.txtAbout.Border.Class = "TextBoxBorder";
            this.txtAbout.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtAbout.Location = new System.Drawing.Point(67, 157);
            this.txtAbout.Margin = new System.Windows.Forms.Padding(4);
            this.txtAbout.Multiline = true;
            this.txtAbout.Name = "txtAbout";
            this.txtAbout.Size = new System.Drawing.Size(436, 67);
            this.txtAbout.TabIndex = 16;
            // 
            // cbLocation
            // 
            this.cbLocation.DisplayMember = "Text";
            this.cbLocation.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbLocation.FormattingEnabled = true;
            this.cbLocation.ItemHeight = 19;
            this.cbLocation.Location = new System.Drawing.Point(67, 87);
            this.cbLocation.Name = "cbLocation";
            this.cbLocation.Size = new System.Drawing.Size(187, 25);
            this.cbLocation.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbLocation.TabIndex = 4;
            // 
            // cbTeacher1
            // 
            this.cbTeacher1.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbTeacher1.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbTeacher1.DisplayMember = "Text";
            this.cbTeacher1.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbTeacher1.FormattingEnabled = true;
            this.cbTeacher1.ItemHeight = 19;
            this.cbTeacher1.Location = new System.Drawing.Point(315, 52);
            this.cbTeacher1.Name = "cbTeacher1";
            this.cbTeacher1.Size = new System.Drawing.Size(188, 25);
            this.cbTeacher1.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbTeacher1.TabIndex = 10;
            // 
            // lbClubName
            // 
            this.lbClubName.AutoSize = true;
            // 
            // 
            // 
            this.lbClubName.BackgroundStyle.Class = "";
            this.lbClubName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbClubName.Location = new System.Drawing.Point(28, 54);
            this.lbClubName.Name = "lbClubName";
            this.lbClubName.Size = new System.Drawing.Size(34, 21);
            this.lbClubName.TabIndex = 1;
            this.lbClubName.Text = "名稱";
            // 
            // lbTeacher1
            // 
            this.lbTeacher1.AutoSize = true;
            // 
            // 
            // 
            this.lbTeacher1.BackgroundStyle.Class = "";
            this.lbTeacher1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbTeacher1.Location = new System.Drawing.Point(268, 54);
            this.lbTeacher1.Name = "lbTeacher1";
            this.lbTeacher1.Size = new System.Drawing.Size(41, 21);
            this.lbTeacher1.TabIndex = 9;
            this.lbTeacher1.Text = "老師1";
            // 
            // lbLocation
            // 
            this.lbLocation.AutoSize = true;
            // 
            // 
            // 
            this.lbLocation.BackgroundStyle.Class = "";
            this.lbLocation.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbLocation.Location = new System.Drawing.Point(28, 89);
            this.lbLocation.Name = "lbLocation";
            this.lbLocation.Size = new System.Drawing.Size(34, 21);
            this.lbLocation.TabIndex = 3;
            this.lbLocation.Text = "場地";
            // 
            // lbAbout
            // 
            this.lbAbout.AutoSize = true;
            // 
            // 
            // 
            this.lbAbout.BackgroundStyle.Class = "";
            this.lbAbout.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbAbout.Location = new System.Drawing.Point(28, 159);
            this.lbAbout.Name = "lbAbout";
            this.lbAbout.Size = new System.Drawing.Size(34, 21);
            this.lbAbout.TabIndex = 15;
            this.lbAbout.Text = "簡介";
            // 
            // lbSchoolYear
            // 
            this.lbSchoolYear.AutoSize = true;
            // 
            // 
            // 
            this.lbSchoolYear.BackgroundStyle.Class = "";
            this.lbSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbSchoolYear.Font = new System.Drawing.Font("微軟正黑體", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.lbSchoolYear.Location = new System.Drawing.Point(28, 16);
            this.lbSchoolYear.Name = "lbSchoolYear";
            this.lbSchoolYear.Size = new System.Drawing.Size(105, 26);
            this.lbSchoolYear.TabIndex = 0;
            this.lbSchoolYear.Text = "學年度 / 學期";
            // 
            // lbCategory
            // 
            this.lbCategory.AutoSize = true;
            // 
            // 
            // 
            this.lbCategory.BackgroundStyle.Class = "";
            this.lbCategory.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbCategory.Location = new System.Drawing.Point(28, 124);
            this.lbCategory.Name = "lbCategory";
            this.lbCategory.Size = new System.Drawing.Size(34, 21);
            this.lbCategory.TabIndex = 5;
            this.lbCategory.Text = "類型";
            // 
            // lbCLUBNumber
            // 
            this.lbCLUBNumber.AutoSize = true;
            // 
            // 
            // 
            this.lbCLUBNumber.BackgroundStyle.Class = "";
            this.lbCLUBNumber.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbCLUBNumber.Location = new System.Drawing.Point(268, 19);
            this.lbCLUBNumber.Name = "lbCLUBNumber";
            this.lbCLUBNumber.Size = new System.Drawing.Size(34, 21);
            this.lbCLUBNumber.TabIndex = 7;
            this.lbCLUBNumber.Text = "代碼";
            // 
            // tbCLUBNumber
            // 
            // 
            // 
            // 
            this.tbCLUBNumber.Border.Class = "TextBoxBorder";
            this.tbCLUBNumber.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbCLUBNumber.Location = new System.Drawing.Point(315, 17);
            this.tbCLUBNumber.Margin = new System.Windows.Forms.Padding(4);
            this.tbCLUBNumber.Name = "tbCLUBNumber";
            this.tbCLUBNumber.Size = new System.Drawing.Size(188, 25);
            this.tbCLUBNumber.TabIndex = 8;
            // 
            // cbCategory
            // 
            this.cbCategory.DisplayMember = "Text";
            this.cbCategory.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbCategory.FormattingEnabled = true;
            this.cbCategory.ItemHeight = 19;
            this.cbCategory.Location = new System.Drawing.Point(67, 122);
            this.cbCategory.Name = "cbCategory";
            this.cbCategory.Size = new System.Drawing.Size(187, 25);
            this.cbCategory.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbCategory.TabIndex = 6;
            // 
            // lbTeacher2
            // 
            this.lbTeacher2.AutoSize = true;
            // 
            // 
            // 
            this.lbTeacher2.BackgroundStyle.Class = "";
            this.lbTeacher2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbTeacher2.Location = new System.Drawing.Point(268, 89);
            this.lbTeacher2.Name = "lbTeacher2";
            this.lbTeacher2.Size = new System.Drawing.Size(41, 21);
            this.lbTeacher2.TabIndex = 11;
            this.lbTeacher2.Text = "老師2";
            // 
            // cbTeacher2
            // 
            this.cbTeacher2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbTeacher2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbTeacher2.DisplayMember = "Text";
            this.cbTeacher2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbTeacher2.FormattingEnabled = true;
            this.cbTeacher2.ItemHeight = 19;
            this.cbTeacher2.Location = new System.Drawing.Point(315, 87);
            this.cbTeacher2.Name = "cbTeacher2";
            this.cbTeacher2.Size = new System.Drawing.Size(188, 25);
            this.cbTeacher2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbTeacher2.TabIndex = 12;
            // 
            // lbTeacher3
            // 
            this.lbTeacher3.AutoSize = true;
            // 
            // 
            // 
            this.lbTeacher3.BackgroundStyle.Class = "";
            this.lbTeacher3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbTeacher3.Location = new System.Drawing.Point(268, 124);
            this.lbTeacher3.Name = "lbTeacher3";
            this.lbTeacher3.Size = new System.Drawing.Size(41, 21);
            this.lbTeacher3.TabIndex = 13;
            this.lbTeacher3.Text = "老師3";
            // 
            // cbTeacher3
            // 
            this.cbTeacher3.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbTeacher3.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbTeacher3.DisplayMember = "Text";
            this.cbTeacher3.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbTeacher3.FormattingEnabled = true;
            this.cbTeacher3.ItemHeight = 19;
            this.cbTeacher3.Location = new System.Drawing.Point(315, 122);
            this.cbTeacher3.Name = "cbTeacher3";
            this.cbTeacher3.Size = new System.Drawing.Size(188, 25);
            this.cbTeacher3.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbTeacher3.TabIndex = 14;
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(505, 54);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(42, 21);
            this.labelX1.TabIndex = 17;
            this.labelX1.Text = "(評分)";
            // 
            // ClubDetailItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.cbCategory);
            this.Controls.Add(this.txtAbout);
            this.Controls.Add(this.txtClubName);
            this.Controls.Add(this.cbLocation);
            this.Controls.Add(this.cbTeacher3);
            this.Controls.Add(this.cbTeacher2);
            this.Controls.Add(this.tbCLUBNumber);
            this.Controls.Add(this.cbTeacher1);
            this.Controls.Add(this.lbTeacher3);
            this.Controls.Add(this.lbTeacher2);
            this.Controls.Add(this.lbCLUBNumber);
            this.Controls.Add(this.lbCategory);
            this.Controls.Add(this.lbSchoolYear);
            this.Controls.Add(this.lbAbout);
            this.Controls.Add(this.lbLocation);
            this.Controls.Add(this.lbTeacher1);
            this.Controls.Add(this.lbClubName);
            this.Controls.Add(this.labelX1);
            this.Name = "ClubDetailItem";
            this.Size = new System.Drawing.Size(550, 240);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal DevComponents.DotNetBar.Controls.TextBoxX txtClubName;
        internal DevComponents.DotNetBar.Controls.TextBoxX txtAbout;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbLocation;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbTeacher1;
        private DevComponents.DotNetBar.LabelX lbClubName;
        private DevComponents.DotNetBar.LabelX lbTeacher1;
        private DevComponents.DotNetBar.LabelX lbLocation;
        private DevComponents.DotNetBar.LabelX lbAbout;
        private DevComponents.DotNetBar.LabelX lbSchoolYear;
        private DevComponents.DotNetBar.LabelX lbCategory;
        private DevComponents.DotNetBar.LabelX lbCLUBNumber;
        internal DevComponents.DotNetBar.Controls.TextBoxX tbCLUBNumber;
        public DevComponents.DotNetBar.Controls.ComboBoxEx cbCategory;
        private DevComponents.DotNetBar.LabelX lbTeacher2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbTeacher2;
        private DevComponents.DotNetBar.LabelX lbTeacher3;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbTeacher3;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}
