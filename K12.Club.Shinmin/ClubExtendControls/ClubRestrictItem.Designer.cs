namespace K12.Club.Shinmin
{
    partial class ClubRestrictItem
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
            this.cbGenderRestrict = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comboItem3 = new DevComponents.Editors.ComboItem();
            this.comboItem1 = new DevComponents.Editors.ComboItem();
            this.comboItem2 = new DevComponents.Editors.ComboItem();
            this.lbGenderRestrict = new DevComponents.DotNetBar.LabelX();
            this.lbDepartment = new DevComponents.DotNetBar.LabelX();
            this.tbLimit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.tbGrade3Limit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.tbGrade2Limit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.tbGrade1Limit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.lbLimit = new DevComponents.DotNetBar.LabelX();
            this.lbGrade3Limit = new DevComponents.DotNetBar.LabelX();
            this.Grade2Limit = new DevComponents.DotNetBar.LabelX();
            this.lbGrade1Limit = new DevComponents.DotNetBar.LabelX();
            this.listDepartment = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.SuspendLayout();
            // 
            // cbGenderRestrict
            // 
            this.cbGenderRestrict.DisplayMember = "Text";
            this.cbGenderRestrict.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbGenderRestrict.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbGenderRestrict.FormattingEnabled = true;
            this.cbGenderRestrict.ItemHeight = 19;
            this.cbGenderRestrict.Items.AddRange(new object[] {
            this.comboItem3,
            this.comboItem1,
            this.comboItem2});
            this.cbGenderRestrict.Location = new System.Drawing.Point(93, 18);
            this.cbGenderRestrict.Name = "cbGenderRestrict";
            this.cbGenderRestrict.Size = new System.Drawing.Size(104, 25);
            this.cbGenderRestrict.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbGenderRestrict.TabIndex = 1;
            // 
            // comboItem1
            // 
            this.comboItem1.Text = "男";
            // 
            // comboItem2
            // 
            this.comboItem2.Text = "女";
            // 
            // lbGenderRestrict
            // 
            this.lbGenderRestrict.AutoSize = true;
            // 
            // 
            // 
            this.lbGenderRestrict.BackgroundStyle.Class = "";
            this.lbGenderRestrict.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbGenderRestrict.Location = new System.Drawing.Point(26, 20);
            this.lbGenderRestrict.Name = "lbGenderRestrict";
            this.lbGenderRestrict.Size = new System.Drawing.Size(57, 21);
            this.lbGenderRestrict.TabIndex = 0;
            this.lbGenderRestrict.Text = "性 　  別";
            // 
            // lbDepartment
            // 
            this.lbDepartment.AutoSize = true;
            // 
            // 
            // 
            this.lbDepartment.BackgroundStyle.Class = "";
            this.lbDepartment.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbDepartment.Location = new System.Drawing.Point(200, 20);
            this.lbDepartment.Name = "lbDepartment";
            this.lbDepartment.Size = new System.Drawing.Size(74, 21);
            this.lbDepartment.TabIndex = 10;
            this.lbDepartment.Text = "科別限制：";
            // 
            // tbLimit
            // 
            // 
            // 
            // 
            this.tbLimit.Border.Class = "TextBoxBorder";
            this.tbLimit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbLimit.Location = new System.Drawing.Point(93, 157);
            this.tbLimit.Name = "tbLimit";
            this.tbLimit.Size = new System.Drawing.Size(167, 25);
            this.tbLimit.TabIndex = 9;
            // 
            // tbGrade3Limit
            // 
            // 
            // 
            // 
            this.tbGrade3Limit.Border.Class = "TextBoxBorder";
            this.tbGrade3Limit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbGrade3Limit.Location = new System.Drawing.Point(94, 122);
            this.tbGrade3Limit.Name = "tbGrade3Limit";
            this.tbGrade3Limit.Size = new System.Drawing.Size(166, 25);
            this.tbGrade3Limit.TabIndex = 7;
            this.tbGrade3Limit.TextChanged += new System.EventHandler(this.tbGrade3Limit_TextChanged);
            // 
            // tbGrade2Limit
            // 
            // 
            // 
            // 
            this.tbGrade2Limit.Border.Class = "TextBoxBorder";
            this.tbGrade2Limit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbGrade2Limit.Location = new System.Drawing.Point(94, 88);
            this.tbGrade2Limit.Name = "tbGrade2Limit";
            this.tbGrade2Limit.Size = new System.Drawing.Size(166, 25);
            this.tbGrade2Limit.TabIndex = 5;
            this.tbGrade2Limit.TextChanged += new System.EventHandler(this.tbGrade2Limit_TextChanged);
            // 
            // tbGrade1Limit
            // 
            // 
            // 
            // 
            this.tbGrade1Limit.Border.Class = "TextBoxBorder";
            this.tbGrade1Limit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbGrade1Limit.Location = new System.Drawing.Point(94, 53);
            this.tbGrade1Limit.Name = "tbGrade1Limit";
            this.tbGrade1Limit.Size = new System.Drawing.Size(166, 25);
            this.tbGrade1Limit.TabIndex = 3;
            this.tbGrade1Limit.TextChanged += new System.EventHandler(this.tbGrade1Limit_TextChanged);
            // 
            // lbLimit
            // 
            this.lbLimit.AutoSize = true;
            // 
            // 
            // 
            this.lbLimit.BackgroundStyle.Class = "";
            this.lbLimit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbLimit.Location = new System.Drawing.Point(24, 159);
            this.lbLimit.Name = "lbLimit";
            this.lbLimit.Size = new System.Drawing.Size(60, 21);
            this.lbLimit.TabIndex = 8;
            this.lbLimit.Text = "人數上限";
            // 
            // lbGrade3Limit
            // 
            this.lbGrade3Limit.AutoSize = true;
            // 
            // 
            // 
            this.lbGrade3Limit.BackgroundStyle.Class = "";
            this.lbGrade3Limit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbGrade3Limit.Location = new System.Drawing.Point(24, 124);
            this.lbGrade3Limit.Name = "lbGrade3Limit";
            this.lbGrade3Limit.Size = new System.Drawing.Size(60, 21);
            this.lbGrade3Limit.TabIndex = 6;
            this.lbGrade3Limit.Text = "三  年  級";
            // 
            // Grade2Limit
            // 
            this.Grade2Limit.AutoSize = true;
            // 
            // 
            // 
            this.Grade2Limit.BackgroundStyle.Class = "";
            this.Grade2Limit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.Grade2Limit.Location = new System.Drawing.Point(24, 90);
            this.Grade2Limit.Name = "Grade2Limit";
            this.Grade2Limit.Size = new System.Drawing.Size(60, 21);
            this.Grade2Limit.TabIndex = 4;
            this.Grade2Limit.Text = "二  年  級";
            // 
            // lbGrade1Limit
            // 
            this.lbGrade1Limit.AutoSize = true;
            // 
            // 
            // 
            this.lbGrade1Limit.BackgroundStyle.Class = "";
            this.lbGrade1Limit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbGrade1Limit.Location = new System.Drawing.Point(24, 55);
            this.lbGrade1Limit.Name = "lbGrade1Limit";
            this.lbGrade1Limit.Size = new System.Drawing.Size(60, 21);
            this.lbGrade1Limit.TabIndex = 2;
            this.lbGrade1Limit.Text = "一  年  級";
            // 
            // listDepartment
            // 
            // 
            // 
            // 
            this.listDepartment.Border.Class = "ListViewBorder";
            this.listDepartment.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.listDepartment.CheckBoxes = true;
            this.listDepartment.Location = new System.Drawing.Point(276, 20);
            this.listDepartment.Name = "listDepartment";
            this.listDepartment.Size = new System.Drawing.Size(251, 162);
            this.listDepartment.TabIndex = 11;
            this.listDepartment.UseCompatibleStateImageBehavior = false;
            this.listDepartment.View = System.Windows.Forms.View.List;
            this.listDepartment.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.listDepartment_ItemChecked);
            // 
            // ClubRestrictItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.listDepartment);
            this.Controls.Add(this.cbGenderRestrict);
            this.Controls.Add(this.tbGrade1Limit);
            this.Controls.Add(this.tbLimit);
            this.Controls.Add(this.tbGrade3Limit);
            this.Controls.Add(this.tbGrade2Limit);
            this.Controls.Add(this.lbGenderRestrict);
            this.Controls.Add(this.lbGrade1Limit);
            this.Controls.Add(this.lbDepartment);
            this.Controls.Add(this.Grade2Limit);
            this.Controls.Add(this.lbGrade3Limit);
            this.Controls.Add(this.lbLimit);
            this.Name = "ClubRestrictItem";
            this.Size = new System.Drawing.Size(550, 200);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cbGenderRestrict;
        private DevComponents.Editors.ComboItem comboItem3;
        private DevComponents.Editors.ComboItem comboItem1;
        private DevComponents.Editors.ComboItem comboItem2;
        private DevComponents.DotNetBar.LabelX lbGenderRestrict;
        private DevComponents.DotNetBar.LabelX lbDepartment;
        private DevComponents.DotNetBar.Controls.TextBoxX tbLimit;
        private DevComponents.DotNetBar.Controls.TextBoxX tbGrade3Limit;
        private DevComponents.DotNetBar.Controls.TextBoxX tbGrade2Limit;
        private DevComponents.DotNetBar.Controls.TextBoxX tbGrade1Limit;
        private DevComponents.DotNetBar.LabelX lbLimit;
        private DevComponents.DotNetBar.LabelX lbGrade3Limit;
        private DevComponents.DotNetBar.LabelX Grade2Limit;
        private DevComponents.DotNetBar.LabelX lbGrade1Limit;
        private DevComponents.DotNetBar.Controls.ListViewEx listDepartment;

    }
}
