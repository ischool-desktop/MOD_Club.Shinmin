namespace K12.Club.Shinmin
{
    partial class NewAddClub
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
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.intSchoolYear = new DevComponents.Editors.IntegerInput();
            this.intSemester = new DevComponents.Editors.IntegerInput();
            this.txtClubName = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.lbSchoolYear = new DevComponents.DotNetBar.LabelX();
            this.lbSemester = new DevComponents.DotNetBar.LabelX();
            this.lbClubName = new DevComponents.DotNetBar.LabelX();
            this.tbAboutClub = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.lbAboutClub = new DevComponents.DotNetBar.LabelX();
            this.groupPanel1 = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.listDepartment = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.cbGenderRestrict = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.comboItem3 = new DevComponents.Editors.ComboItem();
            this.comboItem1 = new DevComponents.Editors.ComboItem();
            this.comboItem2 = new DevComponents.Editors.ComboItem();
            this.tbLimit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.tbGrade3Limit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.tbGrade2Limit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.tbGrade1Limit = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.lbGenderRestrict = new DevComponents.DotNetBar.LabelX();
            this.lbDepartment = new DevComponents.DotNetBar.LabelX();
            this.lbLimit = new DevComponents.DotNetBar.LabelX();
            this.lbGrade3Limit = new DevComponents.DotNetBar.LabelX();
            this.Grade2Limit = new DevComponents.DotNetBar.LabelX();
            this.lbGrade1Limit = new DevComponents.DotNetBar.LabelX();
            this.cbTeacher = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbTeacher = new DevComponents.DotNetBar.LabelX();
            this.lbLocation = new DevComponents.DotNetBar.LabelX();
            this.cbLocation = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbHelp4 = new DevComponents.DotNetBar.LabelX();
            this.lbCategory = new DevComponents.DotNetBar.LabelX();
            this.cbCategory = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.tbClubNumber = new DevComponents.DotNetBar.Controls.TextBoxX();
            this.lbClubNumber = new DevComponents.DotNetBar.LabelX();
            this.cbTeacher2 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbTeacher2 = new DevComponents.DotNetBar.LabelX();
            this.cbTeacher3 = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.lbTeacher3 = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.intSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.intSemester)).BeginInit();
            this.groupPanel1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.AutoSize = true;
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Location = new System.Drawing.Point(413, 386);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 25);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 18;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(494, 386);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 19;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // intSchoolYear
            // 
            this.intSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.intSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.intSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.intSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.intSchoolYear.Location = new System.Drawing.Point(68, 10);
            this.intSchoolYear.MaxValue = 999;
            this.intSchoolYear.MinValue = 90;
            this.intSchoolYear.Name = "intSchoolYear";
            this.intSchoolYear.ShowUpDown = true;
            this.intSchoolYear.Size = new System.Drawing.Size(80, 25);
            this.intSchoolYear.TabIndex = 21;
            this.intSchoolYear.Value = 90;
            this.intSchoolYear.ValueChanged += new System.EventHandler(this.intSchoolYear_ValueChanged);
            // 
            // intSemester
            // 
            this.intSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.intSemester.BackgroundStyle.Class = "DateTimeInputBackground";
            this.intSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.intSemester.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.intSemester.Location = new System.Drawing.Point(220, 10);
            this.intSemester.MaxValue = 2;
            this.intSemester.MinValue = 1;
            this.intSemester.Name = "intSemester";
            this.intSemester.ShowUpDown = true;
            this.intSemester.Size = new System.Drawing.Size(80, 25);
            this.intSemester.TabIndex = 23;
            this.intSemester.Value = 1;
            this.intSemester.ValueChanged += new System.EventHandler(this.intSemester_ValueChanged);
            // 
            // txtClubName
            // 
            // 
            // 
            // 
            this.txtClubName.Border.Class = "TextBoxBorder";
            this.txtClubName.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.txtClubName.Location = new System.Drawing.Point(68, 42);
            this.txtClubName.Name = "txtClubName";
            this.txtClubName.Size = new System.Drawing.Size(232, 25);
            this.txtClubName.TabIndex = 1;
            // 
            // lbSchoolYear
            // 
            this.lbSchoolYear.AutoSize = true;
            this.lbSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbSchoolYear.BackgroundStyle.Class = "";
            this.lbSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbSchoolYear.Location = new System.Drawing.Point(13, 12);
            this.lbSchoolYear.Name = "lbSchoolYear";
            this.lbSchoolYear.Size = new System.Drawing.Size(47, 21);
            this.lbSchoolYear.TabIndex = 20;
            this.lbSchoolYear.Text = "學年度";
            // 
            // lbSemester
            // 
            this.lbSemester.AutoSize = true;
            this.lbSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbSemester.BackgroundStyle.Class = "";
            this.lbSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbSemester.Location = new System.Drawing.Point(171, 12);
            this.lbSemester.Name = "lbSemester";
            this.lbSemester.Size = new System.Drawing.Size(34, 21);
            this.lbSemester.TabIndex = 22;
            this.lbSemester.Text = "學期";
            // 
            // lbClubName
            // 
            this.lbClubName.AutoSize = true;
            this.lbClubName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbClubName.BackgroundStyle.Class = "";
            this.lbClubName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbClubName.Location = new System.Drawing.Point(13, 44);
            this.lbClubName.Name = "lbClubName";
            this.lbClubName.Size = new System.Drawing.Size(47, 21);
            this.lbClubName.TabIndex = 0;
            this.lbClubName.Text = "名　稱";
            // 
            // tbAboutClub
            // 
            // 
            // 
            // 
            this.tbAboutClub.Border.Class = "TextBoxBorder";
            this.tbAboutClub.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbAboutClub.Location = new System.Drawing.Point(68, 266);
            this.tbAboutClub.Multiline = true;
            this.tbAboutClub.Name = "tbAboutClub";
            this.tbAboutClub.Size = new System.Drawing.Size(232, 115);
            this.tbAboutClub.TabIndex = 15;
            // 
            // lbAboutClub
            // 
            this.lbAboutClub.AutoSize = true;
            this.lbAboutClub.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbAboutClub.BackgroundStyle.Class = "";
            this.lbAboutClub.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbAboutClub.Location = new System.Drawing.Point(13, 266);
            this.lbAboutClub.Name = "lbAboutClub";
            this.lbAboutClub.Size = new System.Drawing.Size(47, 21);
            this.lbAboutClub.TabIndex = 14;
            this.lbAboutClub.Text = "簡　介";
            // 
            // groupPanel1
            // 
            this.groupPanel1.BackColor = System.Drawing.Color.Transparent;
            this.groupPanel1.CanvasColor = System.Drawing.SystemColors.Control;
            this.groupPanel1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.groupPanel1.Controls.Add(this.listDepartment);
            this.groupPanel1.Controls.Add(this.cbGenderRestrict);
            this.groupPanel1.Controls.Add(this.tbLimit);
            this.groupPanel1.Controls.Add(this.tbGrade3Limit);
            this.groupPanel1.Controls.Add(this.tbGrade2Limit);
            this.groupPanel1.Controls.Add(this.tbGrade1Limit);
            this.groupPanel1.Controls.Add(this.lbGenderRestrict);
            this.groupPanel1.Controls.Add(this.lbDepartment);
            this.groupPanel1.Controls.Add(this.lbLimit);
            this.groupPanel1.Controls.Add(this.lbGrade3Limit);
            this.groupPanel1.Controls.Add(this.Grade2Limit);
            this.groupPanel1.Controls.Add(this.lbGrade1Limit);
            this.groupPanel1.Location = new System.Drawing.Point(319, 10);
            this.groupPanel1.Name = "groupPanel1";
            this.groupPanel1.Size = new System.Drawing.Size(250, 367);
            // 
            // 
            // 
            this.groupPanel1.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.groupPanel1.Style.BackColorGradientAngle = 90;
            this.groupPanel1.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.groupPanel1.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderBottomWidth = 1;
            this.groupPanel1.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.groupPanel1.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderLeftWidth = 1;
            this.groupPanel1.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderRightWidth = 1;
            this.groupPanel1.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.groupPanel1.Style.BorderTopWidth = 1;
            this.groupPanel1.Style.Class = "";
            this.groupPanel1.Style.CornerDiameter = 4;
            this.groupPanel1.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.groupPanel1.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.groupPanel1.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseDown.Class = "";
            this.groupPanel1.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.groupPanel1.StyleMouseOver.Class = "";
            this.groupPanel1.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.groupPanel1.TabIndex = 16;
            this.groupPanel1.Text = "選社限制";
            // 
            // listDepartment
            // 
            // 
            // 
            // 
            this.listDepartment.Border.Class = "ListViewBorder";
            this.listDepartment.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.listDepartment.CheckBoxes = true;
            this.listDepartment.Location = new System.Drawing.Point(9, 143);
            this.listDepartment.Name = "listDepartment";
            this.listDepartment.Size = new System.Drawing.Size(227, 191);
            this.listDepartment.TabIndex = 11;
            this.listDepartment.UseCompatibleStateImageBehavior = false;
            this.listDepartment.View = System.Windows.Forms.View.List;
            this.listDepartment.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.listDepartment_ItemCheck);
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
            this.cbGenderRestrict.Location = new System.Drawing.Point(68, 82);
            this.cbGenderRestrict.Name = "cbGenderRestrict";
            this.cbGenderRestrict.Size = new System.Drawing.Size(168, 25);
            this.cbGenderRestrict.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbGenderRestrict.TabIndex = 9;
            // 
            // comboItem1
            // 
            this.comboItem1.Text = "男";
            // 
            // comboItem2
            // 
            this.comboItem2.Text = "女";
            // 
            // tbLimit
            // 
            // 
            // 
            // 
            this.tbLimit.Border.Class = "TextBoxBorder";
            this.tbLimit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbLimit.Location = new System.Drawing.Point(187, 46);
            this.tbLimit.Name = "tbLimit";
            this.tbLimit.Size = new System.Drawing.Size(49, 25);
            this.tbLimit.TabIndex = 7;
            // 
            // tbGrade3Limit
            // 
            // 
            // 
            // 
            this.tbGrade3Limit.Border.Class = "TextBoxBorder";
            this.tbGrade3Limit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbGrade3Limit.Location = new System.Drawing.Point(68, 46);
            this.tbGrade3Limit.Name = "tbGrade3Limit";
            this.tbGrade3Limit.Size = new System.Drawing.Size(49, 25);
            this.tbGrade3Limit.TabIndex = 5;
            this.tbGrade3Limit.Text = "0";
            this.tbGrade3Limit.TextChanged += new System.EventHandler(this.tbGrade3Limit_TextChanged);
            // 
            // tbGrade2Limit
            // 
            // 
            // 
            // 
            this.tbGrade2Limit.Border.Class = "TextBoxBorder";
            this.tbGrade2Limit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbGrade2Limit.Location = new System.Drawing.Point(187, 13);
            this.tbGrade2Limit.Name = "tbGrade2Limit";
            this.tbGrade2Limit.Size = new System.Drawing.Size(49, 25);
            this.tbGrade2Limit.TabIndex = 3;
            this.tbGrade2Limit.TextChanged += new System.EventHandler(this.tbGrade2Limit_TextChanged);
            // 
            // tbGrade1Limit
            // 
            // 
            // 
            // 
            this.tbGrade1Limit.Border.Class = "TextBoxBorder";
            this.tbGrade1Limit.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbGrade1Limit.Location = new System.Drawing.Point(68, 10);
            this.tbGrade1Limit.Name = "tbGrade1Limit";
            this.tbGrade1Limit.Size = new System.Drawing.Size(49, 25);
            this.tbGrade1Limit.TabIndex = 1;
            this.tbGrade1Limit.TextChanged += new System.EventHandler(this.tbGrade1Limit_TextChanged);
            // 
            // lbGenderRestrict
            // 
            this.lbGenderRestrict.AutoSize = true;
            // 
            // 
            // 
            this.lbGenderRestrict.BackgroundStyle.Class = "";
            this.lbGenderRestrict.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbGenderRestrict.Location = new System.Drawing.Point(9, 84);
            this.lbGenderRestrict.Name = "lbGenderRestrict";
            this.lbGenderRestrict.Size = new System.Drawing.Size(57, 21);
            this.lbGenderRestrict.TabIndex = 8;
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
            this.lbDepartment.Location = new System.Drawing.Point(9, 115);
            this.lbDepartment.Name = "lbDepartment";
            this.lbDepartment.Size = new System.Drawing.Size(109, 21);
            this.lbDepartment.TabIndex = 10;
            this.lbDepartment.Text = "科別限制(複選)：";
            // 
            // lbLimit
            // 
            this.lbLimit.AutoSize = true;
            // 
            // 
            // 
            this.lbLimit.BackgroundStyle.Class = "";
            this.lbLimit.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbLimit.Location = new System.Drawing.Point(128, 48);
            this.lbLimit.Name = "lbLimit";
            this.lbLimit.Size = new System.Drawing.Size(60, 21);
            this.lbLimit.TabIndex = 6;
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
            this.lbGrade3Limit.Location = new System.Drawing.Point(9, 48);
            this.lbGrade3Limit.Name = "lbGrade3Limit";
            this.lbGrade3Limit.Size = new System.Drawing.Size(60, 21);
            this.lbGrade3Limit.TabIndex = 4;
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
            this.Grade2Limit.Location = new System.Drawing.Point(128, 14);
            this.Grade2Limit.Name = "Grade2Limit";
            this.Grade2Limit.Size = new System.Drawing.Size(60, 21);
            this.Grade2Limit.TabIndex = 2;
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
            this.lbGrade1Limit.Location = new System.Drawing.Point(9, 12);
            this.lbGrade1Limit.Name = "lbGrade1Limit";
            this.lbGrade1Limit.Size = new System.Drawing.Size(60, 21);
            this.lbGrade1Limit.TabIndex = 0;
            this.lbGrade1Limit.Text = "一  年  級";
            // 
            // cbTeacher
            // 
            this.cbTeacher.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbTeacher.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbTeacher.DisplayMember = "Text";
            this.cbTeacher.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbTeacher.FormattingEnabled = true;
            this.cbTeacher.ItemHeight = 19;
            this.cbTeacher.Location = new System.Drawing.Point(68, 170);
            this.cbTeacher.Name = "cbTeacher";
            this.cbTeacher.Size = new System.Drawing.Size(232, 25);
            this.cbTeacher.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbTeacher.TabIndex = 9;
            // 
            // lbTeacher
            // 
            this.lbTeacher.AutoSize = true;
            this.lbTeacher.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbTeacher.BackgroundStyle.Class = "";
            this.lbTeacher.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbTeacher.Location = new System.Drawing.Point(13, 172);
            this.lbTeacher.Name = "lbTeacher";
            this.lbTeacher.Size = new System.Drawing.Size(47, 21);
            this.lbTeacher.TabIndex = 8;
            this.lbTeacher.Text = "老師１";
            // 
            // lbLocation
            // 
            this.lbLocation.AutoSize = true;
            this.lbLocation.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbLocation.BackgroundStyle.Class = "";
            this.lbLocation.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbLocation.Location = new System.Drawing.Point(13, 108);
            this.lbLocation.Name = "lbLocation";
            this.lbLocation.Size = new System.Drawing.Size(47, 21);
            this.lbLocation.TabIndex = 4;
            this.lbLocation.Text = "場　地";
            // 
            // cbLocation
            // 
            this.cbLocation.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbLocation.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbLocation.DisplayMember = "Text";
            this.cbLocation.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbLocation.FormattingEnabled = true;
            this.cbLocation.ItemHeight = 19;
            this.cbLocation.Location = new System.Drawing.Point(68, 106);
            this.cbLocation.Name = "cbLocation";
            this.cbLocation.Size = new System.Drawing.Size(232, 25);
            this.cbLocation.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbLocation.TabIndex = 5;
            // 
            // lbHelp4
            // 
            this.lbHelp4.AutoSize = true;
            this.lbHelp4.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbHelp4.BackgroundStyle.Class = "";
            this.lbHelp4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbHelp4.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.lbHelp4.Location = new System.Drawing.Point(26, 390);
            this.lbHelp4.Name = "lbHelp4";
            this.lbHelp4.Size = new System.Drawing.Size(366, 21);
            this.lbHelp4.TabIndex = 17;
            this.lbHelp4.Text = "說明：\"人數上限\"欄位，具有自動統計效果，也可手動修改。";
            // 
            // lbCategory
            // 
            this.lbCategory.AutoSize = true;
            this.lbCategory.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbCategory.BackgroundStyle.Class = "";
            this.lbCategory.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbCategory.Location = new System.Drawing.Point(13, 140);
            this.lbCategory.Name = "lbCategory";
            this.lbCategory.Size = new System.Drawing.Size(47, 21);
            this.lbCategory.TabIndex = 6;
            this.lbCategory.Text = "類　型";
            // 
            // cbCategory
            // 
            this.cbCategory.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbCategory.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbCategory.DisplayMember = "Text";
            this.cbCategory.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbCategory.FormattingEnabled = true;
            this.cbCategory.ItemHeight = 19;
            this.cbCategory.Location = new System.Drawing.Point(68, 138);
            this.cbCategory.Name = "cbCategory";
            this.cbCategory.Size = new System.Drawing.Size(232, 25);
            this.cbCategory.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbCategory.TabIndex = 7;
            // 
            // tbClubNumber
            // 
            // 
            // 
            // 
            this.tbClubNumber.Border.Class = "TextBoxBorder";
            this.tbClubNumber.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.tbClubNumber.Location = new System.Drawing.Point(68, 74);
            this.tbClubNumber.Name = "tbClubNumber";
            this.tbClubNumber.Size = new System.Drawing.Size(232, 25);
            this.tbClubNumber.TabIndex = 3;
            // 
            // lbClubNumber
            // 
            this.lbClubNumber.AutoSize = true;
            this.lbClubNumber.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbClubNumber.BackgroundStyle.Class = "";
            this.lbClubNumber.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbClubNumber.Location = new System.Drawing.Point(13, 76);
            this.lbClubNumber.Name = "lbClubNumber";
            this.lbClubNumber.Size = new System.Drawing.Size(47, 21);
            this.lbClubNumber.TabIndex = 2;
            this.lbClubNumber.Text = "代　碼";
            // 
            // cbTeacher2
            // 
            this.cbTeacher2.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbTeacher2.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbTeacher2.DisplayMember = "Text";
            this.cbTeacher2.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbTeacher2.FormattingEnabled = true;
            this.cbTeacher2.ItemHeight = 19;
            this.cbTeacher2.Location = new System.Drawing.Point(68, 202);
            this.cbTeacher2.Name = "cbTeacher2";
            this.cbTeacher2.Size = new System.Drawing.Size(232, 25);
            this.cbTeacher2.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbTeacher2.TabIndex = 11;
            // 
            // lbTeacher2
            // 
            this.lbTeacher2.AutoSize = true;
            this.lbTeacher2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbTeacher2.BackgroundStyle.Class = "";
            this.lbTeacher2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbTeacher2.Location = new System.Drawing.Point(13, 204);
            this.lbTeacher2.Name = "lbTeacher2";
            this.lbTeacher2.Size = new System.Drawing.Size(47, 21);
            this.lbTeacher2.TabIndex = 10;
            this.lbTeacher2.Text = "老師２";
            // 
            // cbTeacher3
            // 
            this.cbTeacher3.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.cbTeacher3.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems;
            this.cbTeacher3.DisplayMember = "Text";
            this.cbTeacher3.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cbTeacher3.FormattingEnabled = true;
            this.cbTeacher3.ItemHeight = 19;
            this.cbTeacher3.Location = new System.Drawing.Point(68, 234);
            this.cbTeacher3.Name = "cbTeacher3";
            this.cbTeacher3.Size = new System.Drawing.Size(232, 25);
            this.cbTeacher3.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cbTeacher3.TabIndex = 13;
            // 
            // lbTeacher3
            // 
            this.lbTeacher3.AutoSize = true;
            this.lbTeacher3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lbTeacher3.BackgroundStyle.Class = "";
            this.lbTeacher3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbTeacher3.Location = new System.Drawing.Point(13, 236);
            this.lbTeacher3.Name = "lbTeacher3";
            this.lbTeacher3.Size = new System.Drawing.Size(47, 21);
            this.lbTeacher3.TabIndex = 12;
            this.lbTeacher3.Text = "老師３";
            // 
            // NewAddClub
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(582, 421);
            this.Controls.Add(this.cbTeacher3);
            this.Controls.Add(this.lbTeacher3);
            this.Controls.Add(this.cbTeacher2);
            this.Controls.Add(this.lbTeacher2);
            this.Controls.Add(this.tbClubNumber);
            this.Controls.Add(this.lbClubNumber);
            this.Controls.Add(this.lbCategory);
            this.Controls.Add(this.cbTeacher);
            this.Controls.Add(this.tbAboutClub);
            this.Controls.Add(this.intSemester);
            this.Controls.Add(this.intSchoolYear);
            this.Controls.Add(this.txtClubName);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.lbHelp4);
            this.Controls.Add(this.lbTeacher);
            this.Controls.Add(this.groupPanel1);
            this.Controls.Add(this.lbAboutClub);
            this.Controls.Add(this.lbLocation);
            this.Controls.Add(this.lbClubName);
            this.Controls.Add(this.lbSemester);
            this.Controls.Add(this.lbSchoolYear);
            this.Controls.Add(this.cbCategory);
            this.Controls.Add(this.cbLocation);
            this.Name = "NewAddClub";
            this.Text = "新增社團";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.NewAddClub_FormClosing);
            this.Load += new System.EventHandler(this.NewAddClub_Load);
            ((System.ComponentModel.ISupportInitialize)(this.intSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.intSemester)).EndInit();
            this.groupPanel1.ResumeLayout(false);
            this.groupPanel1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.Editors.IntegerInput intSchoolYear;
        private DevComponents.Editors.IntegerInput intSemester;
        private DevComponents.DotNetBar.Controls.TextBoxX txtClubName;
        private DevComponents.DotNetBar.LabelX lbSchoolYear;
        private DevComponents.DotNetBar.LabelX lbSemester;
        private DevComponents.DotNetBar.LabelX lbClubName;
        private DevComponents.DotNetBar.Controls.TextBoxX tbAboutClub;
        private DevComponents.DotNetBar.LabelX lbAboutClub;
        private DevComponents.DotNetBar.LabelX lbLocation;
        private DevComponents.DotNetBar.Controls.GroupPanel groupPanel1;
        private DevComponents.DotNetBar.LabelX lbLimit;
        private DevComponents.DotNetBar.LabelX lbGrade3Limit;
        private DevComponents.DotNetBar.LabelX Grade2Limit;
        private DevComponents.DotNetBar.LabelX lbGrade1Limit;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbTeacher;
        private DevComponents.DotNetBar.LabelX lbTeacher;
        private DevComponents.DotNetBar.Controls.TextBoxX tbGrade1Limit;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbGenderRestrict;
        private DevComponents.Editors.ComboItem comboItem3;
        private DevComponents.Editors.ComboItem comboItem1;
        private DevComponents.Editors.ComboItem comboItem2;
        private DevComponents.DotNetBar.LabelX lbGenderRestrict;
        private DevComponents.DotNetBar.LabelX lbDepartment;
        private DevComponents.DotNetBar.Controls.TextBoxX tbLimit;
        private DevComponents.DotNetBar.Controls.TextBoxX tbGrade3Limit;
        private DevComponents.DotNetBar.Controls.TextBoxX tbGrade2Limit;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbLocation;
        private DevComponents.DotNetBar.Controls.ListViewEx listDepartment;
        private DevComponents.DotNetBar.LabelX lbHelp4;
        private DevComponents.DotNetBar.LabelX lbCategory;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbCategory;
        private DevComponents.DotNetBar.Controls.TextBoxX tbClubNumber;
        private DevComponents.DotNetBar.LabelX lbClubNumber;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbTeacher2;
        private DevComponents.DotNetBar.LabelX lbTeacher2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cbTeacher3;
        private DevComponents.DotNetBar.LabelX lbTeacher3;
    }
}