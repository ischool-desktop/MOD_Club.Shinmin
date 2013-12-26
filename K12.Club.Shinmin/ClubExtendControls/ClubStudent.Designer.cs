namespace K12.Club.Shinmin
{
    partial class ClubStudent
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
            this.components = new System.ComponentModel.Container();
            this.btnClearStudent = new DevComponents.DotNetBar.ButtonX();
            this.chClass_ = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chSeatNo_ = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chName_ = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chGender_ = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chSNum_ = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnInserStudent = new DevComponents.DotNetBar.ButtonX();
            this.buttonItem2 = new DevComponents.DotNetBar.ButtonItem();
            this.lbCourseCount = new DevComponents.DotNetBar.LabelX();
            this.listViewEx1 = new K12.Club.Shinmin.ListViewEX();
            this.chGen_ = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chStatus = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.chLock = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripMenuItem();
            this.清空學生待處理ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.鎖定學生選社ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.移除選擇學生ToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.移除選擇學生ToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnClearStudent
            // 
            this.btnClearStudent.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnClearStudent.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnClearStudent.Location = new System.Drawing.Point(19, 210);
            this.btnClearStudent.Name = "btnClearStudent";
            this.btnClearStudent.Size = new System.Drawing.Size(100, 23);
            this.btnClearStudent.TabIndex = 1;
            this.btnClearStudent.Text = "移除選擇學生";
            this.btnClearStudent.Click += new System.EventHandler(this.btnClearStudent_Click);
            // 
            // chClass_
            // 
            this.chClass_.Text = "班級";
            this.chClass_.Width = 80;
            // 
            // chSeatNo_
            // 
            this.chSeatNo_.Text = "座號";
            // 
            // chName_
            // 
            this.chName_.Text = "姓名";
            this.chName_.Width = 75;
            // 
            // chGender_
            // 
            this.chGender_.Text = "性別";
            // 
            // chSNum_
            // 
            this.chSNum_.Text = "學號";
            this.chSNum_.Width = 75;
            // 
            // btnInserStudent
            // 
            this.btnInserStudent.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnInserStudent.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnInserStudent.Location = new System.Drawing.Point(127, 210);
            this.btnInserStudent.Name = "btnInserStudent";
            this.btnInserStudent.Size = new System.Drawing.Size(195, 23);
            this.btnInserStudent.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.buttonItem2});
            this.btnInserStudent.TabIndex = 2;
            this.btnInserStudent.Text = "由待處理(學生)加入社員";
            this.btnInserStudent.PopupOpen += new System.EventHandler(this.btnInserStudent_PopupOpen);
            this.btnInserStudent.Click += new System.EventHandler(this.btnInserStudent_Click);
            // 
            // buttonItem2
            // 
            this.buttonItem2.GlobalItem = false;
            this.buttonItem2.Name = "buttonItem2";
            this.buttonItem2.Text = "buttonItem2";
            // 
            // lbCourseCount
            // 
            this.lbCourseCount.AutoSize = true;
            // 
            // 
            // 
            this.lbCourseCount.BackgroundStyle.Class = "";
            this.lbCourseCount.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lbCourseCount.Location = new System.Drawing.Point(338, 211);
            this.lbCourseCount.Name = "lbCourseCount";
            this.lbCourseCount.Size = new System.Drawing.Size(160, 21);
            this.lbCourseCount.TabIndex = 3;
            this.lbCourseCount.Text = "有效社員/所有狀態社員：";
            // 
            // listViewEx1
            // 
            // 
            // 
            // 
            this.listViewEx1.Border.Class = "ListViewBorder";
            this.listViewEx1.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.listViewEx1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chGen_,
            this.chClass_,
            this.chSeatNo_,
            this.chName_,
            this.chGender_,
            this.chSNum_,
            this.chStatus,
            this.chLock});
            this.listViewEx1.ContextMenuStrip = this.contextMenuStrip1;
            this.listViewEx1.FullRowSelect = true;
            this.listViewEx1.HideSelection = false;
            this.listViewEx1.Location = new System.Drawing.Point(19, 8);
            this.listViewEx1.Name = "listViewEx1";
            this.listViewEx1.Size = new System.Drawing.Size(512, 188);
            this.listViewEx1.TabIndex = 0;
            this.listViewEx1.UseCompatibleStateImageBehavior = false;
            this.listViewEx1.View = System.Windows.Forms.View.Details;
            // 
            // chGen_
            // 
            this.chGen_.Text = "年級";
            this.chGen_.Width = 55;
            // 
            // chStatus
            // 
            this.chStatus.Text = "狀態";
            this.chStatus.Width = 55;
            // 
            // chLock
            // 
            this.chLock.Text = "選社鎖定";
            this.chLock.Width = 75;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItem2,
            this.清空學生待處理ToolStripMenuItem,
            this.toolStripSeparator2,
            this.鎖定學生選社ToolStripMenuItem,
            this.移除選擇學生ToolStripMenuItem,
            this.toolStripSeparator1,
            this.移除選擇學生ToolStripMenuItem1});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(167, 126);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(166, 22);
            this.toolStripMenuItem2.Text = "將學生加入待處理";
            this.toolStripMenuItem2.Click += new System.EventHandler(this.toolStripMenuItem2_Click);
            // 
            // 清空學生待處理ToolStripMenuItem
            // 
            this.清空學生待處理ToolStripMenuItem.Name = "清空學生待處理ToolStripMenuItem";
            this.清空學生待處理ToolStripMenuItem.Size = new System.Drawing.Size(166, 22);
            this.清空學生待處理ToolStripMenuItem.Text = "清空待處理";
            this.清空學生待處理ToolStripMenuItem.Click += new System.EventHandler(this.清空學生待處理ToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(163, 6);
            // 
            // 鎖定學生選社ToolStripMenuItem
            // 
            this.鎖定學生選社ToolStripMenuItem.Name = "鎖定學生選社ToolStripMenuItem";
            this.鎖定學生選社ToolStripMenuItem.Size = new System.Drawing.Size(166, 22);
            this.鎖定學生選社ToolStripMenuItem.Text = "鎖定學生選社";
            this.鎖定學生選社ToolStripMenuItem.Click += new System.EventHandler(this.鎖定學生選社ToolStripMenuItem_Click);
            // 
            // 移除選擇學生ToolStripMenuItem
            // 
            this.移除選擇學生ToolStripMenuItem.Name = "移除選擇學生ToolStripMenuItem";
            this.移除選擇學生ToolStripMenuItem.Size = new System.Drawing.Size(166, 22);
            this.移除選擇學生ToolStripMenuItem.Text = "解除鎖定";
            this.移除選擇學生ToolStripMenuItem.Click += new System.EventHandler(this.移除選擇學生ToolStripMenuItem_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(163, 6);
            // 
            // 移除選擇學生ToolStripMenuItem1
            // 
            this.移除選擇學生ToolStripMenuItem1.Name = "移除選擇學生ToolStripMenuItem1";
            this.移除選擇學生ToolStripMenuItem1.Size = new System.Drawing.Size(166, 22);
            this.移除選擇學生ToolStripMenuItem1.Text = "移除選擇學生";
            this.移除選擇學生ToolStripMenuItem1.Click += new System.EventHandler(this.移除選擇學生ToolStripMenuItem1_Click);
            // 
            // ClubStudent
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.listViewEx1);
            this.Controls.Add(this.btnClearStudent);
            this.Controls.Add(this.btnInserStudent);
            this.Controls.Add(this.lbCourseCount);
            this.Name = "ClubStudent";
            this.Size = new System.Drawing.Size(550, 245);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.ButtonX btnClearStudent;
        private System.Windows.Forms.ColumnHeader chClass_;
        private System.Windows.Forms.ColumnHeader chSeatNo_;
        private System.Windows.Forms.ColumnHeader chName_;
        private System.Windows.Forms.ColumnHeader chGender_;
        private System.Windows.Forms.ColumnHeader chSNum_;
        private DevComponents.DotNetBar.ButtonX btnInserStudent;
        private DevComponents.DotNetBar.ButtonItem buttonItem2;
        private DevComponents.DotNetBar.LabelX lbCourseCount;
        private ListViewEX listViewEx1;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem 清空學生待處理ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 鎖定學生選社ToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem 移除選擇學生ToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem 移除選擇學生ToolStripMenuItem1;
        private System.Windows.Forms.ColumnHeader chGen_;
        private System.Windows.Forms.ColumnHeader chLock;
        private System.Windows.Forms.ColumnHeader chStatus;
    }
}
