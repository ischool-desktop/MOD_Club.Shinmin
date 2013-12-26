namespace K12.Club.Shinmin
{
    partial class ClubImageItem
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
            this.contextMenuBar1 = new DevComponents.DotNetBar.ContextMenuBar();
            this.ctxChange1 = new DevComponents.DotNetBar.ButtonItem();
            this.buttonItem1 = new DevComponents.DotNetBar.ButtonItem();
            this.buttonItem5 = new DevComponents.DotNetBar.ButtonItem();
            this.ctxChange2 = new DevComponents.DotNetBar.ButtonItem();
            this.buttonItem3 = new DevComponents.DotNetBar.ButtonItem();
            this.buttonItem6 = new DevComponents.DotNetBar.ButtonItem();
            this.pic1 = new System.Windows.Forms.PictureBox();
            this.pic2 = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.contextMenuBar1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic2)).BeginInit();
            this.SuspendLayout();
            // 
            // contextMenuBar1
            // 
            this.contextMenuBar1.Items.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.ctxChange1,
            this.ctxChange2});
            this.contextMenuBar1.Location = new System.Drawing.Point(113, 9);
            this.contextMenuBar1.Margin = new System.Windows.Forms.Padding(4);
            this.contextMenuBar1.Name = "contextMenuBar1";
            this.contextMenuBar1.Size = new System.Drawing.Size(239, 25);
            this.contextMenuBar1.Stretch = true;
            this.contextMenuBar1.Style = DevComponents.DotNetBar.eDotNetBarStyle.Office2003;
            this.contextMenuBar1.TabIndex = 236;
            this.contextMenuBar1.TabStop = false;
            this.contextMenuBar1.Text = "contextMenuBar1";
            // 
            // ctxChange1
            // 
            this.ctxChange1.AutoExpandOnClick = true;
            this.ctxChange1.Name = "ctxChange1";
            this.ctxChange1.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.buttonItem1,
            this.buttonItem5});
            this.ctxChange1.Text = "Change 1";
            // 
            // buttonItem1
            // 
            this.buttonItem1.Name = "buttonItem1";
            this.buttonItem1.Text = "變更照片";
            this.buttonItem1.Click += new System.EventHandler(this.buttonItem1_Click);
            // 
            // buttonItem5
            // 
            this.buttonItem5.Name = "buttonItem5";
            this.buttonItem5.Text = "清除照片";
            this.buttonItem5.Click += new System.EventHandler(this.buttonItem5_Click);
            // 
            // ctxChange2
            // 
            this.ctxChange2.AutoExpandOnClick = true;
            this.ctxChange2.Name = "ctxChange2";
            this.ctxChange2.SubItems.AddRange(new DevComponents.DotNetBar.BaseItem[] {
            this.buttonItem3,
            this.buttonItem6});
            this.ctxChange2.Text = "Change 2";
            // 
            // buttonItem3
            // 
            this.buttonItem3.Name = "buttonItem3";
            this.buttonItem3.Text = "變更照片";
            this.buttonItem3.Click += new System.EventHandler(this.buttonItem3_Click);
            // 
            // buttonItem6
            // 
            this.buttonItem6.Name = "buttonItem6";
            this.buttonItem6.Text = "清除照片";
            this.buttonItem6.Click += new System.EventHandler(this.buttonItem6_Click);
            // 
            // pic1
            // 
            this.pic1.BackColor = System.Drawing.Color.Gainsboro;
            this.contextMenuBar1.SetContextMenuEx(this.pic1, this.ctxChange1);
            this.pic1.Image = global::K12.Club.Shinmin.Properties.Resources.defImages_new;
            this.pic1.InitialImage = global::K12.Club.Shinmin.Properties.Resources.defImages_new;
            this.pic1.Location = new System.Drawing.Point(14, 9);
            this.pic1.Margin = new System.Windows.Forms.Padding(4);
            this.pic1.Name = "pic1";
            this.pic1.Size = new System.Drawing.Size(257, 193);
            this.pic1.TabIndex = 235;
            this.pic1.TabStop = false;
            this.pic1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.pic1_MouseDoubleClick);
            // 
            // pic2
            // 
            this.pic2.BackColor = System.Drawing.Color.Gainsboro;
            this.contextMenuBar1.SetContextMenuEx(this.pic2, this.ctxChange2);
            this.pic2.Image = global::K12.Club.Shinmin.Properties.Resources.defImages_new;
            this.pic2.InitialImage = global::K12.Club.Shinmin.Properties.Resources.defImages_new;
            this.pic2.Location = new System.Drawing.Point(279, 9);
            this.pic2.Margin = new System.Windows.Forms.Padding(4);
            this.pic2.Name = "pic2";
            this.pic2.Size = new System.Drawing.Size(257, 193);
            this.pic2.TabIndex = 234;
            this.pic2.TabStop = false;
            this.pic2.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.pic2_MouseDoubleClick);
            // 
            // ClubImageItem
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.contextMenuBar1);
            this.Controls.Add(this.pic1);
            this.Controls.Add(this.pic2);
            this.Name = "ClubImageItem";
            this.Size = new System.Drawing.Size(550, 215);
            ((System.ComponentModel.ISupportInitialize)(this.contextMenuBar1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.pic2)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pic1;
        private System.Windows.Forms.PictureBox pic2;
        private DevComponents.DotNetBar.ContextMenuBar contextMenuBar1;
        private DevComponents.DotNetBar.ButtonItem ctxChange1;
        private DevComponents.DotNetBar.ButtonItem buttonItem1;
        private DevComponents.DotNetBar.ButtonItem buttonItem5;
        private DevComponents.DotNetBar.ButtonItem ctxChange2;
        private DevComponents.DotNetBar.ButtonItem buttonItem3;
        private DevComponents.DotNetBar.ButtonItem buttonItem6;
    }
}
