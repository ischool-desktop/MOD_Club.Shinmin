using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using System.IO;

namespace K12.Club.Shinmin
{
    public partial class ViewJPG : BaseForm
    {
        string _FreshmanPhoto;

        public ViewJPG(string FreshmanPhoto)
        {
            InitializeComponent();
            _FreshmanPhoto = FreshmanPhoto;

        }

        private void ViewJPG_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_FreshmanPhoto))
            {
                byte[] bs = Convert.FromBase64String(_FreshmanPhoto);
                MemoryStream ms = new MemoryStream(bs);
                Bitmap bm = new Bitmap(ms);
                pictureBox1.Image = bm;

                //pictureBox1.Width = bm.Width;
                //pictureBox1.Height = bm.Height;

                this.Width = bm.Width + 10;
                this.Height = bm.Height + 35;
            }
        }
    }
}
