using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Data;
using FISCA.UDT;
using System.IO;
using System.Drawing.Imaging;
using K12.Data;
using FISCA.Permission;

namespace K12.Club.Shinmin
{
    //因為照片屬於比較大的檔案
    //因此儲存時多提供一個背景模式
    [FISCA.Permission.FeatureCode("K12.Club.Shinmin.ClubImageItem.cs", "社團照片")]
    public partial class ClubImageItem : DetailContentBase
    {
        //背景模式
        private BackgroundWorker BGW = new BackgroundWorker();
        private AccessHelper _AccessHelper = new AccessHelper();

        CLUBRecord ClubPrimary;

        //權限
        internal static FeatureAce UserPermission;

        //背景忙碌
        private bool BkWBool = false;

        public ClubImageItem()
        {
            InitializeComponent();

            Group = "社團照片";

            UserPermission = UserAcl.Current[FISCA.Permission.FeatureCodeAttribute.GetCode(GetType())];
            this.Enabled = UserPermission.Editable;

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            ClubEvents.ClubChanged += new EventHandler(ClubEvents_ClubChanged);
        }

        void ClubEvents_ClubChanged(object sender, EventArgs e)
        {
            Changed();
        }

        //切換學生
        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            Changed();
        }

        private void Changed()
        {
            #region 更新時
            if (this.PrimaryKey != "")
            {
                this.Loading = true;

                if (BGW.IsBusy)
                {
                    BkWBool = true;
                }
                else
                {
                    BGW.RunWorkerAsync();
                }
            }
            #endregion
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得社團資料
            List<CLUBRecord> ClubPrimaryList = _AccessHelper.Select<CLUBRecord>(string.Format("UID = '{0}'", this.PrimaryKey));
            if (ClubPrimaryList.Count != 1)
            {
                //如果取得2門以上 或 沒取得社團時
                e.Cancel = true;
                return;
            }

            ClubPrimary = ClubPrimaryList[0];
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Loading = false;

            if (e.Cancelled)
            {
                return;
            }

            if (e.Error != null)
            {
                FISCA.Presentation.Controls.MsgBox.Show("取得[社團照片]發生錯誤!!\n" + e.Error.Message);

                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }

            if (BkWBool) //如果有其他的更新事件
            {
                BkWBool = false;
                BGW.RunWorkerAsync();
                return;
            }

            BindData();

        }

        private ImageCodecInfo GetEncoder(ImageFormat format)
        {
            ImageCodecInfo[] codecs = ImageCodecInfo.GetImageDecoders();

            foreach (ImageCodecInfo codec in codecs)
            {
                if (codec.FormatID == format.Guid)
                {
                    return codec;
                }
            }
            return null;
        }

        private void BindData()
        {
            //社團照片1
            if (!string.IsNullOrEmpty(ClubPrimary.Photo1))
            {

                byte[] bs = Convert.FromBase64String(ClubPrimary.Photo1);
                MemoryStream ms = new MemoryStream(bs);
                Bitmap bm = new Bitmap(ms);
                Bitmap newBmp = new Bitmap(bm, pic1.Size);

                pic1.Image = newBmp;
                pic1.Tag = ClubPrimary.Photo1;
            }
            else
            {
                pic1.Image = pic1.InitialImage;
                pic1.Tag = null;
            }

            //社團照片2
            if (!string.IsNullOrEmpty(ClubPrimary.Photo2))
            {
                byte[] bs = Convert.FromBase64String(ClubPrimary.Photo2);
                MemoryStream ms = new MemoryStream(bs);
                Bitmap bm = new Bitmap(ms);
                Bitmap newBmp = new Bitmap(bm, pic2.Size);

                pic2.Image = newBmp;
                pic2.Tag = ClubPrimary.Photo2;
            }
            else
            {
                pic2.Image = pic2.InitialImage;
                pic2.Tag = null;
            }

            SaveButtonVisible = false;
            CancelButtonVisible = false;
        }

        /// <summary>
        /// 按下儲存時
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSaveButtonClick(EventArgs e)
        {
            //當使用者選擇的資料,等於目前系統取得的資料時
            if (this.PrimaryKey == ClubPrimary.UID)
            {
                BackgroundWorker bgwPhoto = new BackgroundWorker();
                bgwPhoto.DoWork += new DoWorkEventHandler(bgwPhoto_DoWork);
                bgwPhoto.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgwPhoto_RunWorkerCompleted);

                if (pic1.Tag != null)
                {
                    ClubPrimary.Photo1 = "" + pic1.Tag;
                }
                else
                {
                    ClubPrimary.Photo1 = string.Empty;
                }

                if (pic2.Tag != null)
                {
                    ClubPrimary.Photo2 = "" + pic2.Tag;
                }
                else
                {
                    ClubPrimary.Photo2 = string.Empty;
                }

                bgwPhoto.RunWorkerAsync();
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("資料不同步\n儲存失敗");
                return;
            }

        }

        //Save
        void bgwPhoto_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                _AccessHelper.UpdateValues(new List<CLUBRecord>() { ClubPrimary });
            }
            catch
            {
                e.Cancel = true;
                return;
            }
            FISCA.LogAgent.ApplicationLog.Log("社團", "變更社團照片", string.Format("已變更照片：\n學年度「{0}」學期「{1}」社團名稱「{2}」", ClubPrimary.SchoolYear.ToString(), ClubPrimary.Semester.ToString(), ClubPrimary.ClubName));
        }

        //儲存完成
        void bgwPhoto_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Loading = false;

            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    ClubEvents.RaiseAssnChanged();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("資料儲存失敗!!\n" + e.Error.Message);
                    return;
                }
            }
            else
            {
                FISCA.Presentation.Controls.MsgBox.Show("儲存作業已取消!!");
                return;
            }
        }

        /// <summary>
        /// 取消儲存時
        /// </summary>
        /// <param name="e"></param>
        protected override void OnCancelButtonClick(EventArgs e)
        {
            SaveButtonVisible = false;
            CancelButtonVisible = false;

            this.Loading = true;

            //判斷是否忙碌後,開始進行資料重置
            Changed();
        }

        #region 照片1
        private void buttonItem1_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "所有影像(*.jpg,*.jpeg)|*.jpg;*.jpeg;";
            if (od.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = null;
                try
                {
                    fs = new FileStream(od.FileName, FileMode.Open);
                    Bitmap orgBmp = new Bitmap(fs);
                    if (orgBmp.Width > 640 || orgBmp.Height > 480)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("照片大小不可超過640(W)*480(H)\n您所選的圖檔" + orgBmp.Width + "(W)*" + orgBmp.Height + "(H)");
                        return;
                    }
                    fs.Close();

                    //照片需要Resize
                    Bitmap newBmp = new Bitmap(orgBmp, pic1.Size);
                    pic1.Image = newBmp;

                    //Tag則Resize為orgBmp的大小
                    Size s = new System.Drawing.Size(orgBmp.Width, orgBmp.Height);
                    Bitmap new2Bmp = new Bitmap(orgBmp, s);

                    pic1.Tag = ResizeX1.ToBase64String_x1(new2Bmp, ImageFormat.Jpeg);

                    SaveButtonVisible = true;
                    CancelButtonVisible = true;
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show(ex.Message);
                }
            }
        }

        private void buttonItem5_Click(object sender, EventArgs e)
        {
            if (FISCA.Presentation.Controls.MsgBox.Show("您確定要清除此照片嗎？", "", MessageBoxButtons.YesNo) == DialogResult.No) return;

            try
            {
                ClubPrimary.Photo1 = string.Empty;
                pic1.Image = pic1.InitialImage;
                pic1.Tag = null;
                SaveButtonVisible = true;
                CancelButtonVisible = true;
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show(ex.Message);
            }
        }
        #endregion

        #region 照片2
        private void buttonItem3_Click(object sender, EventArgs e)
        {
            OpenFileDialog od = new OpenFileDialog();
            od.Filter = "所有影像(*.jpg,*.jpeg)|*.jpg;*.jpeg;";
            if (od.ShowDialog() == DialogResult.OK)
            {
                FileStream fs = null;
                try
                {
                    fs = new FileStream(od.FileName, FileMode.Open);
                    Bitmap orgBmp = new Bitmap(fs);
                    if (orgBmp.Width > 640 || orgBmp.Height > 480)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("照片大小不可超過640(W)*480(H)\n您所選的圖檔" + orgBmp.Width + "(W)*" + orgBmp.Height + "(H)");
                        return;
                    }
                    fs.Close();

                    //照片需要Resize
                    Bitmap newBmp = new Bitmap(orgBmp, pic2.Size);
                    pic2.Image = newBmp;

                    //Tag則Resize為orgBmp的大小
                    Size s = new System.Drawing.Size(orgBmp.Width, orgBmp.Height);
                    Bitmap new2Bmp = new Bitmap(orgBmp, s);
                    //pic2.Tag = ResizeX1.ToBase64String(new2Bmp, ImageFormat.Jpeg);
                    pic2.Tag = ResizeX1.ToBase64String_x1(new2Bmp, ImageFormat.Jpeg);

                    SaveButtonVisible = true;
                    CancelButtonVisible = true;
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show(ex.Message);
                }
            }
        }

        private void buttonItem6_Click(object sender, EventArgs e)
        {
            if (FISCA.Presentation.Controls.MsgBox.Show("您確定要清除此照片嗎？", "", MessageBoxButtons.YesNo) == DialogResult.No) return;

            try
            {
                ClubPrimary.Photo2 = string.Empty;
                pic2.Image = pic2.InitialImage;
                pic2.Tag = null;
                SaveButtonVisible = true;
                CancelButtonVisible = true;
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show(ex.Message);
            }
        }
        #endregion

        private void pic1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (pic1.Tag != null)
            {
                string tag = "" + pic1.Tag;
                ViewJPG view = new ViewJPG(tag);
                view.ShowDialog();
            }
        }

        private void pic2_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (pic2.Tag != null)
            {
                string tag = "" + pic2.Tag;
                ViewJPG view = new ViewJPG(tag);
                view.ShowDialog();
            }
        }

    }
}
