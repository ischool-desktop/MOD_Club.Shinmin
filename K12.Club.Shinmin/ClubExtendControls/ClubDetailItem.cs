using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.UDT;
using K12.Data;
using DevComponents.DotNetBar;
using FISCA.Data;
using Campus.Windows;
using DevComponents.DotNetBar.Controls;
using FISCA.Permission;

namespace K12.Club.Shinmin
{
    [FISCA.Permission.FeatureCode("K12.Club.Shinmin.ClubDetailItem.cs", "基本資料")]
    public partial class ClubDetailItem : DetailContentBase
    {
        //背景模式
        private BackgroundWorker BGW = new BackgroundWorker();
        private BackgroundWorker Save_BGW = new BackgroundWorker();

        CLUBRecord ClubPrimary;
        CLUBRecord Log_ClubPrimary;

        //UDT物件
        private AccessHelper _AccessHelper = new AccessHelper();
        private QueryHelper _QueryHelper = new QueryHelper();

        ErrorProvider ep_ClubName = new ErrorProvider();
        ErrorProvider ep_Teacher1 = new ErrorProvider();
        ErrorProvider ep_Teacher2 = new ErrorProvider();
        ErrorProvider ep_Teacher3 = new ErrorProvider();
        //ErrorProvider ep_President = new ErrorProvider();
        //ErrorProvider ep_VicePresident = new ErrorProvider();

        /// <summary>
        /// 社團學生名稱
        /// </summary>
        //Dictionary<string, string> StudentNameDic = new Dictionary<string, string>();

        //背景忙碌
        private bool BkWBool = false;

        private ChangeListener DataListener { get; set; } //資料變更事件引發器

        //List<SCJoin> ClubStudentList = new List<SCJoin>();

        //Dictionary<string, SCJoin> ClubStudentDic = new Dictionary<string, SCJoin>();

        //上課地點
        List<string> ClubLocation = new List<string>();

        List<string> ClubCategory = new List<string>();

        List<TeacherObj> TeacherList = new List<TeacherObj>();

        Dictionary<string, TeacherObj> TeacherDic = new Dictionary<string, TeacherObj>();

        //權限
        internal static FeatureAce UserPermission;

        public ClubDetailItem()
        {
            InitializeComponent();

            Group = "基本資料";

            UserPermission = UserAcl.Current[FISCA.Permission.FeatureCodeAttribute.GetCode(GetType())];
            this.Enabled = UserPermission.Editable;

            DataListener = new ChangeListener();
            DataListener.Add(new TextBoxSource(txtClubName));
            DataListener.Add(new TextBoxSource(txtAbout));
            //DataListener.Add(new TextBoxSource(tbCategory));
            DataListener.Add(new TextBoxSource(tbCLUBNumber));
            DataListener.Add(new ComboBoxSource(cbTeacher1, ComboBoxSource.ListenAttribute.Text));
            DataListener.Add(new ComboBoxSource(cbTeacher2, ComboBoxSource.ListenAttribute.Text));
            DataListener.Add(new ComboBoxSource(cbTeacher3, ComboBoxSource.ListenAttribute.Text));
            DataListener.Add(new ComboBoxSource(cbCategory, ComboBoxSource.ListenAttribute.Text));
            //DataListener.Add(new ComboBoxSource(cbPresident, ComboBoxSource.ListenAttribute.Text));
            DataListener.Add(new ComboBoxSource(cbLocation, ComboBoxSource.ListenAttribute.Text));
            //DataListener.Add(new ComboBoxSource(cbVicePresident, ComboBoxSource.ListenAttribute.Text));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

            ClubEvents.ClubChanged += new EventHandler(ClubEvents_ClubChanged);

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            Save_BGW.DoWork += new DoWorkEventHandler(Save_BGW_DoWork);
            Save_BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Save_BGW_RunWorkerCompleted);
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

        void Changed()
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
            //StudentNameDic.Clear();

            //取得社團資料
            List<CLUBRecord> ClubPrimaryList = _AccessHelper.Select<CLUBRecord>(string.Format("UID = '{0}'", this.PrimaryKey));
            if (ClubPrimaryList.Count != 1)
            {
                //如果取得2門以上 或 沒取得社團時
                e.Cancel = true;
                return;
            }

            ClubPrimary = ClubPrimaryList[0];
            Log_ClubPrimary = ClubPrimary.CopyExtension();

            //取得場地[GROUP BY]
            string TableName = Tn._CLUBRecordUDT;
            DataTable dt = _QueryHelper.Select("select location from " + TableName.ToLower() + " group by location ORDER by location");
            ClubLocation.Clear();
            foreach (DataRow row in dt.Rows)
            {
                string loc = "" + row[0];
                ClubLocation.Add(loc);
            }
            ClubLocation.Sort();

            //取得社團類型[Group By]
            TableName = Tn._CLUBRecordUDT;
            dt = _QueryHelper.Select("select club_category from " + TableName.ToLower() + " group by club_category ORDER by club_category");
            ClubCategory.Clear();
            foreach (DataRow row in dt.Rows)
            {
                string loc = "" + row[0];
                if (string.IsNullOrEmpty(loc))
                    continue;
                ClubCategory.Add(loc);
            }
            ClubCategory.Sort();

            //取得老師資料
            TeacherList.Clear();
            TeacherDic.Clear();
            dt = _QueryHelper.Select("select teacher.id,teacher.teacher_name,teacher.nickname from teacher ORDER by teacher_name");
            foreach (DataRow row in dt.Rows)
            {
                TeacherObj obj = new TeacherObj();
                obj.TeacherID = "" + row[0];
                obj.TeacherName = "" + row[1];
                obj.TeacherNickName = "" + row[2];
                TeacherList.Add(obj);

                if (!TeacherDic.ContainsKey(obj.TeacherID))
                {
                    TeacherDic.Add(obj.TeacherID, obj);
                }
            }

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
                MsgBox.Show("取得[基本資料]發生錯誤!!\n" + e.Error.Message);

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

        private void BindData()
        {
            DataListener.SuspendListen(); //終止變更判斷

            ep_ClubName.SetError(txtClubName, null);
            ep_Teacher1.SetError(cbTeacher1, null);
            ep_Teacher2.SetError(cbTeacher2, null);
            ep_Teacher3.SetError(cbTeacher3, null);
            //ep_President.SetError(cbPresident, null);
            //ep_VicePresident.SetError(cbVicePresident, null);

            tbCLUBNumber.Text = ClubPrimary.ClubNumber;
            txtClubName.Text = ClubPrimary.ClubName;
            lbSchoolYear.Text = ClubPrimary.SchoolYear + "學年度　第" + ClubPrimary.Semester + "學期";
            txtAbout.Text = ClubPrimary.About.Replace("<br>", "\r\n");
            //tbCategory.Text = ClubPrimary.ClubCategory;

            #region 社團老師

            cbTeacher1.Items.Clear();
            cbTeacher1.Text = "";
            cbTeacher1.DisplayMember = "TeacherFullName";
            cbTeacher1.Items.AddRange(TeacherList.ToArray());
            
            foreach (TeacherObj each in TeacherList)
            {
                if (each.TeacherID == ClubPrimary.RefTeacherID)
                {
                    //理論上只會被執行一次
                    cbTeacher1.Text = each.TeacherFullName;
                }
            }

            cbTeacher2.Items.Clear();
            cbTeacher2.Text = "";
            cbTeacher2.DisplayMember = "TeacherFullName";
            cbTeacher2.Items.AddRange(TeacherList.ToArray());

            foreach (TeacherObj each in TeacherList)
            {
                if (each.TeacherID == ClubPrimary.RefTeacherID2)
                {
                    //理論上只會被執行一次
                    cbTeacher2.Text = each.TeacherFullName;
                }
            }

            cbTeacher3.Items.Clear();
            cbTeacher3.Text = "";
            cbTeacher3.DisplayMember = "TeacherFullName";
            cbTeacher3.Items.AddRange(TeacherList.ToArray());

            foreach (TeacherObj each in TeacherList)
            {
                if (each.TeacherID == ClubPrimary.RefTeacherID3)
                {
                    //理論上只會被執行一次
                    cbTeacher3.Text = each.TeacherFullName;
                }
            }

            #endregion
            
            #region 場地資料

            cbLocation.Items.Clear();
            cbLocation.Text = "";
            cbLocation.Items.AddRange(ClubLocation.ToArray());
            cbLocation.Text = ClubPrimary.Location;

            #endregion

            #region 類型資料

            cbCategory.Items.Clear();
            cbCategory.Text = "";
            cbCategory.Items.AddRange(ClubCategory.ToArray());
            cbCategory.Text = ClubPrimary.ClubCategory;

            #endregion

            #region 社團資料內配對出社團幹部
            //cbPresident.Items.Clear();
            //cbPresident.Text = "";
            //cbVicePresident.Items.Clear();
            //cbVicePresident.Text = "";


            //foreach (string eachstring in StudentNameDic.Keys)
            //{
            //    if (ClubStudentDic.ContainsKey(eachstring))
            //    {
            //        ComboBoxItem item = new ComboBoxItem();
            //        item.Tag = ClubStudentDic[eachstring];

            //        item.Text = StudentNameDic[ClubStudentDic[eachstring].RefStudentID];

            //        cbPresident.Items.Add(item);
            //        cbVicePresident.Items.Add(item);

            //        if (ClubStudentDic[eachstring].RefStudentID == ClubPrimary.President)
            //        {
            //            cbPresident.SelectedItem = item;
            //        }

            //        if (ClubStudentDic[eachstring].RefStudentID == ClubPrimary.VicePresident)
            //        {
            //            cbVicePresident.SelectedItem = item;
            //        }
            //    }
            //}

            ////社長
            //if (!string.IsNullOrEmpty(ClubPrimary.President))
            //{
            //    if (StudentNameDic.ContainsKey(ClubPrimary.President))
            //    {
            //        ep_President.SetError(cbPresident, null);
            //        cbPresident.Text = StudentNameDic[ClubPrimary.President];
            //    }
            //    else
            //    {
            //        StudentRecord sr = Student.SelectByID(ClubPrimary.President);
            //        if (sr != null)
            //        {

            //            ep_President.SetError(cbPresident, "社長[" + sr.Name + "]已不存在本社團");
            //            cbPresident.Text = "";
            //        }
            //        else
            //        {
            //            ep_President.SetError(cbPresident, "社長系統編號[" + ClubPrimary.President + "]不存在系統內");
            //            cbPresident.Text = "";
            //        }

            //    }
            //}

            //if (!string.IsNullOrEmpty(ClubPrimary.VicePresident))
            //{
            //    //副社長
            //    if (StudentNameDic.ContainsKey(ClubPrimary.VicePresident))
            //    {
            //        ep_VicePresident.SetError(cbVicePresident, null);
            //        cbVicePresident.Text = StudentNameDic[ClubPrimary.VicePresident];
            //    }
            //    else
            //    {
            //        StudentRecord sr = Student.SelectByID(ClubPrimary.VicePresident);
            //        if (sr != null)
            //        {
            //            ep_VicePresident.SetError(cbVicePresident, "副社長[" + sr.Name + "]已不存在本社團");
            //            cbVicePresident.Text = "";
            //        }
            //        else
            //        {
            //            ep_VicePresident.SetError(cbVicePresident, "副社長系統編號[" + ClubPrimary.VicePresident + "]不存在系統內");
            //            cbPresident.Text = "";
            //        }

            //    }
            //}

            #endregion

            BkWBool = false;
            SaveButtonVisible = false;
            CancelButtonVisible = false;

            DataListener.Reset();
            DataListener.ResumeListen();
        }

        /// <summary>
        /// 當資料變更時
        /// </summary>
        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            SaveButtonVisible = (e.Status == ValueStatus.Dirty);
            CancelButtonVisible = (e.Status == ValueStatus.Dirty);
        }

        /// <summary>
        /// 按下儲存時
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSaveButtonClick(EventArgs e)
        {
            if (Save_BGW.IsBusy)
            {
                MsgBox.Show("系統忙碌中...");
                return;
            }

            if (this.PrimaryKey != ClubPrimary.UID)
            {
                MsgBox.Show("資料不同步\n儲存失敗");
                return;
            }

            if (!CheckData())
            {
                MsgBox.Show("請修正資料再儲存!!");
                return;
            }

            GetChengeObj();
            Save_BGW.RunWorkerAsync();
        }

        void Save_BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                _AccessHelper.UpdateValues(new List<CLUBRecord>() { ClubPrimary });
            }
            catch (Exception ex)
            {
                MsgBox.Show("資料儲存失敗!\n" + ex.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                return;
            }

            StringBuilder sb = LogSet();

            FISCA.LogAgent.ApplicationLog.Log("社團", "修改基本資料", sb.ToString());
        }

        private StringBuilder LogSet()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("已修改基本資料：(學年度「{0}」學期「{1}」社團「{2}」)", Log_ClubPrimary.SchoolYear.ToString(), Log_ClubPrimary.Semester.ToString(), Log_ClubPrimary.ClubName));

            string chenge = GetString("名稱", Log_ClubPrimary.ClubName, ClubPrimary.ClubName);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("代碼", Log_ClubPrimary.ClubNumber, ClubPrimary.ClubNumber);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("場地", Log_ClubPrimary.Location, ClubPrimary.Location);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("類型", Log_ClubPrimary.ClubCategory, ClubPrimary.ClubCategory);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("簡介", Log_ClubPrimary.About, ClubPrimary.About);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetStringTeacher("教師1", Log_ClubPrimary.RefTeacherID, ClubPrimary.RefTeacherID);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetStringTeacher("教師2", Log_ClubPrimary.RefTeacherID2, ClubPrimary.RefTeacherID2);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetStringTeacher("教師3", Log_ClubPrimary.RefTeacherID3, ClubPrimary.RefTeacherID3);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            return sb;
        }

        private string GetString(string a, string b, string c)
        {
            if (b != c)
                return string.Format("{0}由「{1}」修改為「{2}」", a, b, c);
            else
                return "";
        }

        private string GetStringTeacher(string a, string 修改前ID, string 修改後ID)
        {
            if (修改前ID != 修改後ID)
            {
                string 修改前老師名稱 = "";
                if (TeacherDic.ContainsKey(修改前ID))
                {
                    修改前老師名稱 = TeacherDic[修改前ID].TeacherFullName;
                }


                string 修改後老師名稱 = "";
                if (TeacherDic.ContainsKey(修改後ID))
                {
                    修改後老師名稱 = TeacherDic[修改後ID].TeacherFullName;
                }
                return string.Format("{0}由「{1}」修改為「{2}」", a, 修改前老師名稱, 修改後老師名稱);
            }
            else
            {
                return "";
            }
        }

        void Save_BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!e.Cancelled)
            {
                if (e.Error == null)
                {
                    ClubEvents.RaiseAssnChanged();
                }
                else
                {
                    MsgBox.Show("取得資料發生錯誤!!\n" + e.Error.Message);

                }
            }
            else
            {
                MsgBox.Show("作業中止!!");
            }
        }

        /// <summary>
        /// 將使用者修改內容填入物件
        /// </summary>
        private void GetChengeObj()
        {
            ClubPrimary.ClubName = txtClubName.Text.Trim();
            //社團老師
            if (cbTeacher1.SelectedItem != null)
            {
                TeacherObj teach = (TeacherObj)cbTeacher1.SelectedItem;
                ClubPrimary.RefTeacherID = teach.TeacherID;
            }
            else
            {
                ClubPrimary.RefTeacherID = "";
            }
            //社團老師2
            if (cbTeacher2.SelectedItem != null)
            {
                TeacherObj teach = (TeacherObj)cbTeacher2.SelectedItem;
                ClubPrimary.RefTeacherID2 = teach.TeacherID;
            }
            else
            {
                ClubPrimary.RefTeacherID2 = "";
            }
            //社團老師3
            if (cbTeacher3.SelectedItem != null)
            {
                TeacherObj teach = (TeacherObj)cbTeacher3.SelectedItem;
                ClubPrimary.RefTeacherID3 = teach.TeacherID;
            }
            else
            {
                ClubPrimary.RefTeacherID3 = "";
            }
            //場地
            ClubPrimary.Location = cbLocation.Text.Trim();
            //關於
            string refText = txtAbout.Text.Replace("\r\n", "<br>");
            ClubPrimary.About = refText;
            //類型
            ClubPrimary.ClubCategory = cbCategory.Text;

            //社團編號(8/7)
            ClubPrimary.ClubNumber = tbCLUBNumber.Text;
        }

        private bool CheckData()
        {
            bool a = true;

            #region 社團名稱
            if (string.IsNullOrEmpty(txtClubName.Text.Trim()))
            {
                ep_ClubName.SetError(txtClubName, "社團必須有名稱");
                a = false;
            }
            else
            {
                if (ClubPrimary.ClubName != txtClubName.Text)
                {
                    //社團名稱+學年度+學期不可重覆
                    StringBuilder sb = new StringBuilder();
                    sb.Append("select club_name from " + Tn._CLUBRecordUDT.ToLower() + " ");
                    sb.Append("where club_name = '" + txtClubName.Text + "' ");
                    sb.Append("and school_year = '" + ClubPrimary.SchoolYear.ToString() + "' ");
                    sb.Append("and semester = '" + ClubPrimary.Semester.ToString() + "'");

                    DataTable dt = _QueryHelper.Select(sb.ToString());


                    if (dt.Rows.Count > 0)
                    {
                        ep_ClubName.SetError(txtClubName, "社團名稱重覆");
                        a = false;
                    }
                    else
                    {
                        ep_ClubName.SetError(txtClubName, null);
                    }
                }
                else
                {
                    ep_ClubName.SetError(txtClubName, null);
                }
            }


            #endregion

            //社團老師
            bool b = tool.ComboBoxValueInItemList(cbTeacher1);
            if (!SetComboBoxError(b, cbTeacher1, ep_Teacher1, "社團老師必須存在於下拉清單中!!"))
                a = false;

            bool c = tool.ComboBoxValueInItemList(cbTeacher2);
            if (!SetComboBoxError(c, cbTeacher2, ep_Teacher2, "社團老師必須存在於下拉清單中!!"))
                a = false;

            bool d = tool.ComboBoxValueInItemList(cbTeacher3);
            if (!SetComboBoxError(d, cbTeacher3, ep_Teacher3, "社團老師必須存在於下拉清單中!!"))
                a = false;

            ////社長
            //bool c = tool.ComboBoxValueInItemList(cbPresident);
            //if (!SetComboBoxError(c, cbPresident, ep_President, "社長必須是本社團成員!!"))
            //    a = false;
            //
            ////副社長
            //bool d = tool.ComboBoxValueInItemList(cbVicePresident);
            //if (!SetComboBoxError(d, cbVicePresident, ep_VicePresident, "副社長必須是本社團成員!!"))
            //    a = false;

            return a;
        }

        private bool SetComboBoxError(bool d, ComboBoxEx ex, ErrorProvider ep, string message)
        {
            bool k = true;
            if (!d)
            {
                ep.SetError(ex, message);
                k = false;
            }
            else
            {
                ep.SetError(ex, null);
            }

            return k;
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

            DataListener.SuspendListen(); //終止變更判斷

            //判斷是否忙碌後,開始進行資料重置
            Changed();
        }
    }
}
