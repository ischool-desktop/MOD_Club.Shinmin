using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Data;
using DevComponents.DotNetBar;
using System.Xml;
using FISCA.DSAUtil;
using DevComponents.DotNetBar.Controls;
using FISCA.Data;

namespace K12.Club.Shinmin
{
    public partial class NewAddClub : BaseForm
    {
        BackgroundWorker BGW = new BackgroundWorker();

        List<TeacherObj> TeacherList = new List<TeacherObj>();

        private QueryHelper _QueryHelper = new QueryHelper();

        List<string> deptList;

        string SetStringIsInt = "必須輸入數字!!";

        bool IsChangeNow = false;

        List<string> ClubLocation = new List<string>();

        List<string> ClubCategory = new List<string>();

        ErrorProvider ep_ClubName = new ErrorProvider();
        ErrorProvider ep_Grade1Limit = new ErrorProvider();
        ErrorProvider ep_Grade2Limit = new ErrorProvider();
        ErrorProvider ep_Grade3Limit = new ErrorProvider();
        ErrorProvider ep_Limit = new ErrorProvider();

        Campus.Windows.ChangeListener _ChangeListener = new Campus.Windows.ChangeListener();

        public NewAddClub()
        {
            InitializeComponent();
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            _ChangeListener.StatusChanged += new EventHandler<Campus.Windows.ChangeEventArgs>(_ChangeListener_StatusChanged);
            _ChangeListener.Add(new Campus.Windows.TextBoxSource(txtClubName));
            //_ChangeListener.Add(new Campus.Windows.TextBoxSource(txtCategory));
            _ChangeListener.Add(new Campus.Windows.TextBoxSource(tbClubNumber)); //社團編號
            _ChangeListener.Add(new Campus.Windows.TextBoxSource(tbAboutClub));
            _ChangeListener.Add(new Campus.Windows.TextBoxSource(tbGrade1Limit));
            _ChangeListener.Add(new Campus.Windows.TextBoxSource(tbGrade2Limit));
            _ChangeListener.Add(new Campus.Windows.TextBoxSource(tbGrade3Limit));
            _ChangeListener.Add(new Campus.Windows.ComboBoxSource(cbTeacher, Campus.Windows.ComboBoxSource.ListenAttribute.SelectedIndex));
            //社團類型
            _ChangeListener.Add(new Campus.Windows.ComboBoxSource(cbCategory, Campus.Windows.ComboBoxSource.ListenAttribute.Text));
            _ChangeListener.Add(new Campus.Windows.ComboBoxSource(cbLocation, Campus.Windows.ComboBoxSource.ListenAttribute.Text));
            _ChangeListener.Add(new Campus.Windows.ComboBoxSource(cbGenderRestrict, Campus.Windows.ComboBoxSource.ListenAttribute.SelectedIndex));

        }

        void _ChangeListener_StatusChanged(object sender, Campus.Windows.ChangeEventArgs e)
        {
            IsChangeNow = (e.Status == Campus.Windows.ValueStatus.Dirty);
        }

        int _DefaultSchoolYear;
        int _DefaultSemester;

        private void NewAddClub_Load(object sender, EventArgs e)
        {
            cbGenderRestrict.SelectedIndex = 0;
            _ChangeListener.SuspendListen();

            this.Text = "新增社團(資料讀取中...)";
            SetFrom = false;

            BGW.RunWorkerAsync();
        }

        private bool SetFrom
        {
            set
            {
                intSchoolYear.Enabled = value;
                intSemester.Enabled = value;
                txtClubName.Enabled = value;
                tbClubNumber.Enabled = value;
                cbLocation.Enabled = value;
                cbCategory.Enabled = value;
                cbTeacher.Enabled = value;
                cbTeacher2.Enabled = value;
                cbTeacher3.Enabled = value;
                tbAboutClub.Enabled = value;
                tbGrade1Limit.Enabled = value;
                tbGrade2Limit.Enabled = value;
                tbGrade3Limit.Enabled = value;
                tbLimit.Enabled = value;
                cbGenderRestrict.Enabled = value;
                listDepartment.Enabled = value;
                btnSave.Enabled = value;
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            _DefaultSchoolYear = tool.StringIsInt_DefIsZero(K12.Data.School.DefaultSchoolYear);
            _DefaultSemester = tool.StringIsInt_DefIsZero(K12.Data.School.DefaultSemester);

            //取得老師資料
            TeacherList.Clear();
            DataTable dt = _QueryHelper.Select("select teacher.id,teacher.teacher_name,teacher.nickname from teacher ORDER by teacher_name");
            foreach (DataRow row in dt.Rows)
            {
                TeacherObj obj = new TeacherObj();
                obj.TeacherID = "" + row[0];
                obj.TeacherName = "" + row[1];
                obj.TeacherNickName = "" + row[2];
                TeacherList.Add(obj);
            }

            //取得場地[GROUP BY]
            string TableName = Tn._CLUBRecordUDT;
            dt = _QueryHelper.Select("select location from " + TableName.ToLower() + " group by location ORDER by location");
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

            //取得科別資料
            deptList = tool.GetQueryDeptList();

        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            SetFrom = true;
            this.Text = "新增社團";

            if (e.Cancelled)
            {
                MsgBox.Show("資料取得中止");
                return;
            }

            if (e.Error != null)
            {
                MsgBox.Show("部份資料取得發生錯誤!!");
                return;
            }

            intSchoolYear.Value = _DefaultSchoolYear;
            intSemester.Value = _DefaultSemester;

            #region 社團老師

            cbTeacher.Items.Clear();
            cbTeacher.Text = "";
            cbTeacher.DisplayMember = "TeacherFullName";
            cbTeacher.Items.AddRange(TeacherList.ToArray());

            cbTeacher2.Items.Clear();
            cbTeacher2.Text = "";
            cbTeacher2.DisplayMember = "TeacherFullName";
            cbTeacher2.Items.AddRange(TeacherList.ToArray());

            cbTeacher3.Items.Clear();
            cbTeacher3.Text = "";
            cbTeacher3.DisplayMember = "TeacherFullName";
            cbTeacher3.Items.AddRange(TeacherList.ToArray());

            #endregion

            #region 場地資料

            cbLocation.Items.Clear();
            cbLocation.Items.AddRange(ClubLocation.ToArray());

            #endregion

            #region 社團類型
            cbCategory.Items.Clear();
            cbCategory.Items.AddRange(ClubCategory.ToArray());

            #endregion

            #region 科別資訊
            listDepartment.Items.Clear();
            List<string> list = new List<string>();
            foreach (string each in deptList)
            {
                if (string.IsNullOrEmpty(each))
                    continue;

                if (list.Contains(each))
                    continue;

                list.Add(each);
                ListViewItem item = new ListViewItem();
                item.Tag = each;
                item.Text = each;
                listDepartment.Items.Add(item);
            }
            _ChangeListener.Reset();
            _ChangeListener.ResumeListen();
            IsChangeNow = false;
            #endregion
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //資料是否輸入檢查
            if (!CheckDataIsError())
            {
                MsgBox.Show("請輸入必填欄位!!");
                return;
            }

            if (!CheckClubName())
            {
                MsgBox.Show("社團名稱重覆!!");
                return;
            }

            CLUBRecord club = GetClub();

            BackgroundWorker BGW_Save = new BackgroundWorker();
            BGW_Save.DoWork += new DoWorkEventHandler(BGW_Save_DoWork);
            BGW_Save.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_Save_RunWorkerCompleted);
            //開始儲存資料
            this.Text = "新增社團(儲存中...)";
            SetFrom = false;
            BGW_Save.RunWorkerAsync(new List<CLUBRecord>() { club });

        }

        private CLUBRecord GetClub()
        {
            CLUBRecord club = new CLUBRecord();
            club.ClubName = txtClubName.Text.Trim(); //社團名稱
            club.SchoolYear = intSchoolYear.Value; //學年度
            club.Semester = intSemester.Value; //學期
            club.ClubCategory = cbCategory.Text.Trim(); //類型
            club.ClubNumber = tbClubNumber.Text.Trim(); //類型

            if (!string.IsNullOrEmpty(tbGrade1Limit.Text.Trim())) //一年級人限
                club.Grade1Limit = tool.StringIsInt_DefIsZero(tbGrade1Limit.Text.Trim());
            if (!string.IsNullOrEmpty(tbGrade2Limit.Text.Trim())) //二年級人限
                club.Grade2Limit = tool.StringIsInt_DefIsZero(tbGrade2Limit.Text.Trim());
            if (!string.IsNullOrEmpty(tbGrade3Limit.Text.Trim())) //三年級人限
                club.Grade3Limit = tool.StringIsInt_DefIsZero(tbGrade3Limit.Text.Trim());
            if (!string.IsNullOrEmpty(tbLimit.Text.Trim())) //三年級人限
                club.Limit = tool.StringIsInt_DefIsZero(tbLimit.Text.Trim());

            if (cbGenderRestrict.SelectedItem != null) //男女限制
            {
                string res = cbGenderRestrict.GetItemText(cbGenderRestrict.SelectedItem);
                club.GenderRestrict = res;
            }

            //社團老師

            if (cbTeacher.SelectedItem != null)
            {
                TeacherObj cbi = (TeacherObj)cbTeacher.SelectedItem;
                club.RefTeacherID = cbi.TeacherID;
            }


            if (cbTeacher2.SelectedItem != null)
            {
                TeacherObj cbi = (TeacherObj)cbTeacher2.SelectedItem;
                club.RefTeacherID2 = cbi.TeacherID;
            }


            if (cbTeacher3.SelectedItem != null)
            {
                TeacherObj cbi = (TeacherObj)cbTeacher3.SelectedItem;
                club.RefTeacherID3 = cbi.TeacherID;
            }


            //社團場地
            if (!string.IsNullOrEmpty(cbLocation.Text.Trim()))
                club.Location = cbLocation.Text.Trim();

            //社團限制 - 科別
            DSXmlHelper dsXml = new DSXmlHelper("Department");
            foreach (ListViewItem each in listDepartment.Items)
            {
                if (each.Checked)
                {
                    dsXml.AddElement("Dept");
                    dsXml.AddText("Dept", each.Text);

                }
            }
            club.DeptRestrict = dsXml.BaseElement.OuterXml;
            club.About = tbAboutClub.Text.Trim(); //簡介

            return club;
        }

        void BGW_Save_DoWork(object sender, DoWorkEventArgs e)
        {
            List<CLUBRecord> list = (List<CLUBRecord>)e.Argument;
            AccessHelper _accessHelper = new AccessHelper();
            try
            {
                _accessHelper.InsertValues(list);
            }
            catch (Exception ex)
            {
                e.Cancel = true;
                return;
            }

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("已新增一筆社團記錄：");
            sb.AppendLine(string.Format("學年度「{0}」學期「{1}」", list[0].SchoolYear.ToString(), list[0].Semester));
            sb.AppendLine(string.Format("社團名稱「{0}」", list[0].ClubName));
            FISCA.LogAgent.ApplicationLog.Log("社團", "新增社團", sb.ToString());
        }

        void BGW_Save_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Text = "新增社團";
            SetFrom = false;

            if (e.Cancelled)
            {
                MsgBox.Show("新增社團已中止!\n" + e.Error.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }

            if (e.Error != null)
            {
                MsgBox.Show("新增社團失敗\n" + e.Error.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }

            MsgBox.Show("新增社團成功！");
            ClubEvents.RaiseAssnChanged();
            IsChangeNow = false;
            this.Close();
        }

        private bool CheckDataIsError()
        {
            bool k = true;

            //社團名稱
            if (string.IsNullOrEmpty(txtClubName.Text))
            {
                k = false;
            }

            //選填欄位是否為正確之資料內容
            bool a = SetLimit(tbGrade1Limit, ep_Grade1Limit, SetStringIsInt);
            bool b = SetLimit(tbGrade2Limit, ep_Grade2Limit, SetStringIsInt);
            bool c = SetLimit(tbGrade3Limit, ep_Grade3Limit, SetStringIsInt);
            bool d = SetLimit(tbLimit, ep_Limit, SetStringIsInt);


            if (!(a && b && c && d))
                k = false;

            return k;
        }

        private bool CheckClubName()
        {
            bool k = true;

            //社團名稱+學年度+學期不可重覆
            StringBuilder sb = new StringBuilder();
            sb.Append("select club_name from " + Tn._CLUBRecordUDT.ToLower() + " ");
            sb.Append("where club_name = '" + txtClubName.Text + "' ");
            sb.Append("and school_year = '" + intSchoolYear.Value.ToString() + "' ");
            sb.Append("and semester = '" + intSemester.Value.ToString() + "'");

            DataTable dt = _QueryHelper.Select(sb.ToString());


            if (dt.Rows.Count > 0)
            {
                ep_ClubName.SetError(txtClubName, "社團名稱+學年度+學期不可重覆");
                return false;
            }
            else
            {
                ep_ClubName.SetError(txtClubName, null);
                return true;
            }
        }

        private bool SetLimit(TextBoxX x1, ErrorProvider ep, string ErrorString)
        {
            bool k = true;
            if (!string.IsNullOrEmpty(x1.Text.Trim()))
            {
                if (tool.StringIsInt_Bool(x1.Text.Trim()))
                {
                    ep.SetError(x1, null);
                    k = true;
                }
                else
                {
                    ep.SetError(x1, ErrorString);
                    k = false;
                }
            }
            else
            {
                ep.SetError(x1, null);
                k = true;
            }

            return k;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void tbGrade1Limit_TextChanged(object sender, EventArgs e)
        {
            AutoCountLimit();
            SetLimit(tbGrade1Limit, ep_Grade1Limit, SetStringIsInt);
        }

        private void tbGrade2Limit_TextChanged(object sender, EventArgs e)
        {
            AutoCountLimit();
            SetLimit(tbGrade2Limit, ep_Grade2Limit, SetStringIsInt);
        }

        private void tbGrade3Limit_TextChanged(object sender, EventArgs e)
        {
            AutoCountLimit();
            SetLimit(tbGrade3Limit, ep_Grade3Limit, SetStringIsInt);
        }

        /// <summary>
        /// 統計一/二/三年級之人數限制
        /// 並填入人數限制的欄位
        /// </summary>
        private void AutoCountLimit()
        {
            int a = tool.StringIsInt_DefIsZero(tbGrade1Limit.Text);
            int b = tool.StringIsInt_DefIsZero(tbGrade2Limit.Text);
            int c = tool.StringIsInt_DefIsZero(tbGrade3Limit.Text);
            tbLimit.Text = (a + b + c).ToString();
        }

        private void intSchoolYear_ValueChanged(object sender, EventArgs e)
        {
            IsChangeNow = true;
        }

        private void intSemester_ValueChanged(object sender, EventArgs e)
        {
            IsChangeNow = true;
        }

        private void listDepartment_ItemCheck(object sender, ItemCheckEventArgs e)
        {
            IsChangeNow = true;
        }

        private void NewAddClub_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsChangeNow)
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("確認放棄?", "尚未儲存資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr != System.Windows.Forms.DialogResult.Yes)
                {
                    e.Cancel = true;
                }
            }
        }
    }
}
