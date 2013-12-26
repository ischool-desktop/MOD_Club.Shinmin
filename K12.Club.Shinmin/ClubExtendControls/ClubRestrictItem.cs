using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.UDT;
using Campus.Windows;
using System.Xml;
using FISCA.DSAUtil;
using DevComponents.DotNetBar.Controls;
using FISCA.Data;
using FISCA.Permission;

namespace K12.Club.Shinmin
{
    [FISCA.Permission.FeatureCode("K12.Club.Shinmin.ClubRestrictItem.cs", "選社限制")]
    public partial class ClubRestrictItem : DetailContentBase
    {
        //背景模式
        private BackgroundWorker BGW = new BackgroundWorker();
        private BackgroundWorker Save_BGW = new BackgroundWorker();

        private AccessHelper _AccessHelper = new AccessHelper();
        private QueryHelper _QueryHelper = new QueryHelper();
        private CLUBRecord ClubPrimary;
        private CLUBRecord Log_ClubPrimary;
        string SetStringIsInt = "必須輸入數字!!";

        //資料變更事件引發器
        private ChangeListener DataListener { get; set; }

        //背景忙碌
        private bool BkWBool = false;

        //權限
        internal static FeatureAce UserPermission;

        List<string> deptList;

        //資料檢查功能
        ErrorProvider ep_Grade1Limit = new ErrorProvider();
        ErrorProvider ep_Grade2Limit = new ErrorProvider();
        ErrorProvider ep_Grade3Limit = new ErrorProvider();
        ErrorProvider ep_Limit = new ErrorProvider();

        public ClubRestrictItem()
        {
            InitializeComponent();

            Group = "選社限制";

            UserPermission = UserAcl.Current[FISCA.Permission.FeatureCodeAttribute.GetCode(GetType())];
            this.Enabled = UserPermission.Editable;

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            Save_BGW.DoWork += new DoWorkEventHandler(Save_BGW_DoWork);
            Save_BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Save_BGW_RunWorkerCompleted);

            ClubEvents.ClubChanged += new EventHandler(ClubEvents_ClubChanged);

            DataListener = new ChangeListener();
            DataListener.Add(new TextBoxSource(tbGrade1Limit));
            DataListener.Add(new TextBoxSource(tbGrade2Limit));
            DataListener.Add(new TextBoxSource(tbGrade3Limit));
            DataListener.Add(new TextBoxSource(tbLimit));
            DataListener.Add(new ComboBoxSource(cbGenderRestrict, ComboBoxSource.ListenAttribute.SelectedIndex));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);
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
            Log_ClubPrimary = ClubPrimary.CopyExtension();

            //取得科別資料
            deptList = tool.GetQueryDeptList();
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
                FISCA.Presentation.Controls.MsgBox.Show("取得社團[選社限制]發生錯誤!!\n" + e.Error.Message);

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

            tbGrade1Limit.Text = string.Empty;
            //一年級人數上限
            if (ClubPrimary.Grade1Limit.HasValue)
            {
                tbGrade1Limit.Text = ClubPrimary.Grade1Limit.Value.ToString();
            }
            //二年級人數上限
            tbGrade2Limit.Text = string.Empty;
            if (ClubPrimary.Grade2Limit.HasValue)
            {
                tbGrade2Limit.Text = ClubPrimary.Grade2Limit.Value.ToString();
            }
            //三年級人數上限
            tbGrade3Limit.Text = string.Empty;
            if (ClubPrimary.Grade3Limit.HasValue)
            {
                tbGrade3Limit.Text = ClubPrimary.Grade3Limit.Value.ToString();
            }
            //人數總上限
            tbLimit.Text = string.Empty;
            if (ClubPrimary.Limit.HasValue)
            {
                tbLimit.Text = ClubPrimary.Limit.Value.ToString();
            }

            cbGenderRestrict.Text = ClubPrimary.GenderRestrict;

            List<string> list = new List<string>();
            listDepartment.Items.Clear();
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


            foreach (ListViewItem item in listDepartment.Items)
            {
                item.Checked = false;
            }

            if (ClubPrimary.DeptRestrict != "")
            {
                XmlElement xmlBase = DSXmlHelper.LoadXml(ClubPrimary.DeptRestrict);
                foreach (XmlElement xml in xmlBase.SelectNodes("Dept"))
                {
                    foreach (ListViewItem item in listDepartment.Items)
                    {
                        if (item.Text == xml.InnerText)
                        {
                            item.Checked = true;
                            break;
                        }
                    }
                }
            }

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

            if (!CheckDataIsError())
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
                MsgBox.Show("修改選社限制失敗\n" + ex.Message);
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                return;
            }

            StringBuilder sb = LogSet();

            FISCA.LogAgent.ApplicationLog.Log("社團", "修改選社限制", sb.ToString());
        }

        private StringBuilder LogSet()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(string.Format("已修改選社限制：(學年度「{0}」學期「{1}」社團「{2}」)", Log_ClubPrimary.SchoolYear.ToString(), Log_ClubPrimary.Semester.ToString(), Log_ClubPrimary.ClubName));

            string chenge = GetString("性別", Log_ClubPrimary.GenderRestrict, ClubPrimary.GenderRestrict);
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("一年級人數限制", GetIntString(Log_ClubPrimary.Grade1Limit), GetIntString(ClubPrimary.Grade1Limit));
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("二年級人數限制", GetIntString(Log_ClubPrimary.Grade2Limit), GetIntString(ClubPrimary.Grade2Limit));
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("三年級人數限制", GetIntString(Log_ClubPrimary.Grade3Limit), GetIntString(ClubPrimary.Grade3Limit));
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("總人數上限", GetIntString(Log_ClubPrimary.Limit), GetIntString(ClubPrimary.Limit));
            if (!string.IsNullOrEmpty(chenge))
                sb.AppendLine(chenge);

            chenge = GetString("科別限制", GetDeptString(Log_ClubPrimary.DeptRestrict), GetDeptString(ClubPrimary.DeptRestrict));
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

        private string GetIntString(int? a)
        {
            string name = a.HasValue ? a.Value.ToString() : "";
            return name;
        }

        private string GetDeptString(string a)
        {
            string b = "";
            if (a != "")
            {
                List<string> list = new List<string>();
                XmlElement xmlBase = DSXmlHelper.LoadXml(a);
                foreach (XmlElement xml in xmlBase.SelectNodes("Dept"))
                {
                    list.Add(xml.InnerText);
                }
                b = string.Join(",", list);
            }

            return b;
        }

        private void GetChengeObj()
        {
            //人數限制
            if (!string.IsNullOrEmpty(tbGrade1Limit.Text.Trim())) //一年級人限
                ClubPrimary.Grade1Limit = tool.StringIsInt_DefIsZero(tbGrade1Limit.Text.Trim());
            else
                ClubPrimary.Grade1Limit = null;
            if (!string.IsNullOrEmpty(tbGrade2Limit.Text.Trim())) //二年級人限
                ClubPrimary.Grade2Limit = tool.StringIsInt_DefIsZero(tbGrade2Limit.Text.Trim());
            else
                ClubPrimary.Grade2Limit = null;
            if (!string.IsNullOrEmpty(tbGrade3Limit.Text.Trim())) //三年級人限
                ClubPrimary.Grade3Limit = tool.StringIsInt_DefIsZero(tbGrade3Limit.Text.Trim());
            else
                ClubPrimary.Grade3Limit = null;
            if (!string.IsNullOrEmpty(tbLimit.Text.Trim())) //人數上限
                ClubPrimary.Limit = tool.StringIsInt_DefIsZero(tbLimit.Text.Trim());
            else
                ClubPrimary.Limit = null;

            //性別限制
            if (cbGenderRestrict.SelectedItem != null) //男女限制
            {
                string res = cbGenderRestrict.GetItemText(cbGenderRestrict.SelectedItem);
                if (res == "")
                {
                    ClubPrimary.GenderRestrict = "";
                }
                else
                {
                    ClubPrimary.GenderRestrict = res;
                }
            }

            //科別限制
            DSXmlHelper dsXml = new DSXmlHelper("Department");
            foreach (ListViewItem each in listDepartment.Items)
            {
                if (each.Checked)
                {
                    dsXml.AddElement("Dept");
                    dsXml.AddText("Dept", each.Text);
                }
            }
            ClubPrimary.DeptRestrict = dsXml.BaseElement.OuterXml;
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

        private bool CheckDataIsError()
        {
            bool k = true;

            //選填欄位是否為正確之資料內容
            bool a = SetLimit(tbGrade1Limit, ep_Grade1Limit, SetStringIsInt);
            bool b = SetLimit(tbGrade2Limit, ep_Grade2Limit, SetStringIsInt);
            bool c = SetLimit(tbGrade3Limit, ep_Grade3Limit, SetStringIsInt);
            bool d = SetLimit(tbLimit, ep_Limit, SetStringIsInt);

            if (!(a && b && c && d))
                k = false;

            return k;
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

        private void listDepartment_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            SaveButtonVisible = true;
            CancelButtonVisible = true;
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
    }
}
