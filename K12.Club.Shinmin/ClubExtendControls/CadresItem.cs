using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.UDT;
using FISCA.Data;
using FISCA.Permission;
using K12.Data;
using FISCA.Presentation.Controls;
using DevComponents.DotNetBar.Controls;

namespace K12.Club.Shinmin
{
    [FISCA.Permission.FeatureCode("K12.Club.Shinmin.CadresItem.cs", "社團幹部")]
    public partial class CadresItem : DetailContentBase
    {
        //背景模式
        private BackgroundWorker BGW = new BackgroundWorker();
        private BackgroundWorker Save_BGW = new BackgroundWorker();

        //UDT物件
        private AccessHelper _AccessHelper = new AccessHelper();
        private QueryHelper _QueryHelper = new QueryHelper();

        //權限
        internal static FeatureAce UserPermission;

        //背景忙碌
        private bool BkWBool = false;

        CLUBRecord ClubPrimary;

        //社長
        StudentRecord StudPresident;
        //副社長
        StudentRecord StudVicePresident;
        List<StudentObj> StudRecordlist;

        Dictionary<int, string> StudentIDIndexDic = new Dictionary<int, string>();

        List<CadresRecord> CadresList;

        Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();
        Dictionary<string, StudentRecord> Log_StudentDic = new Dictionary<string, StudentRecord>();

        //資料變更事件引發器
        private Campus.Windows.ChangeListener DataListener { get; set; }

        public CadresItem()
        {
            //ClubPrimary.President
            //ClubPrimary.VicePresident

            InitializeComponent();
            Group = "社團幹部";

            UserPermission = UserAcl.Current[FISCA.Permission.FeatureCodeAttribute.GetCode(GetType())];
            this.Enabled = UserPermission.Editable;

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            Save_BGW.DoWork += new DoWorkEventHandler(Save_BGW_DoWork);
            Save_BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Save_BGW_RunWorkerCompleted);

            ClubEvents.ClubChanged += new EventHandler(ClubEvents_ClubChanged);
        }

        void ClubEvents_ClubChanged(object sender, EventArgs e)
        {
            Changed();
        }

        private void CadresItem_Load(object sender, EventArgs e)
        {
            DataListener = new Campus.Windows.ChangeListener();
            DataListener.Add(new Campus.Windows.DataGridViewSource(dataGridViewX1));
            DataListener.StatusChanged += new EventHandler<Campus.Windows.ChangeEventArgs>(DataListener_StatusChanged);

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
            List<string> studList = new List<string>();
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("ref_club_id='{0}'", ClubPrimary.UID));
            List<SCJoin> SCJoin_LIst = _AccessHelper.Select<SCJoin>(sb.ToString());
            foreach (SCJoin scj in SCJoin_LIst)
            {
                if (!studList.Contains(scj.RefStudentID))
                {
                    studList.Add(scj.RefStudentID);
                }
            }
            //依據社團內的幹部記錄取得
            //社長
            if (!string.IsNullOrEmpty(ClubPrimary.President))
            {
                if (studList.Contains(ClubPrimary.President))
                {
                    StudPresident = Student.SelectByID(ClubPrimary.President);
                }
                else
                {
                    StudPresident = null;
                }
            }
            else
            {
                StudPresident = null;
            }
            //副社長
            if (!string.IsNullOrEmpty(ClubPrimary.VicePresident))
            {
                if (studList.Contains(ClubPrimary.VicePresident))
                {
                    StudVicePresident = Student.SelectByID(ClubPrimary.VicePresident);
                }
                else
                {

                    StudVicePresident = null;
                }
            }
            else
            {
                StudVicePresident = null;
            }

            //取得老師指定的幹部資料CadresRecord
            CadresList = _AccessHelper.Select<CadresRecord>(string.Format("ref_club_id = '{0}'", this.PrimaryKey));

            //取得社團學生
            //建立下拉式選單內的學生清單
            StudentDic.Clear();
            StudentIDIndexDic.Clear();
            Log_StudentDic.Clear();
            List<StudentRecord> _StudRecordlist = Student.SelectByIDs(studList);
            _StudRecordlist = SortClassIndex.K12Data_StudentRecord(_StudRecordlist);
            //studlist.Sort();
            StudRecordlist = new List<StudentObj>();
            foreach (StudentRecord stud in _StudRecordlist)
            {
                StudentObj obj = new StudentObj();
                StringBuilder sb1 = new StringBuilder();
                sb1.Append(stud.Class != null ? stud.Class.Name : "");
                sb1.Append("　");
                sb1.Append(stud.SeatNo.HasValue ? stud.SeatNo.Value.ToString() : "");
                sb1.Append("　");
                sb1.Append(stud.Name);
                obj.TolName = sb1.ToString();
                obj.stuent = stud;
                StudRecordlist.Add(obj);

                if (!StudentDic.ContainsKey(sb1.ToString()))
                {
                    StudentDic.Add(sb1.ToString(), stud);
                }

                if (!Log_StudentDic.ContainsKey(stud.ID))
                {
                    Log_StudentDic.Add(stud.ID, stud);
                }
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                return;
            }

            if (e.Error != null)
            {
                this.Loading = false;
                FISCA.Presentation.Controls.MsgBox.Show("取得幹部資料發生錯誤!!");
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                return;
            }

            if (ClubPrimary == null)
            {
                this.Loading = false;
                return;
            }

            if (BkWBool) //如果有其他的更新事件
            {
                BkWBool = false;
                BGW.RunWorkerAsync();
                return;
            }

            BindData();

            DataListener.Reset();
            DataListener.ResumeListen();

            SaveButtonVisible = false;
            CancelButtonVisible = false;

            this.Loading = false;
        }

        private void BindData()
        {
            if (dataGridViewX1.Columns.Count != 0)
            {
                dataGridViewX1.Rows.Clear();

                //1.填入社長
                DataGridViewRow row1 = new DataGridViewRow();
                row1.CreateCells(dataGridViewX1);
                row1.Cells[0].Value = "社長";
                if (StudPresident != null)
                {
                    StringBuilder sb1 = new StringBuilder();
                    sb1.Append(StudPresident.Class != null ? StudPresident.Class.Name : "");
                    sb1.Append("　");
                    sb1.Append(StudPresident.SeatNo.HasValue ? StudPresident.SeatNo.Value.ToString() : "");
                    sb1.Append("　");
                    sb1.Append(StudPresident.Name);
                    row1.Cells[1].Value = sb1;
                    row1.ErrorText = "";
                }
                else
                {
                    row1.Cells[1].Value = "";
                    //當StudVicePresident為null,ClubPrimary.VicePresident卻又不是空值
                    if (ClubPrimary.President != "")
                    {
                        StudentRecord sp = Student.SelectByID(ClubPrimary.President);
                        row1.ErrorText = "社長：" + sp.Name + "，已不存在本社團";
                    }
                    else
                    {
                        row1.ErrorText = "";
                    }
                }
                dataGridViewX1.Rows.Add(row1);

                //2.填入副社長
                DataGridViewRow row2 = new DataGridViewRow();
                row2.CreateCells(dataGridViewX1);
                row2.Cells[0].Value = "副社長";
                if (StudVicePresident != null)
                {
                    StringBuilder sb1 = new StringBuilder();
                    sb1.Append(StudVicePresident.Class != null ? StudVicePresident.Class.Name : "");
                    sb1.Append("　");
                    sb1.Append(StudVicePresident.SeatNo.HasValue ? StudVicePresident.SeatNo.Value.ToString() : "");
                    sb1.Append("　");
                    sb1.Append(StudVicePresident.Name);
                    row2.Cells[1].Value = sb1;
                    row2.ErrorText = "";
                }
                else
                {
                    row2.Cells[1].Value = "";
                    //當StudVicePresident為null,ClubPrimary.VicePresident卻又不是空值
                    if (ClubPrimary.VicePresident != "")
                    {
                        StudentRecord sp = Student.SelectByID(ClubPrimary.VicePresident);
                        row2.ErrorText = "副社長：" + sp.Name + "，已不存在本社團";
                    }
                    else
                    {
                        row2.ErrorText = "";
                    }
                }
                dataGridViewX1.Rows.Add(row2);

                //3.填入剩餘的幹部記錄
                foreach (CadresRecord cr in CadresList)
                {
                    foreach (StudentRecord each in StudentDic.Values)
                    {
                        if (each.ID == cr.RefStudentID)
                        {
                            StringBuilder sb1 = new StringBuilder();
                            sb1.Append(each.Class != null ? each.Class.Name : "");
                            sb1.Append("　");
                            sb1.Append(each.SeatNo.HasValue ? each.SeatNo.Value.ToString() : "");
                            sb1.Append("　");
                            sb1.Append(each.Name);

                            if (StudentDic.ContainsKey(sb1.ToString()))
                            {
                                DataGridViewRow row3 = new DataGridViewRow();
                                row3.Tag = cr;
                                row3.CreateCells(dataGridViewX1);
                                row3.Cells[0].Value = cr.CadreName;
                                //社團幹部的學生ID,取得學生資料
                                row3.Cells[1].Value = sb1;
                                dataGridViewX1.Rows.Add(row3);
                            }

                            break;
                        }
                    }
                }

                Column2.DataSource = StudRecordlist;
                Column2.DisplayMember = "TolName";
                Column2.FlatStyle = FlatStyle.Standard;
            }
            else
            {

            }
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

            List<CadresRecord> InsertList = GetInsertList();

            Save_BGW.RunWorkerAsync(InsertList);

            #region 註解

            ////儲存時,社長與副社長資料,儲存於社團Record
            //List<CadresRecord> InsertList = new List<CadresRecord>();
            //foreach (DataGridViewRow row in dataGridViewX1.Rows)
            //{
            //    if (row.IsNewRow)
            //        continue;

            //    DevComponents.DotNetBar.Controls.DataGridViewComboBoxExCell cell = (DevComponents.DotNetBar.Controls.DataGridViewComboBoxExCell)row.Cells[1];
            //    if (row.Index == 0) //社長
            //    {
            //        //在Cell-1內所選的學生是誰?

            //        if (StudentDic.ContainsKey("" + cell.Value))
            //        {
            //            StudentRecord sr = StudentDic["" + cell.Value];
            //            ClubPrimary.President = sr.ID;
            //        }
            //        else
            //        {
            //            ClubPrimary.President = string.Empty;
            //        }

            //    }
            //    else if (row.Index == 1) //副社長
            //    {
            //        if (StudentDic.ContainsKey("" + cell.Value))
            //        {
            //            StudentRecord sr = StudentDic["" + cell.Value];
            //            ClubPrimary.VicePresident = sr.ID;
            //        }
            //        else
            //        {
            //            ClubPrimary.VicePresident = string.Empty;
            //        }
            //    }
            //    else //其他Row
            //    {
            //        if (StudentDic.ContainsKey("" + cell.Value))
            //        {
            //            StudentRecord sr = StudentDic["" + cell.Value];
            //            CadresRecord cr = new CadresRecord();
            //            cr.CadreName = "" + row.Cells[0].Value;
            //            cr.RefStudentID = sr.ID;
            //            cr.RefClubID = ClubPrimary.UID;
            //            InsertList.Add(cr);
            //        }
            //    }
            //}
            //List<CLUBRecord> list = new List<CLUBRecord>();
            //list.Add(ClubPrimary);
            //_AccessHelper.DeletedValues(CadresList); //刪除原有社團幹部
            //_AccessHelper.InsertValues(InsertList); //新增社團幹部
            //_AccessHelper.UpdateValues(list); //更新社團資料

            //Changed();

            //ClubEvents.RaiseAssnChanged(); 

            #endregion

        }

        private List<CadresRecord> GetInsertList()
        {
            //儲存時,社長與副社長資料,儲存於社團Record
            List<CadresRecord> InsertList = new List<CadresRecord>();
            Dictionary<string, bool> Dic = new Dictionary<string, bool>();
            Dic.Add("社長", false);
            Dic.Add("副社長", false);
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow)
                    continue;

                DataGridViewTextBoxCell cell_0 = (DataGridViewTextBoxCell)row.Cells[Column1.Index];
                DataGridViewComboBoxExCell cell_1 = (DataGridViewComboBoxExCell)row.Cells[Column2.Index];
                if ("" + cell_0.Value == "社長") //社長
                {
                    Dic["社長"] = true;
                    //在Cell-1內所選的學生是誰?
                    if (StudentDic.ContainsKey("" + cell_1.Value))
                    {
                        StudentRecord sr = StudentDic["" + cell_1.Value];
                        ClubPrimary.President = sr.ID;
                    }
                    else
                    {
                        ClubPrimary.President = string.Empty;
                    }

                }
                else if ("" + cell_0.Value == "副社長") //副社長
                {
                    Dic["副社長"] = true;
                    if (StudentDic.ContainsKey("" + cell_1.Value))
                    {
                        StudentRecord sr = StudentDic["" + cell_1.Value];
                        ClubPrimary.VicePresident = sr.ID;
                    }
                    else
                    {
                        ClubPrimary.VicePresident = string.Empty;
                    }
                }
                else //其他Row
                {
                    if (StudentDic.ContainsKey("" + cell_1.Value))
                    {
                        StudentRecord sr = StudentDic["" + cell_1.Value];
                        CadresRecord cr = new CadresRecord();
                        cr.CadreName = "" + cell_0.Value;
                        cr.RefStudentID = sr.ID;
                        cr.RefClubID = ClubPrimary.UID;
                        InsertList.Add(cr);
                    }
                }
            }


            if (!Dic["社長"])
            {
                ClubPrimary.President = string.Empty;
            }
            if (!Dic["副社長"])
            {
                ClubPrimary.VicePresident = string.Empty;
            }

            return InsertList;
        }

        void Save_BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            List<CadresRecord> InsertList = (List<CadresRecord>)e.Argument;
            List<CLUBRecord> list = new List<CLUBRecord>();
            list.Add(ClubPrimary);

            StringBuilder sb_Log = new StringBuilder();
            sb_Log.AppendLine(string.Format("修改社團幹部資料：(學年度「{0}」學期「{1}」社團「{2}」)", ClubPrimary.SchoolYear.ToString(), ClubPrimary.Semester.ToString(), ClubPrimary.ClubName));

            if (!string.IsNullOrEmpty(ClubPrimary.President))
            {
                if (Log_StudentDic.ContainsKey(ClubPrimary.President))
                {
                    sb_Log.AppendLine(string.Format("社長目前為「{0}」", Log_StudentDic[ClubPrimary.President].Name));
                }
            }
            if (!string.IsNullOrEmpty(ClubPrimary.VicePresident))
            {
                if (Log_StudentDic.ContainsKey(ClubPrimary.VicePresident))
                {
                    sb_Log.AppendLine(string.Format("副社長目前為「{0}」", Log_StudentDic[ClubPrimary.VicePresident].Name));
                }
            }

            List<CadresRecord> list_Insert = GetInsert(InsertList, CadresList);
            List<CadresRecord> list_ReMove = GetRemove(InsertList, CadresList);

            if (list_Insert.Count != 0)
            {
                sb_Log.AppendLine("");
                foreach (CadresRecord each in list_Insert)
                {
                    if (Log_StudentDic.ContainsKey(each.RefStudentID))
                    {
                        sb_Log.AppendLine(string.Format("新增幹部「{0}」學生「{1}」", each.CadreName, Log_StudentDic[each.RefStudentID].Name));
                    }
                }
            }

            if (list_ReMove.Count != 0)
            {
                sb_Log.AppendLine("");
                foreach (CadresRecord each in list_ReMove)
                {
                    if (Log_StudentDic.ContainsKey(each.RefStudentID))
                    {
                        sb_Log.AppendLine(string.Format("移除社團幹部「{0}」學生「{1}」", each.CadreName, Log_StudentDic[each.RefStudentID].Name));
                    }
                }
            }

            try
            {
                _AccessHelper.DeletedValues(CadresList); //刪除原有社團幹部
                _AccessHelper.InsertValues(InsertList); //新增社團幹部
                _AccessHelper.UpdateValues(list); //更新社團資料
            }
            catch (Exception ex)
            {
                MsgBox.Show("取得資料發生錯誤!!\n" + ex.Message);
            }

            FISCA.LogAgent.ApplicationLog.Log("社團", "修改幹部資料", sb_Log.ToString());
        }

        private List<CadresRecord> GetInsert(List<CadresRecord> list1, List<CadresRecord> list2)
        {
            List<CadresRecord> InsertList = new List<CadresRecord>();
            List<CadresRecord> CadresList = new List<CadresRecord>();
            foreach (CadresRecord each1 in list1) //新增的
            {
                InsertList.Add(each1);
            }
            foreach (CadresRecord each1 in list2) //移除的
            {
                CadresList.Add(each1);
            }

            List<CadresRecord> list = new List<CadresRecord>(); //一樣的資料
            foreach (CadresRecord each1 in InsertList) //新增的
            {
                foreach (CadresRecord each2 in CadresList) //原有的
                {
                    //當幹部名稱相同,擔任學生亦相同時
                    //雖然狀態是新增,但是資料顯示須判斷為未改變
                    if (each1.CadreName == each2.CadreName && each1.RefStudentID == each2.RefStudentID)
                    {
                        list.Add(each1);
                        continue;
                    }
                }
            }

            //將未改變的資料於清單內移除
            foreach (CadresRecord each1 in list)
            {
                if (InsertList.Contains(each1))
                {
                    InsertList.Remove(each1);
                }
            }

            return InsertList;
        }

        /// <summary>
        /// list1 - 新增清單
        /// list2 - 移除清單
        /// 移除的社團幹部清單,就是...
        /// 在移除清單中讀資料(CadresList),不在新增清單中(InsertList)
        /// </summary>
        private List<CadresRecord> GetRemove(List<CadresRecord> list1, List<CadresRecord> list2)
        {
            List<CadresRecord> InsertList = new List<CadresRecord>();
            List<CadresRecord> CadresList = new List<CadresRecord>();
            foreach (CadresRecord each1 in list1) //新增的
            {
                InsertList.Add(each1);
            }
            foreach (CadresRecord each1 in list2) //移除的
            {
                CadresList.Add(each1);
            }

            List<CadresRecord> list = new List<CadresRecord>(); //一樣的資料
            foreach (CadresRecord each1 in CadresList) //原有的
            {
                bool check = false;
                foreach (CadresRecord each2 in InsertList) //新增的
                {
                    //當幹部名稱相同,擔任學生亦相同時
                    //雖然狀態是新增,但是資料顯示須判斷為未改變
                    if (each1.CadreName == each2.CadreName && each1.RefStudentID == each2.RefStudentID)
                    {
                        check = true;
                    }
                }

                //有相同資料表示已被新增
                if (check)
                {
                    list.Add(each1);
                }
            }

            foreach (CadresRecord each1 in list) //原有的
            {
                if (CadresList.Contains(each1))
                {
                    CadresList.Remove(each1);
                }
            }

            return CadresList;
        }

        void Save_BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MsgBox.Show("作業中止!!");
                return;
            }

            if (e.Error != null)
            {
                MsgBox.Show("取得資料發生錯誤!!\n" + e.Error.Message);
                return;
            }

            ClubEvents.RaiseAssnChanged();
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


        void DataListener_StatusChanged(object sender, Campus.Windows.ChangeEventArgs e)
        {
            SaveButtonVisible = (e.Status == Campus.Windows.ValueStatus.Dirty);
            CancelButtonVisible = (e.Status == Campus.Windows.ValueStatus.Dirty);
        }

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridViewX1.CurrentCell.ColumnIndex == 1)
            {
                dataGridViewX1.EndEdit();
            }
        }
    }
}
