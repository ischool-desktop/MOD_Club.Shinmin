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
using Campus.Windows;
using FISCA.Permission;

namespace K12.Club.Shinmin
{
    [FISCA.Permission.FeatureCode("K12.Club.Shinmin.StudentResultItem.cs", "社團成績")]
    public partial class StudentResultItem : DetailContentBase
    {
        private BackgroundWorker BGW = new BackgroundWorker();
        private bool BkWBool = false;

        private AccessHelper _AccessHelper = new AccessHelper();

        StudentRecord _Studentecord = new StudentRecord();

        List<ResultScoreRecord> RSRList = new List<ResultScoreRecord>();

        List<ResultScoreRecord> deleteList = new List<ResultScoreRecord>();

        //權限
        internal static FeatureAce UserPermission;

        //資料變更事件引發器
        private ChangeListener DataListener { get; set; }

        public StudentResultItem()
        {
            InitializeComponent();

            Group = "社團成績";

            UserPermission = UserAcl.Current[FISCA.Permission.FeatureCodeAttribute.GetCode(GetType())];
            this.Enabled = UserPermission.Editable;

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            DataListener = new ChangeListener();
            DataListener.Add(new DataGridViewSource(dataGridViewX1));
            DataListener.StatusChanged += new EventHandler<ChangeEventArgs>(DataListener_StatusChanged);

        }

        void DataListener_StatusChanged(object sender, ChangeEventArgs e)
        {
            SaveButtonVisible = (e.Status == ValueStatus.Dirty);
            CancelButtonVisible = (e.Status == ValueStatus.Dirty);
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得學生
            _Studentecord = Student.SelectByID(this.PrimaryKey);

            //取得本名學生的社團成績記錄 string.Format("ref_student_id = '{0}'", this.PrimaryKey)
            RSRList = _AccessHelper.Select<ResultScoreRecord>(string.Format("ref_student_id = '{0}'", this.PrimaryKey));


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
                this.Loading = false;
                FISCA.Presentation.Controls.MsgBox.Show("取得[社團成績]發生錯誤!!\n" + e.Error.Message);
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

        /// <summary>
        /// 排序學期資料
        /// </summary>
        private int SortRSRList(ResultScoreRecord a1, ResultScoreRecord b2)
        {
            string aa1 = a1.SchoolYear.ToString().PadLeft(3, '0');
            aa1 += a1.Semester.ToString().PadLeft(2, '0');

            string bb2 = b2.SchoolYear.ToString().PadLeft(3, '0');
            bb2 += b2.Semester.ToString().PadLeft(2, '0');

            return aa1.CompareTo(bb2);
        }

        private void BindData()
        {
            deleteList.Clear();
            RSRList.Sort(SortRSRList);
            dataGridViewX1.Rows.Clear();

            //是否有"社團參與記錄參考",沒有就是使用者手動增加的Cell
            foreach (ResultScoreRecord each in RSRList)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridViewX1);
                if (string.IsNullOrEmpty(each.RefSCJoinID))
                {
                    row.Cells[0].Value = each.SchoolYear.ToString();
                    row.Cells[1].Value = each.Semester.ToString();
                    row.Cells[2].Value = each.ClubName;
                    row.Cells[3].Value = each.ResultScore.HasValue ? each.ResultScore.Value.ToString() : "";
                    row.Cells[4].Value = each.CadreName;
                    row.Tag = each;
                }
                else
                {
                    row.Cells[0].Value = each.SchoolYear.ToString();
                    row.Cells[1].Value = each.Semester.ToString();
                    row.Cells[2].Value = each.ClubName;
                    row.Cells[3].Value = each.ResultScore.HasValue ? each.ResultScore.Value.ToString() : "";
                    row.Cells[4].Value = each.CadreName;
                    row.Tag = each;

                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        cell.Style.BackColor = Color.LightCyan;
                        //cell.ReadOnly = true;
                    }

                }
                dataGridViewX1.Rows.Add(row);

            }

            BkWBool = false;
            SaveButtonVisible = false;
            CancelButtonVisible = false;

            DataListener.Reset();
            DataListener.ResumeListen();
        }

        /// <summary>
        /// 按下儲存時
        /// </summary>
        /// <param name="e"></param>
        protected override void OnSaveButtonClick(EventArgs e)
        {
            //資料檢查
            if (!CheckCell())
            {
                MsgBox.Show("請修正資料再儲存!!");
                return;
            }

            List<ResultScoreRecord> InsertList = new List<ResultScoreRecord>();
            List<ResultScoreRecord> updateList = new List<ResultScoreRecord>();

            //沒有Tag就是一筆新記錄
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow)
                    continue;

                if (row.Tag == null)
                {
                    ResultScoreRecord rsr = new ResultScoreRecord();
                    rsr.SchoolYear = int.Parse("" + row.Cells[0].Value);
                    rsr.Semester = int.Parse("" + row.Cells[1].Value);

                    //rsr.RefClubID = cr.UID; //社團ID
                    rsr.RefStudentID = _Studentecord.ID; //學生ID
                    //rsr.RefSCJoinID = scj.UID; //參與記錄ID

                    rsr.ClubName = "" + row.Cells[2].Value;
                    decimal xy;
                    if (decimal.TryParse("" + row.Cells[3].Value, out xy))
                    {
                        rsr.ResultScore = xy; //成績              
                    }
                    rsr.CadreName = "" + row.Cells[4].Value; //幹部

                    InsertList.Add(rsr);
                }
                else
                {
                    ResultScoreRecord rsr = (ResultScoreRecord)row.Tag;
                    rsr.SchoolYear = int.Parse("" + row.Cells[0].Value);
                    rsr.Semester = int.Parse("" + row.Cells[1].Value);

                    rsr.ClubName = "" + row.Cells[2].Value;
                    decimal xy;
                    if (decimal.TryParse("" + row.Cells[3].Value, out xy))
                    {
                        rsr.ResultScore = xy; //成績              
                    }
                    rsr.CadreName = "" + row.Cells[4].Value; //幹部

                    updateList.Add(rsr);
                }


            }

            if (InsertList.Count != 0)
                _AccessHelper.InsertValues(InsertList);
            if (updateList.Count != 0)
                _AccessHelper.UpdateValues(updateList);
            if (deleteList.Count != 0)
                _AccessHelper.DeletedValues(deleteList);

            SaveButtonVisible = false;
            CancelButtonVisible = false;

            Changed();
        }

        /// <summary>
        /// 資料檢查
        /// </summary>
        private bool CheckCell()
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                    {
                        return false;
                    }
                }
            }
            return true;
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

        private void dataGridViewX1_UserDeletedRow(object sender, DataGridViewRowEventArgs e)
        {
            if (e.Row.Tag != null)
            {
                ResultScoreRecord rsr = (ResultScoreRecord)e.Row.Tag;
                if (!deleteList.Contains(rsr))
                {
                    deleteList.Add(rsr);
                }
            }
        }

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            DataGridViewCell CurrentCell = dataGridViewX1.CurrentCell;
            if (CurrentCell.ColumnIndex == Column1.Index)
            {
                int schoolyear;
                if (int.TryParse("" + CurrentCell.Value, out schoolyear))
                {
                    CurrentCell.ErrorText = "";
                }
                else
                {
                    CurrentCell.ErrorText = "必須是數字!";
                }

            }
            else if (dataGridViewX1.CurrentCell.ColumnIndex == Column2.Index)
            {
                if ("" + CurrentCell.Value == "1" || "" + CurrentCell.Value == "2")
                {
                    CurrentCell.ErrorText = "";
                }
                else
                {
                    CurrentCell.ErrorText = "必須是1或2";
                }
            }
            else if (dataGridViewX1.CurrentCell.ColumnIndex == Column4.Index)
            {
                decimal schoolyear;
                if (decimal.TryParse("" + CurrentCell.Value, out schoolyear))
                {
                    CurrentCell.ErrorText = "";
                }
                else
                {
                    CurrentCell.ErrorText = "必須是數字!";
                }
            }



        }
    }
}
