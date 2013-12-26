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
using FISCA.Data;
using K12.Data;

namespace K12.Club.Shinmin
{
    public partial class RepeatForm : BaseForm
    {
        BackgroundWorker BGW = new BackgroundWorker();

        AccessHelper _AccessHelper = new AccessHelper();
        QueryHelper _QueryHelper = new QueryHelper();

        Dictionary<string, CLUBRecord> ClubDic { get; set; }
        Dictionary<string, StudentRecord> StudentDic { get; set; }
        Dictionary<string, StudRepeatObj> DataDic { get; set; }

        public RepeatForm()
        {
            InitializeComponent();

            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
        }

        private void RepeatForm_Load(object sender, EventArgs e)
        {
            btnClear.Enabled = false;
            this.Text = "重覆選社檢查(資料讀取中...)";
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            ClubDic = GetClubDic();

            List<SCJoin> SCJoin = _AccessHelper.Select<SCJoin>(UDT_S.PopOneCondition("ref_club_id", ClubDic.Values.ToList().Select(X => X.UID).ToList()));

            StudentDic = GetStudentDic(SCJoin);

            //主要資料內容
            DataDic = new Dictionary<string, StudRepeatObj>();

            //建立學生容器
            foreach (SCJoin each in SCJoin)
            {
                if (!DataDic.ContainsKey(each.RefStudentID))
                {
                    if (StudentDic.ContainsKey(each.RefStudentID))
                    {
                        StudRepeatObj sr = new StudRepeatObj(StudentDic[each.RefStudentID]);
                        sr._SCJoinList.Add(each);

                        DataDic.Add(each.RefStudentID, sr);
                    }
                }
                else
                {
                    DataDic[each.RefStudentID]._SCJoinList.Add(each);
                }
            }


            //依據資料內容呈現在畫面上

            List<StudRepeatObj> list = new List<StudRepeatObj>();
            foreach (string each in DataDic.Keys)
            {
                if (DataDic[each]._SCJoinList.Count > 1)
                {
                    list.Add(DataDic[each]);
                }
            }

            e.Result = list;

        }

        /// <summary>
        /// 取得社團記錄
        /// </summary>
        private Dictionary<string, CLUBRecord> GetClubDic()
        {
            Dictionary<string, CLUBRecord> dic = new Dictionary<string, CLUBRecord>();
            List<CLUBRecord> ClubList = _AccessHelper.Select<CLUBRecord>("school_year=" + School.DefaultSchoolYear + " and semester=" + School.DefaultSemester);
            foreach (CLUBRecord each in ClubList)
            {
                if (!dic.ContainsKey(each.UID))
                {
                    dic.Add(each.UID, each);
                }
            }
            return dic;
        }

        /// <summary>
        /// 由社團參與記錄,取得學生記錄
        /// </summary>
        private Dictionary<string, StudentRecord> GetStudentDic(List<SCJoin> Join)
        {
            Dictionary<string, StudentRecord> dic = new Dictionary<string, StudentRecord>();
            List<string> studentIDList = new List<string>();
            foreach (SCJoin each in Join)
            {
                if (!studentIDList.Contains(each.RefStudentID))
                {
                    studentIDList.Add(each.RefStudentID);
                }
            }

            List<StudentRecord> StudentRecordList = Student.SelectByIDs(studentIDList);
            foreach (StudentRecord each in StudentRecordList)
            {
                if (tool.CheckStatus(each))
                {
                    if (!dic.ContainsKey(each.ID))
                    {
                        dic.Add(each.ID, each);
                    }
                }
            }

            return dic;
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnClear.Enabled = true;
            this.Text = "重覆選社檢查";

            if (e.Cancelled)
            {
                MsgBox.Show("作業已被中止");
            }
            else
            {
                if (e.Error == null)
                {
                    List<StudRepeatObj> list = (List<StudRepeatObj>)e.Result;
                    dataGridViewX1.AutoGenerateColumns = false;
                    dataGridViewX1.DataSource = list;

                    SetColumn(GetColumn());


                    foreach (DataGridViewRow row in dataGridViewX1.Rows)
                    {
                        int columnindex = 4;
                        //取得工作物件
                        StudRepeatObj obj = (StudRepeatObj)row.DataBoundItem;

                        foreach (SCJoin each in obj._SCJoinList)
                        {
                            row.Cells[columnindex].Value = ClubDic[each.RefClubID].ClubName;
                            obj._SCJionDIc.Add(columnindex, each);
                            columnindex++;
                        }
                    }
                }
                else
                {
                    MsgBox.Show("資料取得發生錯誤");
                }
            }
        }

        /// <summary>
        /// 依據傳入數值,
        /// </summary>
        private void SetColumn(int ColumnCount)
        {
            for (int x = 1; x <= ColumnCount; x++)
            {
                DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn();
                column.HeaderText = "社團" + x;
                column.Name = "CLUB" + x;
                column.Width = 100;
                column.ReadOnly = true;
                dataGridViewX1.Columns.Add(column);
            }
        }

        /// <summary>
        /// 取得目前重覆社團最大數
        /// </summary>
        private int GetColumn()
        {
            int count = 0;
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                //取得工作物件
                StudRepeatObj obj = (StudRepeatObj)row.DataBoundItem;
                if (count < obj._SCJoinList.Count)
                {
                    count = obj._SCJoinList.Count;
                }
            }
            return count;
        }

        private void dataGridViewX1_KeyDown(object sender, KeyEventArgs e)
        {
            //當使用者鍵入的是Del 或是 空白鍵
            if (e.KeyData == Keys.Delete || e.KeyData == Keys.Space)
            {
                foreach (DataGridViewCell cell in dataGridViewX1.SelectedCells)
                {
                    if (cell.ColumnIndex > 3)
                    {
                        StudRepeatObj obj = (StudRepeatObj)cell.OwningRow.DataBoundItem;
                        obj.Change = true;
                        if (obj._SCJionDIc.ContainsKey(cell.ColumnIndex))
                        {
                            SCJoin scj = obj._SCJionDIc[cell.ColumnIndex];
                            if (!obj._RemoveList.Contains(scj))
                            {
                                obj._RemoveList.Add(scj);
                            }
                        }
                        cell.Style.BackColor = Color.DarkGray;
                        cell.Style.ForeColor = Color.Red;
                    }
                }
            }
        }
        private void btnClear_Click(object sender, EventArgs e)
        {
            DialogResult dr = MsgBox.Show("您確定要把標記之社團參與記錄移除?", MessageBoxButtons.YesNo, MessageBoxDefaultButton.Button2);
            if (dr == System.Windows.Forms.DialogResult.Yes)
            {
                List<SCJoin> RemoveList = new List<SCJoin>();
                foreach (DataGridViewRow row in dataGridViewX1.Rows)
                {
                    StudRepeatObj obj = (StudRepeatObj)row.DataBoundItem;
                    if (obj.Change) //如果資料有修改
                    {
                        RemoveList.AddRange(obj._RemoveList);
                    }
                }
                try
                {
                    _AccessHelper.DeletedValues(RemoveList);
                }
                catch (Exception ex)
                {
                    MsgBox.Show("清除作業發生錯誤\n" + ex.Message);
                    return;
                }
                MsgBox.Show("社團參與記錄清除成功!!");
                this.Close();
            }
            else
            {
                MsgBox.Show("已中止操作!!");
            }

        }

        private void dataGridViewX1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            DataGridViewCell cell = dataGridViewX1.CurrentCell;

            if (cell == null)
                return;

            if (cell.ColumnIndex <= 3)
                return;

            StudRepeatObj obj = (StudRepeatObj)cell.OwningRow.DataBoundItem;

            if (obj._SCJionDIc.ContainsKey(cell.ColumnIndex))
            {
                SCJoin scj = obj._SCJionDIc[cell.ColumnIndex];

                if (!obj._RemoveList.Contains(scj))
                {
                    obj.Change = true;
                    obj._RemoveList.Add(scj);
                    cell.Style.BackColor = Color.DarkGray;
                    cell.Style.ForeColor = Color.Red;
                }
                else
                {
                    obj.Change = false;
                    obj._RemoveList.Remove(scj);
                    cell.Style.BackColor = Color.White;
                    cell.Style.ForeColor = Color.Black;
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

    }
}
