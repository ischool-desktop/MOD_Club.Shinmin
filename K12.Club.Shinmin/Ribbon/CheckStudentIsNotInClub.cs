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
using DevComponents.DotNetBar;

namespace K12.Club.Shinmin
{
    public partial class CheckStudentIsNotInClub : BaseForm
    {
        /// <summary>
        /// 開啟畫面即檢查學生
        /// 取得本學期的社團資料
        /// </summary>
        BackgroundWorker BGW = new BackgroundWorker();

        /// <summary>
        /// 儲存學生的社團分配記錄
        /// </summary>
        BackgroundWorker BGWSave = new BackgroundWorker();

        //UDT物件
        private AccessHelper _AccessHelper = new AccessHelper();
        private QueryHelper _QueryHelper = new QueryHelper();

        List<CLUBRecord> CLUBRecordList = new List<CLUBRecord>();

        List<SCJoin> SCJoinList = new List<SCJoin>();

        /// <summary>
        /// 已選社團學生
        /// </summary>
        Dictionary<string, SCJoin> StudentScjoinDic = new Dictionary<string, SCJoin>();

        /// <summary>
        /// 未選社團學生
        /// </summary>
        List<StudRecord> IsStudentList = new List<StudRecord>();

        public CheckStudentIsNotInClub()
        {
            InitializeComponent();

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            BGWSave.DoWork += new DoWorkEventHandler(BGWSave_DoWork);
            BGWSave.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGWSave_RunWorkerCompleted);

            K12.Presentation.NLDPanels.Student.TempSourceChanged += new EventHandler(Student_TempSourceChanged);
            labelX1.Text = string.Format("{0}學年度　第{1}學期　未選社清單：", School.DefaultSchoolYear, School.DefaultSemester);
            labelX3.Text = string.Format("待處理學生：共{0}人", K12.Presentation.NLDPanels.Student.TempSource.Count);
        }

        void Student_TempSourceChanged(object sender, EventArgs e)
        {
            labelX3.Text = string.Format("待處理學生：共{0}人", K12.Presentation.NLDPanels.Student.TempSource.Count);
        }

        private void CheckStudentIsNotInClub_Load(object sender, EventArgs e)
        {
            //畫面開啟
            //1.即檢查本學年度學期,是否有未選社團學生
            //每個Row內存學生的ID

            //2.取得目前學年度/學期的可選擇社團
            //每個Row內存社團的Club UID

            //3.

            if (!BGW.IsBusy)
            {
                btnSave.Enabled = false;
                BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("請重開本畫面");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {

            //取得本學期社團資料
            CLUBRecordList.Clear();
            CLUBRecordList = _AccessHelper.Select<CLUBRecord>(string.Format("school_year={0} and semester={1}", School.DefaultSchoolYear, School.DefaultSemester));

            //取得本學期,所有社團參與記錄
            List<string> ClubRefIDList = new List<string>();
            foreach (CLUBRecord record in CLUBRecordList)
            {
                if (!ClubRefIDList.Contains(record.UID))
                {
                    ClubRefIDList.Add(record.UID);
                }
            }

            //取得學校所有學生記錄
            //學生記錄來自於社團ID
            string ClubIdString = string.Join("','", ClubRefIDList);
            List<SCJoin> Scjoin = _AccessHelper.Select<SCJoin>(string.Format("ref_club_id in ('{0}')", ClubIdString));
            foreach (SCJoin join in Scjoin)
            {
                if (!StudentScjoinDic.ContainsKey(join.RefStudentID))
                {
                    StudentScjoinDic.Add(join.RefStudentID, join);
                }
            }

            //取得學校內所有一般生記錄
            //班級/座號/學號/姓名

            DataTable studentDT = _QueryHelper.Select("select student.id,class.class_name,student.seat_no,student.student_number,student.name,class.grade_year from student join class on student.ref_class_id=class.id where student.status=1 or student.status=2 ORDER BY class.grade_year,class.class_name,student.seat_no");

            IsStudentList.Clear();
            foreach (DataRow row in studentDT.Rows)
            {
                StudRecord re = new StudRecord(row);

                if (!StudentScjoinDic.ContainsKey(re.id))
                {
                    IsStudentList.Add(re);
                }
            }

        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;
            if (e.Cancelled)
            {
                MsgBox.Show("檢查作業已停止!");
                return;
            }
            if (e.Error != null)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                MsgBox.Show("檢查學生社團參與記錄發生錯誤\n" + e.Error.Message);
                return;
            }

            #region 學生
            labelX1.Text = string.Format("{0}學年度　第{1}學期　未選社清單(共{2}人)：", School.DefaultSchoolYear, School.DefaultSemester, IsStudentList.Count);
            foreach (StudRecord re in IsStudentList)
            {
                DataGridViewRow dataRow = new DataGridViewRow();
                dataRow.CreateCells(dataGridViewX1);
                dataRow.Tag = re;
                dataRow.Cells[0].Value = re.grade_year;
                dataRow.Cells[1].Value = re.class_name;
                dataRow.Cells[2].Value = re.seat_no;
                dataRow.Cells[3].Value = re.name;
                dataRow.Cells[4].Value = re.student_number;

                dataGridViewX1.Rows.Add(dataRow);
            }
            #endregion

            #region 社團

            foreach (CLUBRecord record in CLUBRecordList)
            {
                ButtonItem btnItem = new ButtonItem();
                btnItem.Text = record.ClubName;
                btnItem.Tag = record;
                btnItem.OptionGroup = "itmPnlTimeName";
                btnItem.ButtonStyle = eButtonStyle.ImageAndText;
                btnItem.Click += new EventHandler(btnItem_Click);

                itmPnlTimeName.Items.Add(btnItem);
            }

            itmPnlTimeName.ResumeLayout();
            itmPnlTimeName.Refresh();
            #endregion
        }

        void btnItem_Click(object sender, EventArgs e)
        {
            if (itmPnlTimeName.SelectedItems.Count == 1)
            {
                //取得目前所選擇的Button
                ButtonItem Buttonitem = itmPnlTimeName.SelectedItems[0] as ButtonItem;
                //取得課程Record
                CLUBRecord club = (CLUBRecord)Buttonitem.Tag;

                foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
                {
                    row.Cells[5].Value = "" + club.ClubName;
                    row.Cells[5].Tag = club;
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            //將指定好的學生
            //建立社團參與記錄
            //並加入該社團內

            btnSave.Enabled = false;

            SCJoinList.Clear();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.Cells[5].Tag != null)
                {


                    StudRecord sr = (StudRecord)row.Tag;

                    CLUBRecord cr = (CLUBRecord)row.Cells[5].Tag;


                    SCJoin sc = new SCJoin();
                    sc.RefClubID = cr.UID;
                    sc.RefStudentID = sr.id;
                    SCJoinList.Add(sc);
                }
            }

            BGWSave.RunWorkerAsync(SCJoinList);
        }

        void BGWSave_DoWork(object sender, DoWorkEventArgs e)
        {

            List<SCJoin> SCJoinList = (List<SCJoin>)e.Argument;

            try
            {
                _AccessHelper.InsertValues(SCJoinList);
            }
            catch (Exception ex)
            {
                e.Cancel = true;
                SmartSchool.ErrorReporting.ReportingService.ReportException(ex);
                return;
            }
        }

        void BGWSave_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;

            if (e.Cancelled)
            {
                MsgBox.Show("儲存作業發生錯誤已停止!");
                return;
            }
            if (e.Error != null)
            {
                SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                MsgBox.Show("儲存作業發生錯誤!\n" + e.Error.Message);
                return;
            }

            MsgBox.Show("學生加入社團成功!!");

            ClubEvents.RaiseAssnChanged();
            this.Close();


        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void 清除指定社團ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
            {
                row.Cells[5].Value = "";
                row.Cells[5].Tag = null;
            }
        }

        private void 加入待處理學生ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            List<string> list = new List<string>();
            foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
            {
                StudRecord sr = (StudRecord)row.Tag;
                list.Add(sr.id);
            }
            K12.Presentation.NLDPanels.Student.AddToTemp(list);
        }

        private void btnOutPut_Click(object sender, EventArgs e)
        {
            #region 匯出
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "匯出未選社團學生清單";
            saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            DataGridViewExport export = new DataGridViewExport(dataGridViewX1);
            export.Save(saveFileDialog1.FileName);

            if (new CompleteForm().ShowDialog() == DialogResult.Yes)
                System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            #endregion
        }
    }

    public class StudRecord
    {
        public StudRecord(DataRow row)
        {
            id = "" + row[0];
            class_name = "" + row[1];
            seat_no = "" + row[2];
            student_number = "" + row[3];
            name = "" + row[4];
            grade_year = "" + row[5];
        }

        public string id { get; set; }
        public string class_name { get; set; }
        public string seat_no { get; set; }
        public string student_number { get; set; }
        public string name { get; set; }
        public string grade_year { get; set; }

    }
}
