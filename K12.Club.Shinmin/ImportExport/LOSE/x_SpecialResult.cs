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
using Aspose.Cells;
using System.Diagnostics;
using FISCA.Presentation;

namespace K12.Club.Shinmin.LOSE
{
    public partial class x_SpecialResult : BaseForm
    {
        AccessHelper helper = new AccessHelper();

        BackgroundWorker BGW = new BackgroundWorker();

        List<string> TitleList = new List<string>();

        public x_SpecialResult()
        {
            InitializeComponent();
        }

        private void SpecialResult_Load(object sender, EventArgs e)
        {
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!BGW.IsBusy)
            {
                btnSave.Enabled = false;
                BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("忙碌中稍後再試...");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得所選社團
            List<string> SelectCLUBIDList = ClubAdmin.Instance.SelectedSource;

            //取得參與學生的社團學期成績
            List<ResultScoreRecord> ResultList = helper.Select<ResultScoreRecord>("ref_club_id in ('" + string.Join("','", SelectCLUBIDList) + "')");
            List<string> StudentIDList = new List<string>();

            foreach (ResultScoreRecord rsr in ResultList)
            {
                if (!StudentIDList.Contains(rsr.RefStudentID))
                {
                    StudentIDList.Add(rsr.RefStudentID);
                }
            }

            #region 取得學生基本資料

            List<StudentRecord> StudentRecordList = Student.SelectByIDs(StudentIDList);
            Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();
            foreach (StudentRecord each in StudentRecordList)
            {
                if (each.Status == StudentRecord.StudentStatus.一般 || each.Status == StudentRecord.StudentStatus.延修)
                {
                    if (!StudentDic.ContainsKey(each.ID))
                    {
                        StudentDic.Add(each.ID, each);
                    }
                }
            } 

            #endregion

            #region 學期歷程

            List<SemesterHistoryRecord> SemesterList = SemesterHistory.SelectByStudentIDs(StudentIDList);

            //學生ID : SemesterHistoryRecord 
            Dictionary<string, SemesterHistoryRecord> SemesterDic = new Dictionary<string, SemesterHistoryRecord>();
            foreach (SemesterHistoryRecord each in SemesterList)
            {
                if (!SemesterDic.ContainsKey(each.RefStudentID))
                {
                    SemesterDic.Add(each.RefStudentID, each);
                }
            } 

            #endregion

            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            //建立標頭
            TitleList = GetTitle();
            int TitleIndex = 0;
            foreach (string each in TitleList)
            {
                ws.Cells[0, TitleIndex].PutValue(each);
                TitleIndex++;
            }

            int RowIndex = 1;
            int ColumnIndex = 0;
            foreach (ResultScoreRecord Result in ResultList)
            {
                if (StudentDic.ContainsKey(Result.RefStudentID))
                {
                    #region 每筆資料

                    StudentRecord sr = StudentDic[Result.RefStudentID];
                    ColumnIndex = 0;

                    ws.Cells[RowIndex, ColumnIndex].PutValue(sr.ID);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(sr.StudentNumber);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(string.IsNullOrEmpty(sr.RefClassID) ? "" : sr.Class.Name);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(sr.SeatNo.HasValue ? sr.SeatNo.Value.ToString() : "");
                    ColumnIndex++;
                    //ws.Cells[RowIndex, ColumnIndex].PutValue("");
                    //ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(sr.Name);
                    ColumnIndex++;

                    #region 其它

                    string 取得學分 = "是";
                    string 科目級別 = "";
                    string 成績年級 = "";

                    //if (Result.ResultScore.HasValue)
                    //{
                    //    if (Result.ResultScore.Value >= 60)
                    //    {
                    //        取得學分 = "是";
                    //    }
                    //}

                    if (SemesterDic.ContainsKey(Result.RefStudentID))
                    {
                        SemesterHistoryRecord shr = SemesterDic[Result.RefStudentID];

                        foreach (SemesterHistoryItem each in shr.SemesterHistoryItems)
                        {
                            if (Result.SchoolYear == each.SchoolYear && Result.Semester == each.Semester)
                            {
                                科目級別 = GetSchoolYearByGradeYear(each);
                                成績年級 = each.GradeYear.ToString();
                            }
                        }
                    }

                    #endregion

                    ws.Cells[RowIndex, ColumnIndex].PutValue("聯課活動");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(科目級別);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue("" + Result.SchoolYear);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue("" + Result.Semester);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue("0");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue("必修");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue("學業");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(成績年級);
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue("部訂");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(Result.ResultScore.HasValue ? Result.ResultScore.Value.ToString() : "");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(Result.ResultScore.HasValue ? Result.ResultScore.Value.ToString() : "");
                    ColumnIndex++;
                    ws.Cells[RowIndex, ColumnIndex].PutValue(取得學分);
                    ColumnIndex++;

                    RowIndex++;

                    #endregion
                }
            }

            e.Result = wb;
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;

            if (e.Cancelled)
            {
                MsgBox.Show("作業已被中止!!");
            }
            else
            {
                if (e.Error == null)
                {
                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                    SaveFileDialog1.Filter = "Excel (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                    SaveFileDialog1.FileName = "匯出資料介接成績";

                    //資料
                    try
                    {
                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            Workbook inResult = (Workbook)e.Result;
                            inResult.Save(SaveFileDialog1.FileName);
                            Process.Start(SaveFileDialog1.FileName);
                            MotherForm.SetStatusBarMessage("匯出資料介接成績,列印完成!!");
                        }
                        else
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                            return;
                        }
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                        MotherForm.SetStatusBarMessage("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                    }
                }
                else
                {
                    MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                    SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                }
            }

        }

        /// <summary>
        /// 取得年級比例
        /// </summary>
        private string GetSchoolYearByGradeYear(SemesterHistoryItem item)
        {
            if (item.GradeYear == 1)
            {
                if (item.Semester == 1)
                {
                    return "1";
                }
                else if (item.Semester == 2)
                {
                    return "2";
                }
            }
            else if (item.GradeYear == 2)
            {
                if (item.Semester == 1)
                {
                    return "3";
                }
                else if (item.Semester == 2)
                {
                    return "4";
                }
            }
            else if (item.GradeYear == 3)
            {
                if (item.Semester == 1)
                {
                    return "5";
                }
                else if (item.Semester == 2)
                {
                    return "6";
                }
            }

            return "";
        }

        /// <summary>
        /// 取得報表標題清單
        /// </summary>
        private List<string> GetTitle()
        {
            List<string> list = new List<string>();

            list.Add("學生系統編號"); //聯課活動
            list.Add("學號");
            list.Add("班級"); //目前欄位
            list.Add("座號"); //目前欄位
            //list.Add("科別");
            list.Add("姓名");

            list.Add("科目"); //聯課活動
            list.Add("科目級別");
            list.Add("學年度"); //目前欄位
            list.Add("學期"); //目前欄位
            list.Add("學分數");
            list.Add("必選修");
            list.Add("分項類別");
            list.Add("成績年級");
            list.Add("校部訂");
            list.Add("科目成績"); //目前欄位
            list.Add("原始成績"); //目前欄位
            list.Add("取得學分");

            return list;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
