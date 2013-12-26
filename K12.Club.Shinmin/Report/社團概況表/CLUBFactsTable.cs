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
using FISCA.Data;
using Aspose.Words;
using System.IO;
using System.Diagnostics;
using System.Xml;
using FISCA.DSAUtil;

namespace K12.Club.Shinmin
{
    public partial class CLUBFactsTable : BaseForm
    {
        //背景模式
        BackgroundWorker BGW = new BackgroundWorker();

        AccessHelper _accessHelper = new AccessHelper();

        QueryHelper _queryHelper = new QueryHelper();


        Document _doc = new Document(); //主文件
        Run _run; //移動使用

        int _SchoolYear { get; set; }
        int _Semester { get; set; }


        public CLUBFactsTable()
        {
            InitializeComponent();
        }

        private void CLUBFactsTable_Load(object sender, EventArgs e)
        {
            intSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            intSemester.Value = int.Parse(K12.Data.School.DefaultSemester);

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!BGW.IsBusy)
            {
                btnSave.Enabled = false;
                _SchoolYear = intSchoolYear.Value;
                _Semester = intSemester.Value;

                BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("系統忙碌中,稍後再試...");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //社團ID : 社團特殊物件
            Dictionary<string, List<FactsObj_big>> DoWorkDic = GetClubData(); //分類

            Document _template = new Document(new MemoryStream(Properties.Resources.社團概況表_範本));
            _doc.Sections.Clear();

            //取得範本樣式
            Document PageOne = (Document)_template.Clone(true);

            #region 資料MailMerge第一步

            List<string> name = new List<string>();
            List<string> value = new List<string>();
            name.Add("學校名稱");
            value.Add(School.ChineseName);

            name.Add("學年度");
            value.Add("" + _SchoolYear);

            name.Add("學期");
            value.Add("" + _Semester);

            name.Add("日期");
            value.Add(DateTime.Now.ToString("yyyy/MM/dd HH:mm"));

            PageOne.MailMerge.Execute(name.ToArray(), value.ToArray());

            #endregion

            _run = new Run(PageOne);
            int count = 0;
            int count_end = 0; //行數
            foreach (string each in DoWorkDic.Keys)
            {
                count_end += DoWorkDic[each].Count;
            }

            DocumentBuilder builder = new DocumentBuilder(PageOne);
            builder.MoveToMergeField("標1");
            Cell cell = (Cell)builder.CurrentParagraph.ParentNode;
            //取得目前Row
            Row row = (Row)builder.CurrentParagraph.ParentNode.ParentNode;

            //建立新行(依異動筆數)
            for (int x = 1; x < count_end; x++)
            {
                (cell.ParentNode.ParentNode as Table).InsertAfter(row.Clone(true), cell.ParentNode);
            }

            foreach (string each in DoWorkDic.Keys)
            {
                DoWorkDic[each].Sort(SortClub); //排序

                foreach (FactsObj_big bin in DoWorkDic[each])
                {
                    count++; //目前行數

                    //類別
                    Write(cell, bin.社團類型);
                    cell = GetMoveRightCell(cell, 1);

                    //代碼
                    Write(cell, bin.社團代碼);
                    cell = GetMoveRightCell(cell, 1);

                    //名稱
                    Write(cell, bin.社團名稱);
                    cell = GetMoveRightCell(cell, 1);

                    //老師
                    Write(cell, bin.老師姓名);
                    cell = GetMoveRightCell(cell, 1);

                    //限制                   
                    Write(cell, bin.科別限制);
                    cell = GetMoveRightCell(cell, 1);

                    //性別限制
                    Write(cell, bin.性別限制);
                    cell = GetMoveRightCell(cell, 1);

                    //1年級人數/目前人數
                    Write(cell, bin.一年級);
                    cell = GetMoveRightCell(cell, 1);

                    //2年級人數/目前人數
                    Write(cell, bin.二年級);
                    cell = GetMoveRightCell(cell, 1);

                    //3年級人數/目前人數
                    Write(cell, bin.三年級);
                    cell = GetMoveRightCell(cell, 1);

                    //總人數/人數上限
                    Write(cell, bin.總人數);
                    cell = GetMoveRightCell(cell, 1);

                    //場地
                    Write(cell, bin.活動場地);
                    GetMoveRightCell(cell, 1);

                    //下一行
                    Row row_k = cell.ParentRow.NextSibling as Row;
                    //第一格
                    if (row_k != null)
                    {
                        if (row_k.FirstCell != null)
                        {
                            cell = row_k.FirstCell;
                        }
                        else
                            break;
                    }
                    else
                    {
                        break;
                    }

                }
            }
            cell.ParentRow.RowFormat.Borders.Bottom.LineWidth = 1.5; //線寬

            _doc.Sections.Add(_doc.ImportNode(PageOne.FirstSection, true));
            e.Result = _doc;

        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;

            if (e.Cancelled)
            {
                MsgBox.Show("列印作業已中止!!");
                return;
            }

            if (e.Error == null)
            {
                Document inResult = (Document)e.Result;

                try
                {
                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();

                    SaveFileDialog1.Filter = "Word (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                    SaveFileDialog1.FileName = "社團概況表";

                    if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        inResult.Save(SaveFileDialog1.FileName);
                        Process.Start(SaveFileDialog1.FileName);
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
                    return;
                }

                this.Close();
            }
            else
            {
                MsgBox.Show("列印發生錯誤:\n" + e.Error.Message);
                return;
            }
        }

        int SortClub(FactsObj_big a, FactsObj_big b)
        {
            return a.社團代碼.CompareTo(b.社團代碼);
        }

        /// <summary>
        /// 取得社團相關資料
        /// </summary>
        private Dictionary<string, List<FactsObj_big>> GetClubData()
        {
            //社團ID : 社團特殊物件
            Dictionary<string, List<FactsObj_big>> dic1 = new Dictionary<string, List<FactsObj_big>>();

            //社團ID : 社團特殊物件
            Dictionary<string, List<FactsObj_big>> dic_Sort = new Dictionary<string, List<FactsObj_big>>();

            //取得本學年度的社團清單
            List<CLUBRecord> CLUBRecordList = _accessHelper.Select<CLUBRecord>(string.Format("school_year={0} and semester={1}", _SchoolYear, _Semester));

            #region 取得學生用

            List<string> CLUBIDList = new List<string>();

            foreach (CLUBRecord each in CLUBRecordList)
            {
                CLUBIDList.Add(each.UID);
            }

            List<SCJoin> SCJoinList = _accessHelper.Select<SCJoin>("ref_club_id in ('" + string.Join("','", CLUBIDList) + "')");

            List<string> StudentIDList = new List<string>();

            //社團 : 參與記錄
            Dictionary<string, List<SCJoin>> SCJDic = new Dictionary<string, List<SCJoin>>();

            foreach (SCJoin each in SCJoinList)
            {
                StudentIDList.Add(each.RefStudentID);

                if (!SCJDic.ContainsKey(each.RefClubID))
                {
                    SCJDic.Add(each.RefClubID, new List<SCJoin>());
                }
                SCJDic[each.RefClubID].Add(each);
            }

            List<StudentRecord> StudentList = Student.SelectByIDs(StudentIDList);
            Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();
            foreach (StudentRecord each in StudentList)
            {
                if (!StudentDic.ContainsKey(each.ID))
                {
                    StudentDic.Add(each.ID, each);
                }

            }

            #endregion

            //取得教師資料
            Dictionary<string, TeacherCrk> TeacherDic = GetTeacher();

            //排序用
            List<string> CLUBNameSortList = new List<string>();

            foreach (CLUBRecord each in CLUBRecordList)
            {
                FactsObj_big F = new FactsObj_big(each);

                if (SCJDic.ContainsKey(each.UID))
                {
                    foreach (SCJoin scj in SCJDic[each.UID])
                    {
                        StudentRecord SR = StudentDic[scj.RefStudentID];
                        if (SR.Status == StudentRecord.StudentStatus.一般 || SR.Status == StudentRecord.StudentStatus.延修)
                        {
                            if (SR.Class != null)
                            {
                                if (SR.Class.GradeYear.HasValue)
                                {
                                    if (SR.Class.GradeYear.Value == 1)
                                    {
                                        F.StudentList_1.Add(SR);
                                    }
                                    else if (SR.Class.GradeYear.Value == 2)
                                    {
                                        F.StudentList_2.Add(SR);
                                    }
                                    else if (SR.Class.GradeYear.Value == 3)
                                    {
                                        F.StudentList_3.Add(SR);
                                    }
                                }
                            }
                        }
                    }
                }

                //是否有老師
                if (!string.IsNullOrEmpty(each.RefTeacherID))
                {
                    if (TeacherDic.ContainsKey(each.RefTeacherID))
                    {
                        F.Teacher = TeacherDic[each.RefTeacherID]; //社團老師
                    }
                }

                //資料盒建立
                if (!dic1.ContainsKey(each.ClubCategory + "_" + each.ClubNumber))
                {
                    dic1.Add(each.ClubCategory + "_" + each.ClubNumber, new List<FactsObj_big>());
                    CLUBNameSortList.Add(each.ClubCategory + "_" + each.ClubNumber);
                }

                dic1[each.ClubCategory + "_" + each.ClubNumber].Add(F);

            }



            //排序資料
            CLUBNameSortList.Sort();
            foreach (string each in CLUBNameSortList)
            {
                dic_Sort.Add(each, dic1[each]);
            }

            return dic_Sort;
        }

        /// <summary>
        /// 取得教師資料
        /// </summary>
        private Dictionary<string, TeacherCrk> GetTeacher()
        {
            Dictionary<string, TeacherCrk> dic = new Dictionary<string, TeacherCrk>();
            string teacherUyery = "select id,teacher_name,nickname from teacher where status='1'";
            DataTable dt = _queryHelper.Select(teacherUyery);
            foreach (DataRow row in dt.Rows)
            {
                TeacherCrk obj = new TeacherCrk(row);
                string TeacherID = "" + row[0];

                if (!dic.ContainsKey(TeacherID))
                {
                    dic.Add(TeacherID, obj);
                }
            }
            return dic;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        /// <summary>
        /// 寫入資料
        /// </summary>
        private void Write(Cell cell, string text)
        {
            if (cell.FirstParagraph == null)
                cell.Paragraphs.Add(new Paragraph(cell.Document));
            cell.FirstParagraph.Runs.Clear();
            _run.Text = text;
            _run.Font.Size = 8;
            _run.Font.Name = "標楷體";
            cell.FirstParagraph.Runs.Add(_run.Clone(true));
        }

        /// <summary>
        /// 以Cell為基準,向右移一格
        /// </summary>
        private Cell GetMoveRightCell(Cell cell, int count)
        {
            if (count == 0) return cell;

            Row row = cell.ParentRow;
            int col_index = row.IndexOf(cell);
            Table table = row.ParentTable;
            int row_index = table.Rows.IndexOf(row);

            try
            {
                return table.Rows[row_index].Cells[col_index + count];
            }
            catch (Exception ex)
            {
                return null;
            }
        }

    }
}
