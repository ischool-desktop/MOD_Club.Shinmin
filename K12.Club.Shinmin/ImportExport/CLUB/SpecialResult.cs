using System.Collections.Generic;
using K12.Data;
using SmartSchool.API.PlugIn;
using System;
using FISCA.UDT;
using System.Text;
using FISCA.LogAgent;
using System.Data;

namespace K12.Club.Shinmin.CLUB
{
    class SpecialResult : SmartSchool.API.PlugIn.Export.Exporter
    {
        Dictionary<string, demoStudent> StudentDic = new Dictionary<string, demoStudent>();

        //建構子
        public SpecialResult()
        {
            this.Image = null;
            this.Text = "匯出聯課活動成績(資料介接)";
        }

        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            List<string> FieldsList = GetList();
            wizard.ExportableFields.AddRange(FieldsList.ToArray());

            wizard.ExportPackage += delegate(object sender, SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
            {
                #region 收集資料

                //取得所選社團
                List<string> SelectCLUBIDList = e.List;

                //取得參與學生的社團學期成績
                List<SCJoin> SCJoinList = tool._A.Select<SCJoin>("ref_club_id in ('" + string.Join("','", SelectCLUBIDList) + "')");
                List<string> SCJoinIDList = new List<string>();
                foreach (SCJoin each in SCJoinList)
                {
                    if (!SCJoinIDList.Contains(each.UID))
                    {
                        SCJoinIDList.Add(each.UID);
                    }
                }

                List<ResultScoreRecord> ResultScoreList = tool._A.Select<ResultScoreRecord>("ref_scjoin_id in ('" + string.Join("','", SCJoinIDList) + "')");

                List<string> StudentIDList = new List<string>();

                foreach (ResultScoreRecord rsr in ResultScoreList)
                {
                    if (!StudentIDList.Contains(rsr.RefStudentID))
                    {
                        StudentIDList.Add(rsr.RefStudentID);
                    }
                }

                #endregion

                #region 取得學生基本資料

                StudentDic.Clear();
                if (StudentIDList.Count > 0)
                {
                    DataTable dt = tool._Q.Select(string.Format("select student.id,student.seat_no,student.student_number,student.name,class.class_name,class.id as class_id,class.grade_year from student join class on student.ref_class_id=class.id where student.id in('{0}')", string.Join("','", StudentIDList)));
                    foreach (DataRow each in dt.Rows)
                    {
                        demoStudent ds = new demoStudent(each);
                        if (!StudentDic.ContainsKey(ds.ref_student_id))
                        {
                            StudentDic.Add(ds.ref_student_id, ds);
                        }
                    }
                }

                #endregion

                ResultScoreList.Sort(SortResult);

                foreach (ResultScoreRecord Result in ResultScoreList)
                {
                    if (StudentDic.ContainsKey(Result.RefStudentID))
                    {
                        demoStudent sr = StudentDic[Result.RefStudentID];

                        #region 其它

                        string 取得學分 = "否";
                        if (Result.ResultScore.HasValue)
                        {
                            if (Result.ResultScore.Value >= 60)
                            {
                                取得學分 = "是";
                            }
                        }

                        string 科目級別 = "";
                        string 成績年級 = "";

                        if (!string.IsNullOrEmpty(sr.class_id))
                        {
                            int a;
                            if (int.TryParse(sr.grade_year,out a))
                            {
                                科目級別 = GetSchoolYearByGradeYear(a, Result.Semester);
                                成績年級 = sr.grade_year;
                            }

                        }

                        #endregion

                        RowData row = new RowData();
                        row.ID = Result.UID;

                        foreach (string field in e.ExportFields)
                        {
                            #region row

                            if (wizard.ExportableFields.Contains(field))
                            {
                                switch (field)
                                {
                                    case "學生系統編號": row.Add(field, sr.ref_student_id); break;
                                    case "學號": row.Add(field, sr.student_number); break;
                                    case "班級": row.Add(field, sr.class_name); break;
                                    case "座號": row.Add(field, sr.seat_no); break;
                                    case "姓名": row.Add(field, sr.name); break;

                                    case "科目": row.Add(field, "聯課活動"); break;
                                    case "科目級別": row.Add(field, 科目級別); break;
                                    case "學年度": row.Add(field, "" + Result.SchoolYear); break;
                                    case "學期": row.Add(field, "" + Result.Semester); break;
                                    case "學分數": row.Add(field, "0"); break;
                                    case "必選修": row.Add(field, "必修"); break;
                                    case "分項類別": row.Add(field, "學業"); break;
                                    case "成績年級": row.Add(field, 成績年級); break;
                                    case "校部訂": row.Add(field, "部訂"); break;
                                    case "科目成績": row.Add(field, Result.ResultScore.HasValue ? Result.ResultScore.Value.ToString() : ""); break;
                                    case "原始成績": row.Add(field, Result.ResultScore.HasValue ? Result.ResultScore.Value.ToString() : ""); break;
                                    case "取得學分": row.Add(field, 取得學分); break;
                                }
                            }

                            #endregion
                        }
                        e.Items.Add(row);
                    }
                }
            };
        }

        private int SortResult(ResultScoreRecord rsr1, ResultScoreRecord rsr2)
        {
            string rsr1SortString = "";
            string rsr2SortString = "";

            if (StudentDic.ContainsKey(rsr1.RefStudentID) && StudentDic.ContainsKey(rsr2.RefStudentID))
            {
                demoStudent sr1 = StudentDic[rsr1.RefStudentID];
                demoStudent sr2 = StudentDic[rsr2.RefStudentID];

                rsr1SortString = string.IsNullOrEmpty(sr1.class_id) ? "000000" : sr1.class_name.PadLeft(6, '0');
                rsr2SortString = string.IsNullOrEmpty(sr2.class_id) ? "000000" : sr2.class_name.PadLeft(6, '0');

                rsr1SortString += !string.IsNullOrEmpty(sr1.seat_no) ? sr1.seat_no.ToString().PadLeft(3, '0') : "000";
                rsr2SortString += !string.IsNullOrEmpty(sr2.seat_no) ? sr2.seat_no.ToString().PadLeft(3, '0') : "000";

            }

            return rsr1SortString.CompareTo(rsr2SortString);
        }

        /// <summary>
        /// 取得年級比例
        /// </summary>
        private string GetSchoolYearByGradeYear(int GradeYear, int Semester)
        {
            if (GradeYear == 1)
            {
                if (Semester == 1)
                {
                    return "1";
                }
                else if (Semester == 2)
                {
                    return "2";
                }
            }
            else if (GradeYear == 2)
            {
                if (Semester == 1)
                {
                    return "3";
                }
                else if (Semester == 2)
                {
                    return "4";
                }
            }
            else if (GradeYear == 3)
            {
                if (Semester == 1)
                {
                    return "5";
                }
                else if (Semester == 2)
                {
                    return "6";
                }
            }

            return "";
        }

        private List<string> GetList()
        {
            List<string> list = new List<string>();
            list.Add("學生系統編號");
            list.Add("學號");
            list.Add("班級");
            list.Add("座號");
            list.Add("姓名");
            list.Add("科目");
            list.Add("科目級別");
            list.Add("學年度");
            list.Add("學期");
            list.Add("學分數");
            list.Add("必選修");
            list.Add("分項類別");
            list.Add("成績年級");
            list.Add("校部訂");
            list.Add("科目成績");
            list.Add("原始成績");
            list.Add("取得學分");
            return list;
        }
    }

    class demoStudent
    {
        public demoStudent(DataRow row)
        {
            ref_student_id = "" + row["id"];
            seat_no = "" + row["seat_no"];
            class_id = "" + row["class_id"];
            class_name = "" + row["class_name"];
            grade_year = "" + row["grade_year"];
            name = "" + row["name"];
            student_number = "" + row["student_number"];
        }

        public string student_number { get; set; }
        public string grade_year { get; set; }
        public string ref_student_id { get; set; }
        public string class_id { get; set; }
        public string class_name { get; set; }
        public string seat_no { get; set; }
        public string name { get; set; }
    }
}
