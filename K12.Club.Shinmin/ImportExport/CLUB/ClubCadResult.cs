using System.Collections.Generic;
using K12.Data;
using SmartSchool.API.PlugIn;
using System;
using FISCA.UDT;
using System.Text;
using FISCA.LogAgent;

namespace K12.Club.Shinmin.CLUB
{
    class ClubCadResult : SmartSchool.API.PlugIn.Export.Exporter
    {
        AccessHelper helper = new AccessHelper();

        Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();
        Dictionary<string, CLUBRecord> CLUBDic = new Dictionary<string, CLUBRecord>();
        //建構子
        public ClubCadResult()
        {
            this.Image = null;
            this.Text = "匯出社團幹部清單";
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
                List<ClubCadresObj> _startList = new List<ClubCadresObj>();
                List<CLUBRecord> clubList = helper.Select<CLUBRecord>(SelectCLUBIDList);
                CLUBDic.Clear();
                foreach (CLUBRecord each in clubList)
                {
                    if (!CLUBDic.ContainsKey(each.UID))
                    {
                        CLUBDic.Add(each.UID, each);
                    }

                    #region 處理學生擔任之幹部
                    if (!string.IsNullOrEmpty(each.President))
                    {
                        ClubCadresObj obj = new ClubCadresObj();
                        obj._Club = each;
                        obj.ref_student_id = each.President;
                        obj.CadreName = "社長";
                        _startList.Add(obj);
                    }

                    if (!string.IsNullOrEmpty(each.VicePresident))
                    {
                        ClubCadresObj obj = new ClubCadresObj();
                        obj._Club = each;
                        obj.ref_student_id = each.VicePresident;
                        obj.CadreName = "副社長";
                        _startList.Add(obj);
                    }
                    #endregion
                }
                //取得社團學生的幹部記錄
                List<CadresRecord> newList = helper.Select<CadresRecord>(string.Format("ref_club_id in ('{0}')", string.Join("','", SelectCLUBIDList)));
                foreach (CadresRecord each in newList)
                {
                    if (CLUBDic.ContainsKey(each.RefClubID))
                    {
                        ClubCadresObj obj = new ClubCadresObj();
                        obj._Club = CLUBDic[each.RefClubID];
                        obj.CadreName = each.CadreName;
                        obj.ref_student_id = each.RefStudentID;
                        _startList.Add(obj);
                    }
                }

                List<string> StudentIDList = new List<string>();
                foreach (ClubCadresObj rsr in _startList)
                {
                    if (!StudentIDList.Contains(rsr.ref_student_id))
                    {
                        StudentIDList.Add(rsr.ref_student_id);
                    }
                }

                #endregion

                #region 取得學生基本資料

                StudentDic.Clear();
                List<StudentRecord> StudentRecordList = Student.SelectByIDs(StudentIDList);
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


                foreach (ClubCadresObj Result in _startList)
                {
                    if (StudentDic.ContainsKey(Result.ref_student_id))
                    {
                        Result._Student = StudentDic[Result.ref_student_id];
                    }
                }

                #endregion

                _startList.Sort(SortResult);

                foreach (ClubCadresObj Result in _startList)
                {
                    StudentRecord sr = Result._Student;

                    //社團代碼
                    string CLUBCode = Result._Club.ClubNumber;

                    RowData row = new RowData();
                    row.ID = Result.ref_student_id;

                    foreach (string field in e.ExportFields)
                    {
                        #region row

                        if (wizard.ExportableFields.Contains(field))
                        {
                            switch (field)
                            {
                                case "學年度": row.Add(field, "" + Result._Club.SchoolYear); break;
                                case "學期": row.Add(field, "" + Result._Club.Semester); break;
                                case "社團名稱": row.Add(field, "" + Result._Club.ClubName); break;
                                case "代碼": row.Add(field, CLUBCode); break;
                                case "班級": row.Add(field, string.IsNullOrEmpty(sr.RefClassID) ? "" : sr.Class.Name); break;
                                case "座號": row.Add(field, sr.SeatNo.HasValue ? sr.SeatNo.Value.ToString() : ""); break;
                                case "學號": row.Add(field, sr.StudentNumber); break;
                                case "姓名": row.Add(field, sr.Name); break;
                                case "幹部職稱": row.Add(field, Result.CadreName); break;
                            }
                        }

                        #endregion
                    }
                    e.Items.Add(row);

                }
            };
        }

        /// <summary>
        /// 排序依據
        /// 社團代碼 / 社團名稱 / 幹部名稱 / 學生班級 / 學生座號
        /// </summary>
        private int SortResult(ClubCadresObj rsr1, ClubCadresObj rsr2)
        {
            //代碼
            string rsr1Code = rsr1._Club.ClubNumber.PadLeft(10, '0');
            string rsr2Code = rsr2._Club.ClubNumber.PadLeft(10, '0');
            //社團名稱
            rsr1Code += rsr1._Club.ClubName.PadLeft(10, '0');
            rsr2Code += rsr2._Club.ClubName.PadLeft(10, '0');

            //幹部名稱
            rsr1Code += GetCadNowName(rsr1.CadreName);
            rsr2Code += GetCadNowName(rsr2.CadreName);

            //學生班級
            rsr1Code += string.IsNullOrEmpty(rsr1._Student.RefClassID) ? "000000" : rsr1._Student.Class.Name.PadLeft(6, '0');
            rsr2Code += string.IsNullOrEmpty(rsr2._Student.RefClassID) ? "000000" : rsr2._Student.Class.Name.PadLeft(6, '0');

            //學生座號
            rsr1Code += rsr1._Student.SeatNo.HasValue ? rsr1._Student.SeatNo.Value.ToString().PadLeft(3, '0') : "000";
            rsr2Code += rsr2._Student.SeatNo.HasValue ? rsr2._Student.SeatNo.Value.ToString().PadLeft(3, '0') : "000";

            return rsr1Code.CompareTo(rsr2Code);
        }

        private string GetCadNowName(string p)
        {
            if (p == "社長")
            {
                return "01";
            }
            else if (p == "副社長")
            {
                return "02";
            }
            else
            {
                return "03";
            }
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
            list.Add("學年度");
            list.Add("學期");
            list.Add("社團名稱");
            list.Add("代碼");
            list.Add("班級");
            list.Add("座號");
            list.Add("學號");
            list.Add("姓名");
            list.Add("幹部職稱");
            return list;
        }
    }
}
