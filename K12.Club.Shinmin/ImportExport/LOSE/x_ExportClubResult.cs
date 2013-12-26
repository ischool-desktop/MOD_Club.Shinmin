using System.Collections.Generic;
using System.Xml;
using FISCA.UDT;
using SmartSchool.API.PlugIn;
using K12.Data;

namespace K12.Club.Shinmin.LOSE
{
    class x_ExportClubResult : SmartSchool.API.PlugIn.Export.Exporter
    {
        //建構子
        public x_ExportClubResult()
        {
            this.Image = null;
            this.Text = "匯出聯課活動成績(社團記錄)";
        }

        //覆寫
        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            List<string> FieldList = new List<string>();
            FieldList.Add("科目"); //聯課活動
            FieldList.Add("科目級別");
            FieldList.Add("學年度"); //目前欄位
            FieldList.Add("學期"); //目前欄位
            FieldList.Add("學分數");
            FieldList.Add("必選修");
            FieldList.Add("分項類別");
            FieldList.Add("成績年級");
            FieldList.Add("校部訂");
            FieldList.Add("科目成績"); //目前欄位
            FieldList.Add("原始成績"); //目前欄位
            FieldList.Add("取得學分");

            FieldList.Add("社團名稱"); //目前欄位
            FieldList.Add("社團成績"); //目前欄位
            FieldList.Add("幹部記錄"); //目前欄位

            wizard.ExportableFields.AddRange(FieldList);

            wizard.ExportPackage += (sender, e) =>
            {
                //取得學生清單

                AccessHelper helper = new AccessHelper();

                string strCondition = string.Empty;

                foreach (string ID in e.List)
                    strCondition += strCondition == string.Empty ? "'" + ID + "'" : ",'" + ID + "'";

                List<ResultScoreRecord> records = helper.Select<ResultScoreRecord>("ref_student_id in (" + strCondition + ")");

                //取得學生學期歷程
                List<SemesterHistoryRecord> SemesterList = SemesterHistory.SelectByStudentIDs(e.List);

                //學生ID : SemesterHistoryRecord
                Dictionary<string, SemesterHistoryRecord> SemesterDic = new Dictionary<string, SemesterHistoryRecord>();
                foreach (SemesterHistoryRecord each in SemesterList)
                {
                    if (!SemesterDic.ContainsKey(each.RefStudentID))
                    {
                        SemesterDic.Add(each.RefStudentID, each);
                    }
                }

                //records.Sort();

                for (int i = 0; i < records.Count; i++)
                {

                    RowData row = new RowData();
                    row.ID = records[i].RefStudentID;

                    foreach (string field in e.ExportFields)
                    {
                        if (wizard.ExportableFields.Contains(field))
                        {
                            string 取得學分 = "否";
                            string 科目級別 = "";
                            string 成績年級 = "";

                            if (records[i].ResultScore.HasValue)
                            {
                                if (records[i].ResultScore.Value >= 60)
                                {
                                    取得學分 = "是";
                                }
                            }

                            if (SemesterDic.ContainsKey(records[i].RefStudentID))
                            {
                                SemesterHistoryRecord shr = SemesterDic[records[i].RefStudentID];

                                foreach (SemesterHistoryItem each in shr.SemesterHistoryItems)
                                {
                                    if (records[i].SchoolYear == each.SchoolYear && records[i].Semester == each.Semester)
                                    {
                                        科目級別 = GetSchoolYearByGradeYear(each);
                                        成績年級 = each.GradeYear.ToString();
                                    }
                                }
                            }

                            switch (field)
                            {
                                case "科目": row.Add(field, "聯課活動"); break;

                                //需要依據學期歷程進行判斷
                                case "科目級別": row.Add(field, 科目級別); break;
                                case "學年度": row.Add(field, "" + records[i].SchoolYear); break;
                                case "學期": row.Add(field, "" + records[i].Semester); break;
                                case "學分數": row.Add(field, "0"); break;
                                case "必選修": row.Add(field, "必修"); break;
                                case "分項類別": row.Add(field, "學業"); break;
                                case "成績年級": row.Add(field, 成績年級); break;
                                case "校部訂": row.Add(field, "部訂"); break;
                                case "科目成績": row.Add(field, records[i].ResultScore.HasValue ? records[i].ResultScore.Value.ToString() : ""); break;
                                case "原始成績": row.Add(field, records[i].ResultScore.HasValue ? records[i].ResultScore.Value.ToString() : ""); break;
                                case "取得學分": row.Add(field, 取得學分); break;

                                case "社團名稱": row.Add(field, records[i].ClubName); break;
                                case "社團成績": row.Add(field, records[i].ResultScore.HasValue ? records[i].ResultScore.Value.ToString() : ""); break;
                                case "幹部記錄": row.Add(field, records[i].CadreName); break;
                            }
                        }
                    }
                    e.Items.Add(row);
                }
            };
        }

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
    }
}