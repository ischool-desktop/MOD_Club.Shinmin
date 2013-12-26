using System.Collections.Generic;
using System.Xml;
using FISCA.UDT;
using SmartSchool.API.PlugIn;
using K12.Data;

namespace K12.Club.Shinmin
{
    class ExportStudentClubResult : SmartSchool.API.PlugIn.Export.Exporter
    {
        //建構子
        public ExportStudentClubResult()
        {
            this.Image = null;
            this.Text = "匯出社團學期成績";
        }

        //覆寫
        public override void InitializeExport(SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            List<string> FieldList = new List<string>();
            FieldList.Add("學年度"); //目前欄位
            FieldList.Add("學期"); //目前欄位
            FieldList.Add("社團名稱"); //目前欄位
            FieldList.Add("社團成績"); //目前欄位

            wizard.ExportableFields.AddRange(FieldList);

            wizard.ExportPackage += (sender, e) =>
            {
                //取得學生清單

                AccessHelper helper = new AccessHelper();

                string strCondition = string.Join("','", e.List);
                List<ResultScoreRecord> records = helper.Select<ResultScoreRecord>("ref_student_id in ('" + strCondition + "')");

                for (int i = 0; i < records.Count; i++)
                {
                    RowData row = new RowData();
                    row.ID = records[i].RefStudentID;

                    foreach (string field in e.ExportFields)
                    {
                        if (wizard.ExportableFields.Contains(field))
                        {
                            switch (field)
                            {
                                case "學年度": row.Add(field, "" + records[i].SchoolYear); break;
                                case "學期": row.Add(field, "" + records[i].Semester); break;
                                case "社團名稱": row.Add(field, records[i].ClubName); break;
                                case "社團成績": row.Add(field, records[i].ResultScore.HasValue ? records[i].ResultScore.Value.ToString() : ""); break;
                            }
                        }
                    }
                    e.Items.Add(row);
                }
            };
        }
    }
}