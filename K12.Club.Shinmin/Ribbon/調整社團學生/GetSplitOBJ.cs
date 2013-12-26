using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using FISCA.Data;
using System.Data;

namespace K12.Club.Shinmin
{
    static class GetSplitOBJ
    {
        //UDT物件
        private static AccessHelper _AccessHelper = new AccessHelper();
        private static QueryHelper _QueryHelper = new QueryHelper();

        /// <summary>
        /// 取得選擇社團的學生社團參與記錄
        /// </summary>
        public static Dictionary<string, SCJoin> GetSCJoin()
        {
            Dictionary<string, SCJoin> dic = new Dictionary<string, SCJoin>();
            foreach (SCJoin each in _AccessHelper.Select<SCJoin>(UDT_S.PopOneCondition("ref_club_id", ClubAdmin.Instance.SelectedSource)))
            {
                if (!dic.ContainsKey(each.UID))
                {
                    dic.Add(each.UID, each);
                }
            }
            return dic;
        }

        /// <summary>
        /// 取得選擇社團的資料
        /// </summary>
        public static Dictionary<string, CLUBRecord> GetClubRecord()
        {
            Dictionary<string, CLUBRecord> dic = new Dictionary<string, CLUBRecord>();
            foreach (CLUBRecord each in _AccessHelper.Select<CLUBRecord>(UDT_S.PopOneCondition("UID", ClubAdmin.Instance.SelectedSource)))
            {
                if (!dic.ContainsKey(each.UID))
                {
                    dic.Add(each.UID, each);
                }
            }
            return dic;
        }

        /// <summary>
        /// 根據 new_SCJoinDic 取得學生記錄ID
        /// </summary>
        public static List<string> GetStudentIDList(List<SCJoin> SCJlist)
        {
            List<string> list = new List<string>();
            foreach (SCJoin each in SCJlist)
            {
                if (!list.Contains(each.RefStudentID))
                {
                    list.Add(each.RefStudentID);
                }
            }
            return list;

        }

        /// <summary>
        /// 取得傳入學生的班級資訊/年級資訊
        /// </summary>
        public static Dictionary<string, DataRow_clsRecord> GetClassRecord(List<string> studentList)
        {
            Dictionary<string, DataRow_clsRecord> dic = new Dictionary<string, DataRow_clsRecord>();
            if (studentList.Count != 0)
            {
                StringBuilder sb_class = new StringBuilder();
                sb_class.Append("select class.id,class.class_name,class.grade_year from student ");
                sb_class.Append("join class on student.ref_class_id=class.id ");
                sb_class.Append(string.Format("where student.id in('{0}') ", string.Join("','", studentList)));
                sb_class.Append("group by class.id,class.class_name,class.grade_year");

                //需要把class去除重覆
                //studentList不可為0
                DataTable td_class = _QueryHelper.Select(sb_class.ToString());
                foreach (DataRow row in td_class.Rows)
                {
                    DataRow_clsRecord cr = new DataRow_clsRecord(row);
                    if (!dic.ContainsKey(cr.id))
                    {
                        dic.Add(cr.id, cr);
                    }
                }
            }
            return dic;
        }
    }
}
