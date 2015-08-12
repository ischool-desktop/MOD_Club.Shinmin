using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using K12.Data;

namespace K12.Club.Shinmin
{
    class ScJoinMag
    {
        private AccessHelper _AccessHelper = new AccessHelper();

        /// <summary>
        /// 社團ID - 社團參與記錄
        /// </summary>
        public Dictionary<string, List<SCJoin>> all_SCJoinDic = new Dictionary<string, List<SCJoin>>();

        /// <summary>
        /// 社團參與記錄清單
        /// </summary>
        public List<SCJoin> SCJoin_LIst = new List<SCJoin>();

        /// <summary>
        /// 社團參與記錄清單(ID:Record)
        /// </summary>
        public Dictionary<string, List<SCJoin>> SCJoin_Dic = new Dictionary<string, List<SCJoin>>();

        /// <summary>
        /// 社團參與記錄清單(學生記錄)
        /// </summary>
        public List<StudentRecord> SCJoinStudent_LIst = new List<StudentRecord>();

        public Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();

        public Dictionary<string, StudentRecord> StudentAllDic = new Dictionary<string, StudentRecord>();

        public List<string> SCJoin_Lock = new List<string>();

        /// <summary>
        /// 提供多種取得[社團學生參與記錄]的方法
        /// </summary>
        public ScJoinMag()
        {
            //取得目前學年期,所有社團參與記錄




        }

        /// <summary>
        /// 傳入單一社團ID
        /// 本功能用以取得社團學生清單
        /// </summary>
        public ScJoinMag(string ClubRecordID)
        {
            //取得參與記錄
            StringBuilder sb = new StringBuilder();
            sb.Append(string.Format("ref_club_id='{0}'", ClubRecordID));
            SCJoin_LIst = _AccessHelper.Select<SCJoin>(sb.ToString());

            string Test = "";

            //取得學生物件
            List<string> list = new List<string>();
            foreach (SCJoin each in SCJoin_LIst)
            {
                if (!list.Contains(each.RefStudentID))
                {
                    list.Add(each.RefStudentID);
                }

                if (!SCJoin_Dic.ContainsKey(each.RefStudentID))
                {
                    SCJoin_Dic.Add(each.RefStudentID,new List<SCJoin>());
                }

                SCJoin_Dic[each.RefStudentID].Add(each);

                if (each.Lock)
                {
                    if (!SCJoin_Lock.Contains(each.RefStudentID))
                    {
                        SCJoin_Lock.Add(each.RefStudentID);
                    }
                }
            }

            SCJoinStudent_LIst = Student.SelectByIDs(list);
            foreach (StudentRecord each in SCJoinStudent_LIst)
            {
                if (tool.CheckStatus(each))
                {
                    if (!StudentDic.ContainsKey(each.ID))
                    {
                        StudentDic.Add(each.ID, each);
                    }
                }

                if (!StudentAllDic.ContainsKey(each.ID))
                {
                    StudentAllDic.Add(each.ID, each);
                }
            }
            //排序
            SCJoinStudent_LIst = SortClassIndex.K12Data_StudentRecord(SCJoinStudent_LIst);
        }

        /// <summary>
        /// 傳入社團ID清單
        /// 取得每個課社團的參與記錄
        /// </summary>
        public ScJoinMag(List<string> ClubRecordList)
        {
            all_SCJoinDic.Clear();

            StringBuilder sb = new StringBuilder();
            string StringList = string.Join(",", ClubRecordList);
            sb.Append(string.Format("ref_club_id in ('{0}')", StringList));
            List<SCJoin> SCJoinLIst = _AccessHelper.Select<SCJoin>(sb.ToString());

            foreach (SCJoin each in SCJoinLIst)
            {
                if (all_SCJoinDic.ContainsKey(each.RefClubID))
                    continue;

                all_SCJoinDic.Add(each.RefClubID, new List<SCJoin>());
                all_SCJoinDic[each.RefClubID].Add(each);
            }
        }
    }
}
