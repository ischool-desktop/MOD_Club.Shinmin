using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using K12.Data;

namespace K12.Club.Shinmin
{
    class SCJoinDataLoad
    {

        AccessHelper _AccessHelper = new AccessHelper();

        /// <summary>
        /// 目前畫面上所選擇的"社團清單"
        /// </summary>
        List<CLUBRecord> ClubList { get; set; }

        /// <summary>
        /// 目前畫面上所選擇"社團清單"
        /// 該社團清單的所有社團學生參與記錄
        /// </summary>
        List<SCJoin> SCJoinList { get; set; }

        /// <summary>
        /// 社團參與記錄(學生ID)清單
        /// </summary>
        List<string> StudentIDList { get; set; }

        /// <summary>
        /// 學生記錄
        /// </summary>
        List<StudentRecord> StudentRecordList { get; set; }

        /// <summary>
        /// 學生資料
        /// </summary>
        public Dictionary<string, StudentRecord> StudentRecordDic { get; set; }

        /// <summary>
        /// 老師資料
        /// </summary>
        public Dictionary<string, TeacherRecord> TeacherDic { get; set; }

        /// <summary>
        /// 社團資料
        /// </summary>
        public Dictionary<string, CLUBRecord> CLUBRecordDic { get; set; }

        /// <summary>
        /// 列印報表資料用
        /// 社團ID:學生清單
        /// </summary>
        public Dictionary<string, List<StudentRecord>> ClubByStudentList { get; set; }

        public SCJoinDataLoad()
        {
            //取得社團清單
            ClubList = _AccessHelper.Select<CLUBRecord>("UID in ('" + string.Join("','", ClubAdmin.Instance.SelectedSource) + "')");
            //取得學生修課記錄
            SCJoinList = _AccessHelper.Select<SCJoin>("ref_club_id in ('" + string.Join("','", ClubAdmin.Instance.SelectedSource) + "')");

            //取得學生ID - 1
            //StudentIDList = SCJoinList.Select(x => x.RefStudentID).ToList();
            //取得學生ID - 2
            StudentIDList = SCJoinSetByStudentIDList();

            //取得學生Record
            StudentRecordList = K12.Data.Student.SelectByIDs(StudentIDList);

            //<學生ID:學生Record>
            StudentRecordDic = GetStudentDic();

            //<社團ID:社團Record>
            CLUBRecordDic = GetCLUBDic();

            //<社團ID:學生清單>
            ClubByStudentList = GetClubIDByStudentDic();

            //依據修課記錄,加入學生
            foreach (SCJoin each in SCJoinList)
            {
                if (CLUBRecordDic.ContainsKey(each.RefClubID))
                {
                    if (StudentRecordDic.ContainsKey(each.RefStudentID))
                    {
                        ClubByStudentList[each.RefClubID].Add(StudentRecordDic[each.RefStudentID]);
                    }
                }
            }

            //把學生清單進行排序
            ClubByStudentList = SortStudentDic(ClubByStudentList);

            //取得社團老師
            List<string> TeacherIDList = GetTeacherIDByCLUB();

            //老師資料也必須建立
            //老師ID:老師Record
            TeacherDic = GetTeacherDic(TeacherIDList);
        }

        /// <summary>
        /// 社團學生排序
        /// </summary>
        private Dictionary<string, List<StudentRecord>> SortStudentDic(Dictionary<string, List<StudentRecord>> ClubByStudentList)
        {
            Dictionary<string, List<StudentRecord>> dic = new Dictionary<string, List<StudentRecord>>();

            foreach (string each in ClubByStudentList.Keys)
            {
                if (!dic.ContainsKey(each))
                {
                    dic.Add(each, new List<StudentRecord>());
                }
                dic[each] = SortClassIndex.K12Data_StudentRecord(ClubByStudentList[each]);

            }

            return dic;
        }

        /// <summary>
        /// 依據社團清單,取得目前的社團老師
        /// 判斷基準為社團老師1~社團老師3
        /// </summary>
        private List<string> GetTeacherIDByCLUB()
        {
            List<string> TeacherIDList = new List<string>();
            foreach (CLUBRecord each in ClubList)
            {
                if (!string.IsNullOrEmpty(each.RefTeacherID))
                {
                    if (!TeacherIDList.Contains(each.RefTeacherID))
                    {
                        TeacherIDList.Add(each.RefTeacherID);
                    }
                }
                if (!string.IsNullOrEmpty(each.RefTeacherID2))
                {
                    if (!TeacherIDList.Contains(each.RefTeacherID2))
                    {
                        TeacherIDList.Add(each.RefTeacherID2);
                    }
                }
                if (!string.IsNullOrEmpty(each.RefTeacherID3))
                {
                    if (!TeacherIDList.Contains(each.RefTeacherID3))
                    {
                        TeacherIDList.Add(each.RefTeacherID3);
                    }
                }
            }
            return TeacherIDList;
        }

        /// <summary>
        /// 取得社團老師資料
        /// </summary>
        private Dictionary<string, TeacherRecord> GetTeacherDic(List<string> TeacherIDList)
        {
            Dictionary<string, TeacherRecord> dic = new Dictionary<string, TeacherRecord>();
            foreach (TeacherRecord each in Teacher.SelectByIDs(TeacherIDList))
            {
                if (!dic.ContainsKey(each.ID))
                {
                    dic.Add(each.ID, each);
                }
            }
            return dic;
        }

        /// <summary>
        /// 取得社團字典
        /// </summary>
        private Dictionary<string, CLUBRecord> GetCLUBDic()
        {
            Dictionary<string, CLUBRecord> dic_list = new Dictionary<string, CLUBRecord>();
            foreach (CLUBRecord each in ClubList)
            {
                if (!dic_list.ContainsKey(each.UID))
                {
                    dic_list.Add(each.UID, each);
                }

            }
            return dic_list;
        }

        /// <summary>
        /// 依據社團記錄,建立一個可放入學生清單的字典
        /// </summary>
        private Dictionary<string, List<StudentRecord>> GetClubIDByStudentDic()
        {
            Dictionary<string, List<StudentRecord>> dic_list = new Dictionary<string, List<StudentRecord>>();
            foreach (CLUBRecord each in ClubList)
            {
                if (!dic_list.ContainsKey(each.UID))
                {
                    dic_list.Add(each.UID, new List<StudentRecord>());
                }
            }
            return dic_list;
        }

        /// <summary>
        /// 將學生修課記錄,取得學生ID清單
        /// </summary>
        private List<string> SCJoinSetByStudentIDList()
        {
            List<string> list = new List<string>();
            foreach (SCJoin each in SCJoinList)
            {
                if (!list.Contains(each.RefStudentID))
                {
                    list.Add(each.RefStudentID);
                }
            }
            return list;
        }

        /// <summary>
        /// 取得學生字典
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, StudentRecord> GetStudentDic()
        {
            Dictionary<string, StudentRecord> Dic = new Dictionary<string, StudentRecord>();
            foreach (StudentRecord student in StudentRecordList)
            {
                if (tool.CheckStatus(student))
                {
                    if (!Dic.ContainsKey(student.ID))
                    {
                        Dic.Add(student.ID, student);
                    }
                }
            }
            return Dic;
        }


    }
}
