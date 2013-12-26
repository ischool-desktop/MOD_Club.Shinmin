using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using K12.Data;

namespace K12.Club.Shinmin
{
    class ClubTraMag
    {
        /// <summary>
        /// 社團ID : 社團記錄
        /// </summary>
        public Dictionary<string, CLUBRecord> CLUBDic = new Dictionary<string, CLUBRecord>();

        /// <summary>
        /// 課程ID : 社團老師
        /// </summary>
        public Dictionary<string, TeacherRecord> TeacherDic = new Dictionary<string, TeacherRecord>();

        /// <summary>
        /// 學生ID : 學生記錄
        /// </summary>
        public Dictionary<string, StudentRecord> StudentDic = new Dictionary<string, StudentRecord>();

        /// <summary>
        /// 社團參與記錄ID : 學期成績
        /// </summary>
        Dictionary<string, ResultScoreRecord> RSRDic = new Dictionary<string, ResultScoreRecord>();

        /// <summary>
        /// 社團 : ClubTraObj(StudentRecord/SCJoin/ResultScoreRecord
        /// </summary>
        public Dictionary<string, List<ClubTraObj>> TraDic = new Dictionary<string, List<ClubTraObj>>();


        public ClubTraMag()
        {
            //取得社團
            #region 資料取得

            List<string> TeacherList = new List<string>();
            List<CLUBRecord> CLUBList = tool._A.Select<CLUBRecord>(ClubAdmin.Instance.SelectedSource);
            foreach (CLUBRecord cr in CLUBList)
            {
                if (!TraDic.ContainsKey(cr.UID))
                {
                    TraDic.Add(cr.UID, new List<ClubTraObj>());
                }

                if (!CLUBDic.ContainsKey(cr.UID))
                {
                    CLUBDic.Add(cr.UID, cr);
                }

                #region 收集老師資訊

                if (!string.IsNullOrEmpty(cr.RefTeacherID))
                {
                    TeacherList.Add(cr.RefTeacherID);
                }

                if (!string.IsNullOrEmpty(cr.RefTeacherID2))
                {
                    TeacherList.Add(cr.RefTeacherID2);
                }

                if (!string.IsNullOrEmpty(cr.RefTeacherID3))
                {
                    TeacherList.Add(cr.RefTeacherID3);
                }

                #endregion
            }

            //社團老師Record
            TeacherDic = GetTeacher(TeacherList);

            //取得社團修課記錄
            List<SCJoin> SCJoinList = tool._A.Select<SCJoin>(string.Format("ref_club_id in ('{0}')", string.Join("','", ClubAdmin.Instance.SelectedSource)));
            List<string> SCJoinIDList = new List<string>();
            List<string> StudentList = new List<string>();
            foreach (SCJoin each in SCJoinList)
            {
                if (!SCJoinIDList.Contains(each.UID))
                {
                    SCJoinIDList.Add(each.UID);
                }

                if (!StudentList.Contains(each.RefStudentID))
                {
                    StudentList.Add(each.RefStudentID);
                }

            }
            //取得學生
            foreach (StudentRecord student in Student.SelectByIDs(StudentList))
            {
                if (tool.CheckStatus(student))
                {
                    if (!StudentDic.ContainsKey(student.ID))
                    {
                        StudentDic.Add(student.ID, student);
                    }
                }
            }

            //取得社團結算記錄
            List<ResultScoreRecord> RSRList = tool._A.Select<ResultScoreRecord>(string.Format("ref_scjoin_id in ('{0}')", string.Join("','", SCJoinIDList)));
            foreach (ResultScoreRecord each in RSRList)
            {
                if (!RSRDic.ContainsKey(each.RefSCJoinID))
                {
                    RSRDic.Add(each.RefSCJoinID, each);
                }
            }

            #endregion

            #region 建立主資料清單

            foreach (SCJoin each in SCJoinList)
            {
                if (TraDic.ContainsKey(each.RefClubID))
                {
                    ClubTraObj ctObj = new ClubTraObj();
                    ctObj.SCJ = each;
                    if (StudentDic.ContainsKey(each.RefStudentID))
                    {
                        ctObj.student = StudentDic[each.RefStudentID];
                        if (RSRDic.ContainsKey(each.UID))
                        {
                            ctObj.RSR = RSRDic[each.UID];
                        }
                        TraDic[each.RefClubID].Add(ctObj);
                    }
                }

            }

            #endregion
        }

        /// <summary>
        /// 傳入老師ID,取得老師清單
        /// </summary>
        private Dictionary<string, TeacherRecord> GetTeacher(List<string> list)
        {
            Dictionary<string, TeacherRecord> dic = new Dictionary<string, TeacherRecord>();
            foreach (TeacherRecord teacher in Teacher.SelectByIDs(list))
            {
                if (!dic.ContainsKey(teacher.ID))
                {
                    dic.Add(teacher.ID, teacher);
                }
            }
            return dic;
        }

    }
}
