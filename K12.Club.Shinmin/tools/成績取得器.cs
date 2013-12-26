using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    class 成績取得器
    {
        #region 資料盒

        /// <summary>
        /// 學生ID : 學生Record
        /// </summary>
        public Dictionary<string, StudentRecord> _StudentDic { get; set; }

        /// <summary>
        /// 社團ID : 社團Record
        /// </summary>
        public Dictionary<string, CLUBRecord> _ClubDic { get; set; }

        /// <summary>
        /// 學生ID : List SCJoin Record
        /// </summary>
        public Dictionary<string, List<SCJoin>> _SCJoinDic { get; set; }

        /// <summary>
        /// SCJoin UID : List ResultScore Record
        /// </summary>
        public Dictionary<string, ResultScoreRecord> _RSRDic { get; set; }

        /// <summary>
        /// 學生ID : List Record
        /// </summary>
        public Dictionary<string, List<ResultScoreRecord>> _RSRDic_s { get; set; }

        /// <summary>
        /// 成績比例原則
        /// </summary>
        public WeightProportion _wp { get; set; }

        #endregion

        /// <summary>
        /// 本功能將會把傳入的學生ID清單
        /// 組織成一個可以方便使用的物件模型
        /// </summary>
        public 成績取得器()
        {
            _StudentDic = new Dictionary<string, StudentRecord>();
            _ClubDic = new Dictionary<string, CLUBRecord>();
            _SCJoinDic = new Dictionary<string, List<SCJoin>>();
            _RSRDic = new Dictionary<string, ResultScoreRecord>();
            _RSRDic_s = new Dictionary<string, List<ResultScoreRecord>>();
        }

        /// <summary>
        /// 由學生清單取得成績
        /// (將會取得不分學年度學期之社團參與記錄)
        /// </summary>
        public void GetSCJoinByStudentIDList(List<string> StudentList)
        {
            if (StudentList.Count == 0)
                return;

            #region 取得學生
            StudentList = StudentList.Distinct().ToList();
            foreach (StudentRecord student in Student.SelectByIDs(StudentList))
            {
                if (tool.CheckStatus(student))
                {
                    if (!_StudentDic.ContainsKey(student.ID))
                    {
                        _StudentDic.Add(student.ID, student);
                    }
                }
            }

            #endregion

            #region 取得社團參與記錄

            string qu = string.Join("','", _StudentDic.Keys);
            List<SCJoin> SCJoinList = tool._A.Select<SCJoin>(string.Format("ref_student_id in ('{0}')", qu));
            foreach (SCJoin scj in SCJoinList)
            {
                if (!_SCJoinDic.ContainsKey(scj.RefStudentID))
                {
                    _SCJoinDic.Add(scj.RefStudentID, new List<SCJoin>());
                }
                _SCJoinDic[scj.RefStudentID].Add(scj);
            }

            #endregion

            #region 取得社團結算成績
            List<string> scjList = new List<string>();
            foreach (SCJoin scj in SCJoinList)
            {
                if (!scjList.Contains(scj.UID))
                {
                    scjList.Add(scj.UID);
                }
            }
            string ju = string.Join("','", scjList);
            List<ResultScoreRecord> RSRList = tool._A.Select<ResultScoreRecord>(string.Format("ref_scjoin_id in ('{0}')", ju));
            foreach (ResultScoreRecord rsr in RSRList)
            {
                if (!_RSRDic.ContainsKey(rsr.RefSCJoinID))
                {
                    _RSRDic.Add(rsr.RefSCJoinID, rsr);
                }
            }

            string su = string.Join("','", _StudentDic.Keys.ToList());
            List<ResultScoreRecord> RSRList_s = tool._A.Select<ResultScoreRecord>(string.Format("ref_student_id in ('{0}')", su));
            foreach (ResultScoreRecord rsr in RSRList_s)
            {

                if (!_RSRDic_s.ContainsKey(rsr.RefStudentID))
                {
                    _RSRDic_s.Add(rsr.RefStudentID, new List<ResultScoreRecord>());
                }
                _RSRDic_s[rsr.RefStudentID].Add(rsr);
            }
            
            #endregion

            #region 取得課程

            //從社團參與記錄內,取得相關社團資料
            List<string> ClubList = new List<string>();
            foreach (SCJoin scj in SCJoinList)
            {
                if (!ClubList.Contains(scj.RefClubID))
                {
                    ClubList.Add(scj.RefClubID);
                }
            }

            _ClubDic = GetAllData.GetClub(ClubList);

            #endregion
        }

        /// <summary>
        /// 由課程清單取得成績
        /// </summary>
        public void GetSCJoinByClubIDList(List<string> ClubList)
        {
            if (ClubList.Count == 0)
                return;

            #region 取得課程

            _ClubDic = GetAllData.GetClub(ClubList);


            #endregion

            #region 取得社團參與記錄

            string qu = string.Join("','", ClubList);
            List<SCJoin> SCJoinList = tool._A.Select<SCJoin>(string.Format("ref_club_id in ('{0}')", qu));

            foreach (SCJoin scj in SCJoinList)
            {
                if (!_SCJoinDic.ContainsKey(scj.RefStudentID))
                {
                    _SCJoinDic.Add(scj.RefStudentID, new List<SCJoin>());
                }
                _SCJoinDic[scj.RefStudentID].Add(scj);
            }

            #endregion

            #region 取得社團結算成績

            List<string> scjList = new List<string>();
            foreach (SCJoin scj in SCJoinList)
            {
                if (!scjList.Contains(scj.UID))
                {
                    scjList.Add(scj.UID);
                }
            }
            string ju = string.Join("','", scjList);
            List<ResultScoreRecord> RSRList = tool._A.Select<ResultScoreRecord>(string.Format("ref_scjoin_id in ('{0}')", ju));
            foreach (ResultScoreRecord rsr in RSRList)
            {
                if (!_RSRDic.ContainsKey(rsr.RefSCJoinID))
                {
                    _RSRDic.Add(rsr.RefSCJoinID, rsr);
                }
            }
            #endregion


            #region 取得學生

            List<StudentRecord> StudentList = Student.SelectByIDs(_SCJoinDic.Keys);
            foreach (StudentRecord student in StudentList)
            {
                if (tool.CheckStatus(student))
                {
                    if (!_StudentDic.ContainsKey(student.ID))
                    {
                        _StudentDic.Add(student.ID, student);
                    }
                }
            }

            #endregion
        }

        /// <summary>
        /// 取得比例原則
        /// </summary>
        public void SetWeightProportion()
        {
            List<WeightProportion> wpList = tool._A.Select<WeightProportion>();
            if (wpList.Count >= 1)
            {
                _wp = wpList[0];
            }
        }

        /// <summary>
        /// 計算比例成績
        /// </summary>
        public decimal GetDecimalValue(SCJoin scj)
        {
            decimal results = 0;

            if (_wp != null)
            {
                decimal? 平時活動比例 = scj.PAScore;
                decimal? 出缺率比例 = scj.ARScore;
                decimal? 活動力及服務比例 = scj.AASScore;
                decimal? 成品成果考驗比例 = scj.FARScore;

                if (平時活動比例.HasValue)
                    平時活動比例 = _wp.PA_Weight * 平時活動比例 / 100;

                if (出缺率比例.HasValue)
                    出缺率比例 = _wp.AR_Weight * 出缺率比例 / 100;

                if (活動力及服務比例.HasValue)
                    活動力及服務比例 = _wp.AAS_Weight * 活動力及服務比例 / 100;

                if (成品成果考驗比例.HasValue)
                    成品成果考驗比例 = _wp.FAR_Weight * 成品成果考驗比例 / 100;

                if (平時活動比例.HasValue)
                    results += 平時活動比例.Value;

                if (出缺率比例.HasValue)
                    results += 出缺率比例.Value;

                if (活動力及服務比例.HasValue)
                    results += 活動力及服務比例.Value;

                if (成品成果考驗比例.HasValue)
                    results += 成品成果考驗比例.Value;
            }

            return results = Math.Round(results, MidpointRounding.AwayFromZero);
        }

    }
}
