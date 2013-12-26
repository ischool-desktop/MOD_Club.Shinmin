using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using System.Data;

namespace K12.Club.Shinmin
{
    static class GetAllData
    {
        /// <summary>
        /// 取得傳入的社團ID清單
        /// (含依據社團序號/社團名稱排序)
        /// </summary>
        static public Dictionary<string, CLUBRecord> GetClub(List<string> ClubIDList)
        {
            Dictionary<string, CLUBRecord> dic = new Dictionary<string, CLUBRecord>();
            List<CLUBRecord> ClubList = tool._A.Select<CLUBRecord>(ClubIDList);
            ClubList.Sort(SortClub);
            foreach (CLUBRecord club in ClubList)
            {
                if (!dic.ContainsKey(club.UID))
                {
                    dic.Add(club.UID, club);
                }
            }
            return dic;
        }

        /// <summary>
        /// 排序社團依據:代碼/名稱排序
        /// </summary>
        static private int SortClub(CLUBRecord cr1, CLUBRecord cr2)
        {
            string Comp1 = cr1.ClubNumber.PadLeft(5, '0');
            Comp1 += cr1.ClubName.PadLeft(20, '0');

            string Comp2 = cr2.ClubNumber.PadLeft(5, '0');
            Comp2 += cr2.ClubName.PadLeft(20, '0');

            return Comp1.CompareTo(Comp2);
        }

        /// <summary>
        /// 取得本學年度學期的社團清單
        /// </summary>
        static public Dictionary<string, CLUBRecord> GetSchoolYearClub()
        {
            Dictionary<string, CLUBRecord> dic = new Dictionary<string, CLUBRecord>();
            List<CLUBRecord> ClubList = tool._A.Select<CLUBRecord>(string.Format("school_year={0} and semester={1}", School.DefaultSchoolYear, School.DefaultSemester));
            foreach (CLUBRecord club in ClubList)
            {
                if (!dic.ContainsKey(club.UID))
                {
                    dic.Add(club.UID, club);
                }
            }
            return dic;
        }

        /// <summary>
        /// 取得傳入的學生ID清單
        /// </summary>
        static public Dictionary<string, StudentRecord> GetStudent(List<string> StudentIDList)
        {
            Dictionary<string, StudentRecord> dic = new Dictionary<string, StudentRecord>();

            List<StudentRecord> StudentList = Student.SelectByIDs(StudentIDList);
            foreach (StudentRecord sr in StudentList)
            {
                if (!dic.ContainsKey(sr.ID))
                {
                    dic.Add(sr.ID, sr);
                }
            }
            return dic;
        }
    }
}
