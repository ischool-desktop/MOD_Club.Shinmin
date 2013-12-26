using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Club.Shinmin
{
    class CopyClubRecord
    {

        //社團
        public CLUBRecord _cr { get; set; }
        //社團學生清單
        public List<SCJoin> _scj { get; set; }

        //新社團
        public CLUBRecord _new_cr { get; set; }
        //新社團記錄
        public List<SCJoin> _new_scj { get; set; }

        public CopyClubRecord(CLUBRecord cr)
        {
            _cr = cr;
            //建立一個社團 / 社團學生的物件
            //並且傳入新社團的相關資料

            //本物件的功能
            //就是傳入新社團的ID
            //而本物件依據ID來建立學生記錄


        }

        /// <summary>
        /// 設定本社團的參與學生
        /// </summary>
        /// <param name="scj"></param>
        public void SetSCJ(List<SCJoin> scj)
        {
            _scj = scj;
        }

        /// <summary>
        /// 建立學生的社團記錄Record
        /// </summary>
        public List<SCJoin> GetNewSCJoinList()
        {
            List<SCJoin> list = new List<SCJoin>();
            if (_new_cr != null && _scj.Count != 0)
            {
                foreach (SCJoin each in _scj)
                {
                    SCJoin scj = new SCJoin();
                    scj.RefStudentID = each.RefStudentID;
                    scj.RefClubID = _new_cr.UID;
                    list.Add(scj);
                }
            }
            return list;
        }
    }
}
