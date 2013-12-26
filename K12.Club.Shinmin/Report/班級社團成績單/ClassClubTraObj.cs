using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace K12.Club.Shinmin
{
    //每名學生的基本資料
    class ClassClubTraObj
    {
        /// <summary>
        /// 學生
        /// </summary>
        public StudentRecord studentRecord { get; set; }

        /// <summary>
        /// 社團參與記錄
        /// </summary>
        public SCJoin SCJoin { get; set; }

        /// <summary>
        /// 社團基本資料
        /// </summary>
        public CLUBRecord club { get; set; }

        /// <summary>
        /// 學期成績
        /// </summary>
        public ResultScoreRecord RSR { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public ClassRecord classRecord { get; set; }


        public ClassClubTraObj()
        {


        }
    }
}
