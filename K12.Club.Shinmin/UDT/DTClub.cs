using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    //設定學校的開始選社時間
    [TableName("K12.DTClub.Shinmin")]
    class DTClub : ActiveRecord
    {
        /// <summary>
        /// 年級
        /// </summary>
        [Field(Field = "gradeyear", Indexed = false)]
        public int GradeYear { get; set; }

        /// <summary>
        /// 開始時間
        /// </summary>
        [Field(Field = "startDate", Indexed = false)]
        public DateTime? Start { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        [Field(Field = "endDate", Indexed = false)]
        public DateTime? End { get; set; }
    }
}
