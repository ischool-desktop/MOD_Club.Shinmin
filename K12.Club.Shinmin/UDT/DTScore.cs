using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    //設定社團成績輸入時間
    [TableName("K12.DTScore.Shinmin")]
    class DTScore : ActiveRecord
    {
        /// <summary>
        /// 開始時間
        /// </summary>
        [Field(Field = "start", Indexed = false)]
        public DateTime? Start { get; set; }

        /// <summary>
        /// 結束時間
        /// </summary>
        [Field(Field = "end", Indexed = false)]
        public DateTime? End { get; set; }
    }
}
