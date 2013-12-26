 using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    //社團成績計算比例原則
    [TableName("K12.WeightProportion.Shinmin")]
    class WeightProportion : ActiveRecord
    {
        /// <summary>
        /// 平時活動比例
        /// </summary>
        [Field(Field = "pa_weight", Indexed = false)]
        public int PA_Weight { get; set; }

        /// <summary>
        /// 出缺率比例
        /// </summary>
        [Field(Field = "ar_weight", Indexed = false)]
        public int AR_Weight { get; set; }

        /// <summary>
        /// 活動力及服務比例
        /// </summary>
        [Field(Field = "aas_weight", Indexed = false)]
        public int AAS_Weight { get; set; }

        /// <summary>
        /// 成品成果考驗比例
        /// </summary>
        [Field(Field = "far_weight", Indexed = false)]
        public int FAR_Weight { get; set; }
    }
}
