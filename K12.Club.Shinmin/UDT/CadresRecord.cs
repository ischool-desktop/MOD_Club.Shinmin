using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    [TableName("K12.CadresRecord.Shinmin")]
    class CadresRecord : ActiveRecord
    {
        /// <summary>
        /// 幹部名稱
        /// </summary>
        [Field(Field = "cadre_name", Indexed = true)]
        public string CadreName { get; set; }

        /// <summary>
        /// 社團參考
        /// </summary>
        [Field(Field = "ref_club_id", Indexed = true)]
        public string RefClubID { get; set; }

        /// <summary>
        /// 學生參考
        /// </summary>
        [Field(Field = "ref_student_id", Indexed = true)]
        public string RefStudentID { get; set; }
    }
}
