using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace K12.Club.Shinmin
{
    class ClubCadresObj
    {
        /// <summary>
        /// 幹部名稱
        /// </summary>
        public string CadreName { get; set; }

        /// <summary>
        /// 社團物件
        /// </summary>
        public CLUBRecord _Club { get; set; }

        /// <summary>
        /// 學生
        /// </summary>
        public StudentRecord _Student { get; set; }

        public string ref_student_id { get; set; }
    }
}
