using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace K12.Club.Shinmin
{
    class TeacherObj
    {
        /// <summary>
        /// 導師ID
        /// </summary>
        public string TeacherID { get; set; }
        /// <summary>
        /// 導師姓名
        /// </summary>
        public string TeacherName { get; set; }

        /// <summary>
        /// 導師暱稱
        /// </summary>
        public string TeacherNickName { get; set; }

        public string TeacherFullName
        {
            get
            {
                string teachername = string.Empty;

                if (!string.IsNullOrEmpty(TeacherNickName))
                {
                    teachername = TeacherName + "(" + TeacherNickName + ")";
                }
                else
                {
                    teachername = TeacherName;
                }

                return teachername;
            }
        }
    }
}
