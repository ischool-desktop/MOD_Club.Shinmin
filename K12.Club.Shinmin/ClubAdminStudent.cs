using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace K12.Club.Shinmin
{
    class ClubAdminStudent
    {
        public string StudentID { get; set; }
        public string StudentName { get; set; }
        public string ClassName { get; set; }
        public string GradeYear { get; set; }

        public ClubAdminStudent(DataRow row)
        {
            StudentID = "" + row[0];
            StudentName = "" + row[1];
            ClassName = "" + row[2];
            GradeYear = "" + row[3];
        }
    }
}
