using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace K12.Club.Shinmin
{
    /// <summary>
    /// 社團老師
    /// </summary>
    public class TeacherCrk
    {
        public TeacherCrk(DataRow row)
        {
            ID = "" + row[0];
            Name = "" + row[1];
            NickName = "" + row[2];
        }

        public string ID { get; set; }
        public string Name { get; set; }
        public string NickName { get; set; }
        public string FullName
        {
            get
            {
                if (!string.IsNullOrEmpty(NickName))
                {
                    return Name + "(" + NickName + ")";
                }
                else
                {
                    return Name;
                }

            }
        }
    }

}
