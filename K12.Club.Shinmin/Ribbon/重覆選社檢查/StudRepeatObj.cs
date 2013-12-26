using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace K12.Club.Shinmin
{
    class StudRepeatObj
    {
        /// <summary>
        /// 學生ID
        /// </summary>
        public string _StudentID { get; set; }

        /// <summary>
        /// 學生姓名
        /// </summary>
        public string _StudentName { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public string _StudentClass { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string _StudentSeatNo { get; set; }

        /// <summary>
        /// 學號
        /// </summary>
        public string _StudentNumber { get; set; }

        /// <summary>
        /// 學生記錄
        /// </summary>
        public StudentRecord _Student { get; set; }

        /// <summary>
        /// 學生社團參與記錄
        /// </summary>
        public List<SCJoin> _SCJoinList { get; set; }

        /// <summary>
        /// 清除之記錄
        /// </summary>
        public List<SCJoin> _RemoveList { get; set; }

        public Dictionary<int, SCJoin> _SCJionDIc { get; set; }

        public bool Change { get; set; }

        public StudRepeatObj(StudentRecord Student)
        {
            _Student = Student;
            _StudentID = Student.ID;
            _StudentClass = string.IsNullOrEmpty(Student.RefClassID) ? "" : Student.Class.Name;
            _StudentSeatNo = Student.SeatNo.HasValue ? Student.SeatNo.Value.ToString() : "";
            _StudentNumber = Student.StudentNumber;
            _StudentName = Student.Name;
            _SCJoinList = new List<SCJoin>();
            _SCJionDIc = new Dictionary<int, SCJoin>();
            _RemoveList = new List<SCJoin>();
            Change = false;
        }


    }
}
