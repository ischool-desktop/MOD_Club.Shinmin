using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;

namespace K12.Club.Shinmin
{
    class SCJoinRow
    {
        /// <summary>
        /// 社團參與記錄
        /// </summary>
        public SCJoin SCJ { get; set; }

        /// <summary>
        /// 社團成績結算記錄
        /// </summary>
        public ResultScoreRecord RSR { get; set; }

        /// <summary>
        /// 學生基本資料
        /// </summary>
        public StudentRecord student { get; set; }

        /// <summary>
        /// 社團資料
        /// </summary>
        public CLUBRecord club { get; set; }

        /// <summary>
        /// 老師名稱
        /// </summary>
        public TeacherRecord teacher { get; set; }

        /// <summary>
        /// 老師名稱
        /// </summary>
        public string TeacherName
        {
            get
            {
                if (teacher != null)
                {
                    if (!string.IsNullOrEmpty(teacher.Nickname))
                    {
                        return teacher.Name + "(" + teacher.Nickname + ")";
                    }
                    else
                    {
                        return teacher.Name;
                    }
                }
                else
                {
                    return "";
                }
            }
        }

        /// <summary>
        /// 社團參與記錄ID
        /// </summary>
        public string SCJoinID
        {
            get
            {
                return SCJ.UID;
            }
        }


        /// <summary>
        /// 班級名稱
        /// </summary>
        public string ClassName
        {
            get
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                    return Class.SelectByID(student.RefClassID).Name;
                else
                    return "";
            }
        }

        /// <summary>
        /// 班級排序
        /// </summary>
        public string ClassIndex
        {
            get
            {
                if (!string.IsNullOrEmpty(student.RefClassID))
                    return Class.SelectByID(student.RefClassID).DisplayOrder;
                else
                    return "";
            }
        }

        /// <summary>
        /// 學生座號
        /// </summary>
        public string SeatNo
        {
            get
            {
                return student.SeatNo.HasValue ? student.SeatNo.Value.ToString() : "";
            }
        }


        /// <summary>
        /// 學生姓名
        /// </summary>
        public string StudentName
        {
            get
            {
                return student.Name;
            }
        }

        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber
        {
            get
            {
                return student.StudentNumber;
            }
        }

        /// <summary>
        /// 社團名稱
        /// </summary>
        public string ClubName
        {
            get
            {
                return club.ClubName;
            }
        }

        /// <summary>
        /// 是否進行資料修改
        /// </summary>
        public bool HasChange { get; set; }

        public decimal? pa_Score
        {
            get
            {
                return SCJ.PAScore;
            }
            set
            {
                HasChange = true;
                SCJ.PAScore = value;
            }
        }
        public decimal? ar_Score
        {
            get
            {
                return SCJ.ARScore;
            }
            set
            {
                HasChange = true;
                SCJ.ARScore = value;
            }
        }
        public decimal? aas_Score
        {
            get
            {
                return SCJ.AASScore;
            }
            set
            {
                HasChange = true;
                SCJ.AASScore = value;
            }
        }
        public decimal? far_Score
        {
            get
            {
                return SCJ.FARScore;
            }
            set
            {
                HasChange = true;
                SCJ.FARScore = value;
            }
        }


        public decimal? all_Score
        {
            get
            {
                if (RSR != null)
                    return RSR.ResultScore;
                else
                    return null;
            }
        }
    }
}
