using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using System.Xml;
using FISCA.DSAUtil;

namespace K12.Club.Shinmin
{
    /// <summary>
    /// 本物件模型
    /// 用以存放每個分類大項的資料
    /// </summary>
    class FactsObj_big
    {
        public FactsObj_big(CLUBRecord cr)
        {
            Club = cr;
            StudentList_1 = new List<StudentRecord>();
            StudentList_2 = new List<StudentRecord>();
            StudentList_3 = new List<StudentRecord>();
        }

        /// <summary>
        /// 社團ID
        /// </summary>
        public string CLUBID { get; set; }

        /// <summary>
        /// 社團物件
        /// </summary>
        public CLUBRecord Club { get; set; }

        /// <summary>
        /// 老師物件
        /// </summary>
        public TeacherCrk Teacher { get; set; }

        /// <summary>
        /// 1年級學生
        /// </summary>
        public List<StudentRecord> StudentList_1 { get; set; }

        /// <summary>
        /// 2年級學生
        /// </summary>
        public List<StudentRecord> StudentList_2 { get; set; }

        /// <summary>
        /// 3年級學生
        /// </summary>
        public List<StudentRecord> StudentList_3 { get; set; }

        public string 社團名稱
        {
            get
            {
                return Club.ClubName;
            }
        }

        public string 社團類型
        {
            get
            {
                return Club.ClubCategory;
            }
        }

        public string 社團代碼
        {
            get
            {
                return Club.ClubNumber;
            }
        }

        public string 老師姓名
        {
            get
            {
                if (Teacher != null)
                {
                    return Teacher.FullName;
                }
                else
                {
                    return "";
                }
            }
        }

        public string 活動場地
        {
            get
            {
                return Club.Location;
            }
        }

        public string 科別限制
        {
            get
            {
                if (!string.IsNullOrEmpty(Club.DeptRestrict))
                {
                    List<string> list = new List<string>();
                    XmlElement xml = DSXmlHelper.LoadXml(Club.DeptRestrict);
                    foreach (XmlElement a in xml.SelectNodes("Dept"))
                    {
                        list.Add(a.InnerText);
                    }
                    return string.Join(",", list);
                }
                else
                {

                    return "";
                }
            }
        }

        public string 性別限制
        {
            get
            {
                if (!string.IsNullOrEmpty(Club.GenderRestrict))
                {
                    return Club.GenderRestrict;
                }
                else
                {
                    return "";
                }
            }
        }

        public string 一年級限制
        {
            get
            {
                if (Club.Grade1Limit.HasValue)
                {
                    if (Club.Grade1Limit.Value != 0)
                    {
                        return Club.Grade1Limit.Value.ToString();
                    }
                    else
                    {
                        return "不開放";
                    }
                }
                else
                {
                    return "無限制";
                }
            }
        }

        public string 二年級限制
        {
            get
            {
                if (Club.Grade2Limit.HasValue)
                {
                    if (Club.Grade2Limit.Value != 0)
                    {
                        return Club.Grade2Limit.Value.ToString();
                    }
                    else
                    {
                        return "不開放";
                    }
                }
                else
                {
                    return "無限制";
                }
            }
        }

        public string 三年級限制
        {
            get
            {
                if (Club.Grade3Limit.HasValue)
                {
                    if (Club.Grade3Limit.Value != 0)
                    {
                        return Club.Grade3Limit.Value.ToString();
                    }
                    else
                    {
                        return "不開放";
                    }
                }
                else
                {
                    return "無限制";
                }
            }
        }


        public string 人數上限
        {
            get
            {
                if (Club.Limit.HasValue)
                {
                    if (Club.Limit.Value != 0)
                    {
                        return Club.Limit.Value.ToString();
                    }
                    else
                    {
                        return "不開放";
                    }
                }
                else
                {
                    return "無限制";
                }
            }
        }

        public string 一年級
        {
            get
            {
                return StudentList_1.Count + "/" + 一年級限制;
            }
        }

        public string 二年級
        {
            get
            {
                return StudentList_2.Count + "/" + 二年級限制;
            }
        }

        public string 三年級
        {
            get
            {
                return StudentList_3.Count + "/" + 三年級限制;
            }
        }

        public string 總人數
        {
            get
            {
                int count = StudentList_1.Count + StudentList_2.Count + StudentList_3.Count;
                return count + "/" + 人數上限;
            }
        }
    }
}
