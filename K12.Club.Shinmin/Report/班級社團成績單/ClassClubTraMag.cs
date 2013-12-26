using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    //資料取得器
    class ClassClubTraMag
    {
        int _SchoolYear { get; set; }
        int _Semester { get; set; }

        /// <summary>
        /// 當PrintLost為True
        /// 系統將只列印不及格清單
        /// </summary>
        bool _PrintLost { get; set; }

        AccessHelper _AccessHelper = new AccessHelper();

        /// <summary>
        /// 班級ID : 學生社團清單
        /// </summary>
        public Dictionary<string, List<ClassClubTraObj>> TraDic = new Dictionary<string, List<ClassClubTraObj>>();

        /// <summary>
        /// 班級ID : 班級Record
        /// </summary>
        public Dictionary<string, ClassRecord> ClassDic = new Dictionary<string, ClassRecord>();

        /// <summary>
        /// 老師ID : 老師Record
        /// </summary>
        public Dictionary<string, TeacherRecord> TeacherDic = new Dictionary<string, TeacherRecord>();


        public ClassClubTraMag(int SchoolYear, int Semester, bool PrintLost)
        {
            //學年度:學期
            _SchoolYear = SchoolYear;
            _Semester = Semester;
            _PrintLost = PrintLost;

            List<string> ClassIDList = GetClass();

            List<string> StudentIDList = GetStudent(ClassIDList);

            TeacherDic = GetTeacher();

            成績取得器 GetPoint = new 成績取得器();
            GetPoint.GetSCJoinByStudentIDList(StudentIDList);

            //建立資料模型
            TraDic = GetTraDic(GetPoint);

            List<ClassRecord> ClassSortList = Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource);
            ClassSortList.Sort(GetClassSort);
            Dictionary<string, List<ClassClubTraObj>> tDic = new Dictionary<string, List<ClassClubTraObj>>();
            foreach (ClassRecord cr in ClassSortList)
            {
                tDic.Add(cr.ID, TraDic[cr.ID]);
            }

            TraDic = tDic;
        }

        private int GetClassSort(ClassRecord cr1, ClassRecord cr2)
        {
            string SortValueA = cr1.GradeYear.HasValue ? cr1.GradeYear.Value.ToString().PadLeft(1, '9') : "";
            SortValueA += cr1.DisplayOrder.PadLeft(3, '9');
            SortValueA += cr1.Name.PadLeft(10, '0');

            string SortValueB = cr2.GradeYear.HasValue ? cr2.GradeYear.Value.ToString().PadLeft(1, '9') : "";
            SortValueB += cr2.DisplayOrder.PadLeft(3, '9');
            SortValueB += cr2.Name.PadLeft(10, '0');

            return SortValueA.CompareTo(SortValueB);
        }

        /// <summary>
        /// 取得資料模型
        /// </summary>
        private Dictionary<string, List<ClassClubTraObj>> GetTraDic(成績取得器 Point)
        {
            成績取得器 _GetPoint = Point;
            Dictionary<string, List<ClassClubTraObj>> dic = new Dictionary<string, List<ClassClubTraObj>>();
            foreach (ClassRecord each in ClassDic.Values)
            {
                if (!dic.ContainsKey(each.ID))
                {
                    dic.Add(each.ID, new List<ClassClubTraObj>());
                }
            }

            //建立資料內容
            foreach (StudentRecord student in _GetPoint._StudentDic.Values)
            {
                if (string.IsNullOrEmpty(student.RefClassID))
                    continue;

                if (!ClassDic.ContainsKey(student.RefClassID))
                    continue;

                ClassRecord cr = ClassDic[student.RefClassID];

                ClassClubTraObj obj = new ClassClubTraObj();
                obj.studentRecord = student; //學生
                obj.classRecord = cr; //班級

                //如果有社團記錄
                if (_GetPoint._SCJoinDic.ContainsKey(student.ID))
                {
                    List<SCJoin> scjList = _GetPoint._SCJoinDic[student.ID];
                    foreach (SCJoin each in scjList)
                    {
                        if (_GetPoint._ClubDic.ContainsKey(each.RefClubID))
                        {
                            CLUBRecord cc = _GetPoint._ClubDic[each.RefClubID];
                            if (cc.SchoolYear == _SchoolYear && cc.Semester == _Semester && cc.UID == each.RefClubID)
                            {
                                obj.club = _GetPoint._ClubDic[each.RefClubID];
                                obj.SCJoin = each; //社團記錄

                                if (_GetPoint._RSRDic.ContainsKey(each.UID))
                                {
                                    obj.RSR = _GetPoint._RSRDic[each.UID]; //社團記錄
                                }
                            }
                        }
                    }
                }

                if (obj.RSR == null)
                {
                    if (_GetPoint._RSRDic_s.ContainsKey(obj.studentRecord.ID))
                    {
                        foreach (ResultScoreRecord each in _GetPoint._RSRDic_s[obj.studentRecord.ID])
                        {
                            if (each.SchoolYear == _SchoolYear && each.Semester == _Semester)
                            {
                                obj.RSR = each;
                            }
                        }
                    }
                }


                if (dic.ContainsKey(cr.ID))
                {
                    if (_PrintLost)
                    {
                        if (obj.RSR != null)
                        {
                            if (obj.RSR.ResultScore.HasValue)
                            {
                                if (obj.RSR.ResultScore.Value < 60) //當成績低於60分才加入清單內
                                {
                                    dic[cr.ID].Add(obj);
                                }
                            }
                        }
                        else if (obj.RSR == null && obj.SCJoin != null) //有修社團,卻無結算分數
                        {
                            dic[cr.ID].Add(obj);
                        }
                        else if (obj.RSR == null && obj.SCJoin == null) //沒社團,也無結算分數
                        {
                            dic[cr.ID].Add(obj);
                        }
                    }
                    else
                    {
                        dic[cr.ID].Add(obj);
                    }
                }
            }

            return dic;
        }

        /// <summary>
        /// 取得班級老師
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, TeacherRecord> GetTeacher()
        {
            List<string> TeacherList = new List<string>();
            foreach (ClassRecord cr in ClassDic.Values)
            {
                if (string.IsNullOrEmpty(cr.RefTeacherID))
                    continue;

                if (!TeacherList.Contains(cr.RefTeacherID))
                {
                    TeacherList.Add(cr.RefTeacherID);
                }
            }

            Dictionary<string, TeacherRecord> dic = new Dictionary<string, TeacherRecord>();
            foreach (TeacherRecord cr in Teacher.SelectByIDs(TeacherList))
            {
                if (!dic.ContainsKey(cr.ID))
                {
                    dic.Add(cr.ID, cr);
                }
            }
            return dic;
        }

        /// <summary>
        /// 取得班級ID清單
        /// </summary>
        private List<string> GetClass()
        {
            List<string> ClassIDList = new List<string>();
            foreach (ClassRecord classObj in Class.SelectByIDs(K12.Presentation.NLDPanels.Class.SelectedSource))
            {
                //建立字典
                if (!ClassDic.ContainsKey(classObj.ID))
                {
                    ClassDic.Add(classObj.ID, classObj);
                }


                if (ClassIDList.Contains(classObj.ID))
                    continue;
                //建立ID清單
                ClassIDList.Add(classObj.ID);
            }
            return ClassIDList;
        }

        /// <summary>
        /// 取得學生清單
        /// </summary>
        private List<string> GetStudent(List<string> ClassIDList)
        {
            List<string> StudentIDList = new List<string>();
            foreach (StudentRecord student in Student.SelectByClassIDs(ClassIDList))
            {
                if (StudentIDList.Contains(student.ID))
                    continue;

                if (tool.CheckStatus(student))
                {
                    StudentIDList.Add(student.ID);
                }
            }
            return StudentIDList;
        }


    }
}
