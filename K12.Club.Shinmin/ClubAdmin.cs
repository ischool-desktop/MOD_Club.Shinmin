using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation;
using System.Text.RegularExpressions;
using K12.Data;
using System.ComponentModel;
using FISCA.UDT;
using System.Xml;
using FISCA.Data;
using System.Data;
using FISCA.DSAUtil;

namespace K12.Club.Shinmin
{
    //主要控制社團資料來源用
    public partial class ClubAdmin : NLDPanel
    {
        /// <summary>
        /// 取得社團資料背景模式
        /// </summary>
        private BackgroundWorker BGW = new BackgroundWorker();

        /// <summary>
        /// 預設學年度學期
        /// </summary>
        private string FiltedSemester = School.DefaultSchoolYear.ToString().PadLeft(3, ' ') + "學年度 第" + School.DefaultSemester + "學期";

        /// <summary>
        /// 學年度學期社團 / 社團ID
        /// </summary>
        private Dictionary<string, List<string>> SemesterClub = new Dictionary<string, List<string>>();

        /// <summary>
        /// 社團老師清單
        /// </summary>
        private Dictionary<string, TeacherRecord> TeacherDic = new Dictionary<string, TeacherRecord>();

        /// <summary>
        /// 學生清單
        /// </summary>
        private Dictionary<string, ClubAdminStudent> StudentDic = new Dictionary<string, ClubAdminStudent>();

        private Dictionary<string, ClubAdminClub> ClubCountSCJoin = new Dictionary<string, ClubAdminClub>();

        private AccessHelper _AccessHelper = new AccessHelper();
        private QueryHelper _QueryHelper = new QueryHelper();

        #region Field

        /// <summary>
        /// 社團代碼
        /// </summary>
        public ListPaneField Field0;
        /// <summary>
        /// 名稱
        /// </summary>
        public ListPaneField Field1;
        /// <summary>
        /// 代碼
        /// </summary>
        public ListPaneField Field12;
        /// <summary>
        /// 老師1
        /// </summary>
        public ListPaneField Field2;
        /// <summary>
        /// 老師2
        /// </summary>
        public ListPaneField Field2_2;
        /// <summary>
        /// 老師3
        /// </summary>
        public ListPaneField Field2_3;
        /// <summary>
        /// 場地 
        /// </summary>
        public ListPaneField Field3;
        /// <summary>
        ///社長
        /// </summary>
        public ListPaneField Field4;
        /// <summary>
        /// 副社長
        /// </summary>
        public ListPaneField Field5;
        /// <summary>
        ///性 別條件
        /// </summary>
        public ListPaneField Field6;
        /// <summary>
        ///科別限制
        /// </summary>
        public ListPaneField Field7;
        /// <summary>
        /// 一年級人數限制
        /// </summary>
        public ListPaneField Field8;
        /// <summary>
        /// 二年級人數限制
        /// </summary>
        public ListPaneField Field9;
        /// <summary>
        /// 三年級人數限制
        /// </summary>
        public ListPaneField Field10;
        /// <summary>
        /// 目前人數 / 社團人數上限
        /// </summary>
        public ListPaneField Field11;

        #endregion

        /// <summary>
        /// 搜尋內容
        /// </summary>
        private MenuButton SearchClubName;
        private MenuButton SearchClubTeacher;
        private MenuButton SearchStudentName;
        private MenuButton SearchAddress;

        /// <summary>
        /// 社團取得資料背景,是否忙錄中
        /// </summary>
        private bool isbusy = false;


        /// <summary>
        /// 搜尋
        /// </summary>
        private SearchEventArgs SearEvArgs = null;

        /// <summary>
        /// 所有系統內的社團資料
        /// </summary>
        private Dictionary<string, CLUBRecord> AllClubDic = new Dictionary<string, CLUBRecord>();

        public ClubAdmin()
        {
            Group = "新民社團";

            #region 事件

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);

            //社團自我更新事件
            ClubEvents.ClubChanged += new EventHandler(ClubEvents_ClubChanged);

            #endregion

            #region 代碼
            Field0 = new ListPaneField("代碼");
            Field0.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    e.Value = AllClubDic[e.Key].ClubNumber;
                }
            };
            this.AddListPaneField(Field0);
            #endregion
            #region 名稱
            Field1 = new ListPaneField("名稱");
            Field1.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    e.Value = AllClubDic[e.Key].ClubName;
                }
            };
            this.AddListPaneField(Field1);
            #endregion
            #region 類型
            Field12 = new ListPaneField("類型");
            Field12.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    e.Value = AllClubDic[e.Key].ClubCategory;
                }
            };
            this.AddListPaneField(Field12);
            #endregion
            #region 老師
            Field2 = new ListPaneField("老師1");
            Field2.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    if (TeacherDic.ContainsKey(AllClubDic[e.Key].RefTeacherID))
                    {
                        if (string.IsNullOrEmpty(TeacherDic[AllClubDic[e.Key].RefTeacherID].Nickname))
                        {
                            e.Value = TeacherDic[AllClubDic[e.Key].RefTeacherID].Name;
                        }
                        else
                        {
                            string terName = TeacherDic[AllClubDic[e.Key].RefTeacherID].Name;
                            string Nicname = TeacherDic[AllClubDic[e.Key].RefTeacherID].Nickname;
                            e.Value = terName + "(" + Nicname + ")";
                        }
                    }
                }
            };
            this.AddListPaneField(Field2);

            Field2_2 = new ListPaneField("老師2");
            Field2_2.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    if (TeacherDic.ContainsKey(AllClubDic[e.Key].RefTeacherID2))
                    {
                        if (string.IsNullOrEmpty(TeacherDic[AllClubDic[e.Key].RefTeacherID2].Nickname))
                        {
                            e.Value = TeacherDic[AllClubDic[e.Key].RefTeacherID2].Name;
                        }
                        else
                        {
                            string terName = TeacherDic[AllClubDic[e.Key].RefTeacherID2].Name;
                            string Nicname = TeacherDic[AllClubDic[e.Key].RefTeacherID2].Nickname;
                            e.Value = terName + "(" + Nicname + ")";
                        }
                    }
                }
            };
            this.AddListPaneField(Field2_2);

            Field2_3 = new ListPaneField("老師3");
            Field2_3.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    if (TeacherDic.ContainsKey(AllClubDic[e.Key].RefTeacherID3))
                    {
                        if (string.IsNullOrEmpty(TeacherDic[AllClubDic[e.Key].RefTeacherID3].Nickname))
                        {
                            e.Value = TeacherDic[AllClubDic[e.Key].RefTeacherID3].Name;
                        }
                        else
                        {
                            string terName = TeacherDic[AllClubDic[e.Key].RefTeacherID3].Name;
                            string Nicname = TeacherDic[AllClubDic[e.Key].RefTeacherID3].Nickname;
                            e.Value = terName + "(" + Nicname + ")";
                        }
                    }
                }
            };
            this.AddListPaneField(Field2_3);
            #endregion
            #region 場地
            Field3 = new ListPaneField("場地");
            Field3.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    e.Value = AllClubDic[e.Key].Location;
                }
            };
            this.AddListPaneField(Field3);
            #endregion
            #region 社長
            Field4 = new ListPaneField("社長");
            Field4.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    //需取得社員資料
                    //對照出社長姓名
                    if (StudentDic.ContainsKey(AllClubDic[e.Key].President))
                    {
                        e.Value = StudentDic[AllClubDic[e.Key].President].StudentName;
                    }
                }
            };
            this.AddListPaneField(Field4);
            #endregion
            #region 副社長
            Field5 = new ListPaneField("副社長");
            Field5.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    //需取得副社員資料
                    //對照出副社長姓名
                    if (StudentDic.ContainsKey(AllClubDic[e.Key].VicePresident))
                    {
                        e.Value = StudentDic[AllClubDic[e.Key].VicePresident].StudentName;
                    }
                }
            };
            this.AddListPaneField(Field5);
            #endregion
            #region 性別條件
            Field6 = new ListPaneField("性別條件");
            Field6.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    e.Value = AllClubDic[e.Key].GenderRestrict;
                }
            };
            this.AddListPaneField(Field6);
            #endregion
            #region 科別限制
            Field7 = new ListPaneField("科別限制");
            Field7.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    if (AllClubDic[e.Key].DeptRestrict != "")
                    {
                        List<string> list = new List<string>();
                        XmlElement xml = DSXmlHelper.LoadXml(AllClubDic[e.Key].DeptRestrict);
                        foreach (XmlElement mini in xml.SelectNodes("Dept"))
                        {
                            list.Add(mini.InnerText);
                        }
                        string deptString = string.Join(",", list);
                        e.Value = deptString;
                    }
                }
            };
            this.AddListPaneField(Field7);
            #endregion
            #region 一年級人數限制
            Field8 = new ListPaneField("一年級人數/上限");
            Field8.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    //1年級的人數??
                    int count1 = 0;

                    if (ClubCountSCJoin.ContainsKey(e.Key))
                    {
                        count1 = ClubCountSCJoin[e.Key].一年級人數;
                    }

                    //人數上限
                    string count2 = AllClubDic[e.Key].Grade1Limit.HasValue ? AllClubDic[e.Key].Grade1Limit.Value.ToString() : "";
                    if (string.IsNullOrEmpty(count2))
                    {
                        e.Value = count1 + "/無限制";
                    }
                    else
                    {
                        e.Value = count1 + "/" + count2;
                    }
                }
            };
            this.AddListPaneField(Field8);
            #endregion
            #region 二年級人數限制
            Field9 = new ListPaneField("二年級人數/上限");
            Field9.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    //2年級的人數??
                    int count1 = 0;

                    if (ClubCountSCJoin.ContainsKey(e.Key))
                    {
                        count1 = ClubCountSCJoin[e.Key].二年級人數;
                    }
                    
                    //人數上限
                    string count2 = AllClubDic[e.Key].Grade2Limit.HasValue ? AllClubDic[e.Key].Grade2Limit.Value.ToString() : "";
                    if (string.IsNullOrEmpty(count2))
                    {
                        e.Value = count1 + "/無限制";
                    }
                    else
                    {
                        e.Value = count1 + "/" + count2;
                    }
                }
            };
            this.AddListPaneField(Field9);
            #endregion
            #region 三年級人數限制
            Field10 = new ListPaneField("三年級人數/上限");
            Field10.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    //3年級的人數??
                    int count1 = 0;

                    if (ClubCountSCJoin.ContainsKey(e.Key))
                    {
                        count1 = ClubCountSCJoin[e.Key].三年級人數;
                    }

                    //人數上限
                    string count2 = AllClubDic[e.Key].Grade3Limit.HasValue ? AllClubDic[e.Key].Grade3Limit.Value.ToString() : "";
                    if (string.IsNullOrEmpty(count2))
                    {
                        e.Value = count1 + "/無限制";
                    }
                    else
                    {
                        e.Value = count1 + "/" + count2;
                    }
                }
            };
            this.AddListPaneField(Field10);
            #endregion
            #region 目前人數 / 社團人數上限
            Field11 = new ListPaneField("目前總人數/上限");
            Field11.GetVariable += delegate(object sender, GetVariableEventArgs e)
            {
                if (AllClubDic.ContainsKey(e.Key))
                {
                    //count出本社團的SCJoin數量
                    //就是人數
                    int count1 = 0;

                    if (ClubCountSCJoin.ContainsKey(e.Key))
                    {
                        count1 = ClubCountSCJoin[e.Key].社團人數;
                    }

                    string count2 = AllClubDic[e.Key].Limit.HasValue ? AllClubDic[e.Key].Limit.Value.ToString() : "";
                    if (string.IsNullOrEmpty(count2))
                    {
                        e.Value = count1 + "/無限制";
                    }
                    else
                    {
                        e.Value = count1 + "/" + count2;
                    }
                }

            };
            this.AddListPaneField(Field11);
            #endregion


            FilterMenu.SupposeHasChildern = true;
            FilterMenu.PopupOpen += new EventHandler<PopupOpenEventArgs>(FilterMenu_PopupOpen);
            //目前預設學年期設為標題
            FilterMenu.Text = FiltedSemester;

            //設定預設學年期
            //SetClubList(FiltedSemester);

            #region Search

            Campus.Configuration.ConfigData cd = Campus.Configuration.Config.User["AssociationSearchOptionPreference"];

            SearchClubName = SetSearchButton("名稱", "SearchClubName", cd);
            SearchClubTeacher = SetSearchButton("老師", "SearchClubTeacher", cd);
            SearchStudentName = SetSearchButton("參與學生", "SearchStudentName", cd);
            SearchAddress = SetSearchButton("場地", "SearchAddress", cd);
            #endregion

            this.Search += new EventHandler<SearchEventArgs>(ClubAdmin_Search);

            //取得社團基本資料
            BGW.RunWorkerAsync();
        }

        void ClubEvents_ClubChanged(object sender, EventArgs e)
        {
            if (BGW.IsBusy)
            {
                isbusy = true;
            }
            else
            {
                BGW.RunWorkerAsync();
            }
        }

        /// <summary>
        /// 設定可搜尋欄位
        /// </summary>
        private MenuButton SetSearchButton(string MenuName, string BoolMenuName, Campus.Configuration.ConfigData cd)
        {
            MenuButton SearchName = SearchConditionMenu[MenuName];
            SearchName.AutoCheckOnClick = true;
            SearchName.AutoCollapseOnClick = false;
            SearchName.Checked = CheckStringIsBool(cd[BoolMenuName]);
            SearchName.Click += delegate
            {
                cd[BoolMenuName] = SearchName.Checked.ToString(); ;
                BackgroundWorker async = new BackgroundWorker();
                async.DoWork += delegate(object sender, DoWorkEventArgs e) { (e.Argument as Campus.Configuration.ConfigData).Save(); };
                async.RunWorkerAsync(cd);
            };

            return SearchName;
        }

        void FilterMenu_PopupOpen(object sender, PopupOpenEventArgs e)
        {
            if (!BGW.IsBusy)
            {

                List<string> list = new List<string>();

                foreach (string item in SemesterClub.Keys)
                {
                    list.Add(item);
                }
                list.Sort();

                foreach (string item in list)
                {
                    MenuButton mb = e.VirtualButtons[item];
                    mb.AutoCheckOnClick = true;
                    mb.AutoCollapseOnClick = true;
                    mb.Checked = (item == FiltedSemester);
                    mb.Tag = item;
                    mb.CheckedChanged += delegate(object sender1, EventArgs e1)
                    {
                        MenuButton mb1 = sender1 as MenuButton;
                        SetClubList(mb1.Text);
                        FiltedSemester = FilterMenu.Text = mb1.Text;
                        mb1.Checked = true;
                    };
                }
            }
            else //如果忙碌中則提醒使用者
            {
                e.Cancel = true;
                e.VirtualButtons.Text = "資料下載中...";
                FISCA.Presentation.Controls.MsgBox.Show("資料下載中...\n請稍後再試");
            }
        }

        /// <summary>
        /// 由此背景模式
        /// 取得系統所需的各種社團資料
        /// 社團基本資料
        /// 導師資料
        /// 場地資料...
        /// </summary>
        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            ReflashClubList();
        }

        /// <summary>
        /// 重新取得所有社團資料
        /// </summary>
        private void ReflashClubList()
        {
            AllClubDic.Clear();
            SemesterClub.Clear();
            TeacherDic.Clear();
            StudentDic.Clear();
            ClubCountSCJoin.Clear();

            //社團基本資料
            List<CLUBRecord> ClubList = _AccessHelper.Select<CLUBRecord>();
            foreach (CLUBRecord each in ClubList)
            {
                if (!AllClubDic.ContainsKey(each.UID))
                {
                    AllClubDic.Add(each.UID, each);
                }

                string SchoolYearSemester = each.SchoolYear.ToString().PadLeft(3, ' ') + "學年度 第" + each.Semester.ToString() + "學期";

                if (!SemesterClub.ContainsKey(SchoolYearSemester))
                {
                    //學年度學期名稱/當學期的社團課程ID清單
                    SemesterClub.Add(SchoolYearSemester, new List<string>());
                }

                SemesterClub[SchoolYearSemester].Add(each.UID);
            }


            //老師資料       
            foreach (TeacherRecord each in Teacher.SelectAll())
            {
                if (!TeacherDic.ContainsKey(each.ID))
                {
                    TeacherDic.Add(each.ID, each);
                }
            }

            StringBuilder sb_c = new StringBuilder();
            sb_c.Append("select student.id,student.name,class.class_name,class.grade_year from student");
            sb_c.Append(" LEFT join class on class.id=student.ref_class_id");
            sb_c.Append(" where student.status=1 or student.status=2");

            //取得系統內狀態為(一般 / 延修)之學生
            DataTable dTable = _QueryHelper.Select(sb_c.ToString());
            foreach (DataRow row in dTable.Rows)
            {
                ClubAdminStudent cs = new ClubAdminStudent(row);

                if (!StudentDic.ContainsKey(cs.StudentID))
                {
                    StudentDic.Add(cs.StudentID, cs);
                }
            }

            //學生社團參與記錄(統計人數用)

            //避免沒有SCJoin Table
            List<SCJoin> list = _AccessHelper.Select<SCJoin>("uid = 00000");

            //DataTable dt_scjoin = _QueryHelper.Select("select * from " + Tn._SCJoinUDT.ToLower() + " where ref_student_id='103802'");

            DataTable dt_scjoin = _QueryHelper.Select("select ref_club_id,ref_student_id from " + Tn._SCJoinUDT.ToLower());

            foreach (DataRow row in dt_scjoin.Rows)
            {
                string ref_club_id = "" + row[0];
                string ref_student_id = "" + row[1];

                if (!ClubCountSCJoin.ContainsKey(ref_club_id))
                {
                    ClubCountSCJoin.Add(ref_club_id, new ClubAdminClub());
                }

                if (StudentDic.ContainsKey(ref_student_id))
                {
                    ClubCountSCJoin[ref_club_id].社團人數++;

                    if (StudentDic[ref_student_id].GradeYear == "1")
                    {
                        ClubCountSCJoin[ref_club_id].一年級人數++;
                    }
                    else if (StudentDic[ref_student_id].GradeYear == "2")
                    {
                        ClubCountSCJoin[ref_club_id].二年級人數++;
                    }
                    else if (StudentDic[ref_student_id].GradeYear == "3")
                    {
                        ClubCountSCJoin[ref_club_id].三年級人數++;
                    }
                }
            }




            isbusy = false;

        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (isbusy)
            {
                BGW.RunWorkerAsync();
            }
            else
            {
                //傳入使用者預設所選的學年度 / 學期
                SetClubList(FilterMenu.Text);
            }

            //依社團名稱排一下
            Field1.Column.DataGridView.Sort(Field0.Column, ListSortDirection.Ascending);

        }

        /// <summary>
        /// 設定學期資料
        /// </summary>
        /// <param name="NowSemester"></param>
        private void SetClubList(string NowSemester)
        {
            if (SemesterClub.ContainsKey(NowSemester))
            {
                SetFilteredSource(SemesterClub[NowSemester]);
            }
            else //傳入非為目前有資料之學年期,則加入空清單
            {
                SetFilteredSource(new List<string>());
            }
        }

        /// <summary>
        /// 檢查傳入String是否為Bool資料
        /// 是:傳出True
        /// 否:傳出False(預設)
        /// </summary>
        private bool CheckStringIsBool(string p)
        {
            bool k = false;
            bool.TryParse(p, out k);
            return k;
        }

        #region 搜尋功能(最後再處理)

        void ClubAdmin_Search(object sender, SearchEventArgs e)
        {
            SearEvArgs = e;
            Campus.Windows.BlockMessage.Display("資料搜尋中,請稍候....", new Campus.Windows.ProcessInvoker(ProcessSearch));
        }

        /// <summary>
        /// 搜尋資料
        /// </summary>
        /// <param name="args"></param>
        private void ProcessSearch(Campus.Windows.MessageArgs args)
        {
            #region 取得學生

            Dictionary<string, StudentRecord> SearchStudentDic = new Dictionary<string, StudentRecord>();
            StringBuilder sb_c = new StringBuilder();
            sb_c.Append("select student.id from student");
            sb_c.Append(" LEFT join class on class.id=student.ref_class_id");

            //取得系統內狀態為(一般 / 延修)之學生
            DataTable dTable = _QueryHelper.Select(sb_c.ToString());
            List<string> StudentList = new List<string>();
            foreach (DataRow row in dTable.Rows)
            {
                StudentList.Add("" + row[0]);
            }

            foreach (StudentRecord stud in Student.SelectByIDs(StudentList))
            {
                if (!SearchStudentDic.ContainsKey(stud.ID))
                {
                    SearchStudentDic.Add(stud.ID, stud);
                }
            }



            #endregion



            List<string> results = new List<string>();
            Regex rx = new Regex(SearEvArgs.Condition, RegexOptions.IgnoreCase);

            #region 社團名稱
            if (SearchClubName.Checked)
            {
                foreach (string each in AllClubDic.Keys)
                {
                    if (rx.Match(AllClubDic[each].ClubName).Success)
                    {
                        if (!results.Contains(each))
                            results.Add(each);
                    }
                }
            }
            #endregion

            #region 社團老師
            if (SearchClubTeacher.Checked)
            {
                foreach (string each in AllClubDic.Keys)
                {
                    if (TeacherDic.ContainsKey(AllClubDic[each].RefTeacherID))
                    {
                        string terName = TeacherDic[AllClubDic[each].RefTeacherID].Name;
                        string Nicname = TeacherDic[AllClubDic[each].RefTeacherID].Nickname;

                        if (rx.Match(terName).Success || rx.Match(Nicname).Success)
                        {
                            if (!results.Contains(each))
                                results.Add(each);
                        }
                    }
                }
            }
            #endregion

            #region 社團學生

            if (SearchStudentName.Checked)
            {
                //取得所有社團參與記錄
                List<SCJoin> list = _AccessHelper.Select<SCJoin>();

                foreach (SCJoin _join in list)
                {
                    if (SearchStudentDic.ContainsKey(_join.RefStudentID))
                    {
                        if (rx.Match(SearchStudentDic[_join.RefStudentID].Name).Success)
                        {
                            //是否為存在系統內的社團資料???
                            if (AllClubDic.ContainsKey(_join.RefClubID))
                            {
                                if (!results.Contains(_join.RefClubID))
                                    results.Add(_join.RefClubID);
                            }
                        }
                    }
                }
            }

            #endregion

            #region 場地
            if (SearchAddress.Checked)
            {
                foreach (string each in AllClubDic.Keys)
                {
                    if (rx.Match(AllClubDic[each].Location).Success)
                    {
                        if (!results.Contains(each))
                            results.Add(each);
                    }
                }
            }
            SearEvArgs.Result.AddRange(results);
            #endregion
        }

        #endregion

        private static ClubAdmin _ClubAdmin;

        public static ClubAdmin Instance
        {
            get
            {
                if (_ClubAdmin == null)
                {
                    _ClubAdmin = new ClubAdmin();
                }
                return _ClubAdmin;
            }
        }
    }
}
