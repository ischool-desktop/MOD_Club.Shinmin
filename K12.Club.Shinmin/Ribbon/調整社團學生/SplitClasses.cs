using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using DevComponents.DotNetBar;
using FISCA.UDT;
using FISCA.Data;
using System.Drawing.Drawing2D;
using K12.Data;

namespace K12.Club.Shinmin
{
    public partial class SplitClasses : BaseForm
    {
        BackgroundWorker Save_BGW = new BackgroundWorker();
        BackgroundWorker BGW_FormLoad = new BackgroundWorker();
        //UDT物件
        private AccessHelper _AccessHelper = new AccessHelper();
        //private QueryHelper _QueryHelper = new QueryHelper();

        int ButtonCount = 0;


        private SortedList<string, string> _ClubNameList = new SortedList<string, string>();
        private SortedList<string, Color> _ClubColorList = new SortedList<string, Color>();
        Dictionary<string, List<StudentJoinRow>> _StudentAttenCourses = new Dictionary<string, List<StudentJoinRow>>();

        Color[] colors = new Color[] { Color.Red, Color.Yellow, Color.Blue, Color.PowderBlue, Color.Orange, Color.Green, Color.Purple };

        private Dictionary<DataGridViewRow, int> _RowIndex = new Dictionary<DataGridViewRow, int>();


        List<StudentRecord> studentRecordList = new List<StudentRecord>();

        private Dictionary<string, DataRow_clsRecord> new_ClassDic = new Dictionary<string, DataRow_clsRecord>();
        Dictionary<string, SCJoin> new_SCJoinDic = new Dictionary<string, SCJoin>();
        Dictionary<string, CLUBRecord> new_ClubDic = new Dictionary<string, CLUBRecord>();
        List<StudentJoinRow> Rowlist = new List<StudentJoinRow>();



        public SplitClasses()
        {
            InitializeComponent();
        }

        private void SplitClasses_Load(object sender, EventArgs e)
        {
            Save_BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Save_BGW_RunWorkerCompleted);
            Save_BGW.DoWork += new DoWorkEventHandler(Save_BGW_DoWork);

            BGW_FormLoad.DoWork += new DoWorkEventHandler(BGW_FormLoad_DoWork);
            BGW_FormLoad.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_FormLoad_RunWorkerCompleted);

            btnSave.Enabled = false;
            BGW_FormLoad.RunWorkerAsync();
        }

        void BGW_FormLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            //社團
            new_ClubDic = GetSplitOBJ.GetClubRecord();

            //學生參與記錄
            new_SCJoinDic = GetSplitOBJ.GetSCJoin();

            Dictionary<string, List<SCJoin>> new_StudentBySCJoin = new Dictionary<string, List<SCJoin>>();
            foreach (SCJoin each in new_SCJoinDic.Values)
            {
                if (!new_StudentBySCJoin.ContainsKey(each.RefStudentID))
                    new_StudentBySCJoin.Add(each.RefStudentID, new List<SCJoin>());

                new_StudentBySCJoin[each.RefStudentID].Add(each);
            }

            //學生參與記錄
            //new_StudentIDBySCJoin = GetSplitOBJ.GetSCJoin();

            //由SCJoin 取得不重覆的學生ID
            List<string> StudentIDList = GetSplitOBJ.GetStudentIDList(new_SCJoinDic.Values.ToList());

            //取得班級相關資訊清單
            new_ClassDic = GetSplitOBJ.GetClassRecord(StudentIDList);

            //取得學生StudentRecord
            studentRecordList = Student.SelectByIDs(StudentIDList);

            //依據學生建立資料行
            Rowlist.Clear();
            foreach (StudentRecord studentRec in studentRecordList)
            {
                if (tool.CheckStatus(studentRec))
                {
                    StudentJoinRow newJoinRow = new StudentJoinRow();
                    newJoinRow.Ref_Student_Id = studentRec.ID;
                    newJoinRow.Name = studentRec.Name;
                    newJoinRow.SeatNo = studentRec.SeatNo.HasValue ? studentRec.SeatNo.Value.ToString() : "";
                    newJoinRow.StudentNumber = studentRec.StudentNumber;
                    newJoinRow.Gender = studentRec.Gender;

                    //填入學生班級/年級資訊
                    if (studentRec.RefClassID != null)
                    {
                        if (new_ClassDic.ContainsKey(studentRec.RefClassID))
                        {
                            newJoinRow.GradeYear = new_ClassDic[studentRec.RefClassID].gradeyear;
                            newJoinRow.ClassName = new_ClassDic[studentRec.RefClassID].classname;
                        }
                    }

                    //取得社員參與記錄
                    if (new_StudentBySCJoin.ContainsKey(studentRec.ID))
                    {
                        newJoinRow.SCJoinList = new_StudentBySCJoin[studentRec.ID];
                    }

                    //取得目前社團資料
                    if (newJoinRow.SCJoinList.Count == 1)
                    {
                        newJoinRow.CurrentClubRecord = new_ClubDic[newJoinRow.SCJoinList[0].RefClubID];
                    }

                    Rowlist.Add(newJoinRow);
                }
            }
            Rowlist.Sort(SortCLUBRecord);

        }

        private int SortCLUBRecord(StudentJoinRow x,StudentJoinRow y)
        {
            //課程名稱
            string ax = x.CurrentClubName.PadLeft(20, '0');
            ax += x.GradeYear.PadLeft(2, '0');
            ax += x.ClassName.PadLeft(8, '0');
            ax += x.SeatNo.PadLeft(3, '0');

            string ay = y.CurrentClubName.PadLeft(20, '0');
            ay += y.GradeYear.PadLeft(2, '0');
            ay += y.ClassName.PadLeft(8, '0');
            ay += y.SeatNo.PadLeft(3, '0');

            return ax.CompareTo(ay);
        }

        void BGW_FormLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;

            if (e.Cancelled)
            {
                MsgBox.Show("背景作業已取消...");
            }
            else
            {
                if (e.Error == null)
                {
                    _ClubNameList.Clear();
                    _ClubColorList.Clear();
                    ButtonCount = 0;
                    //加入Button
                    this.flowLayoutPanel1.Controls.Clear();
                    foreach (CLUBRecord each in new_ClubDic.Values)
                    {
                        ButtonX item1 = SetBotton(each);
                        this.flowLayoutPanel1.Controls.Add(item1); //加入到Panel裡面
                    }

                    dataGridViewX1.AutoGenerateColumns = false;
                    dataGridViewX1.DataSource = new BindingList<StudentJoinRow>(Rowlist);

                    foreach (DataGridViewRow row in dataGridViewX1.Rows)
                    {
                        StudentJoinRow scjRow = (StudentJoinRow)row.DataBoundItem;
                        if (scjRow.SCJoinList.Count > 1)
                        {
                            row.ReadOnly = true;
                            foreach (DataGridViewCell cell in row.Cells)
                            {
                                cell.Style.BackColor = Color.Gray;
                            }
                        }
                        else if (scjRow.SCJoinList.Count == 1)
                        {
                            ((DataGridViewColorBallTextCell)row.Cells[7]).Color = _ClubColorList[scjRow.CurrentClubRecord.UID];
                            ((DataGridViewColorBallTextCell)row.Cells[7]).Value = _ClubNameList[scjRow.CurrentClubRecord.UID];

                            ((DataGridViewColorBallTextCell)row.Cells[6]).Color = _ClubColorList[scjRow.CurrentClubRecord.UID];
                            ((DataGridViewColorBallTextCell)row.Cells[6]).Value = _ClubNameList[scjRow.CurrentClubRecord.UID];


                        }
                    }

                    #region 舊資料
                    //DataGridViewRow row = new DataGridViewRow();
                    //row.CreateCells(dataGridViewX1);
                    //row.Cells[0].Tag = studentRec; //Cell0 - 儲存StudentRecord
                    //string cr = studentRec.RefClassID;
                    //if (ClassDic.ContainsKey(cr))
                    //{
                    //    //年級
                    //    row.Cells[0].Value = ClassDic[cr].gradeyear;
                    //    //班級
                    //    row.Cells[1].Value = ClassDic[cr].classname;

                    //}
                    ////座號
                    //row.Cells[2].Value = studentRec.SeatNo.HasValue ? studentRec.SeatNo.Value.ToString() : "";
                    ////學號
                    //row.Cells[3].Value = studentRec.StudentNumber;
                    ////性別
                    //row.Cells[4].Value = studentRec.Gender;
                    ////姓名
                    //row.Cells[5].Value = studentRec.Name;
                    ////"Count 大於 1"表示此學生重覆參與社團
                    //row.Cells[6].Value = _StudentAttenCourses[studentRec.ID].Count == 1 ? _ClubNameList[_StudentAttenCourses[studentRec.ID][0].ClubID] : "重覆參與社團";
                    //row.Cells[7].Value = _StudentAttenCourses[studentRec.ID].Count == 1 ? _ClubNameList[_StudentAttenCourses[studentRec.ID][0].ClubID] : "重覆參與社團";
                    //int newRowIndex = dataGridViewX1.Rows.Add(row);

                    ////"Count==1"表示此學生重覆參與社團
                    //if (_StudentAttenCourses[studentRec.ID].Count == 1)
                    //{
                    //    dataGridViewX1.Rows[newRowIndex].Tag = _StudentAttenCourses[studentRec.ID][0];
                    //    ((DataGridViewColorBallTextCell)dataGridViewX1.Rows[newRowIndex].Cells[6]).Color = _CourseColor[_StudentAttenCourses[studentRec.ID][0].CourseID];
                    //    ((DataGridViewColorBallTextCell)dataGridViewX1.Rows[newRowIndex].Cells[7]).Color = _CourseColor[_StudentAttenCourses[studentRec.ID][0].CourseID];
                    //}
                    //else
                    //{
                    //    //1.顏色為Gray
                    //    //2.整Row欄位皆為"唯讀"
                    //    dataGridViewX1.Rows[newRowIndex].ReadOnly = true;
                    //    foreach (DataGridViewCell cell in dataGridViewX1.Rows[newRowIndex].Cells)
                    //    {
                    //        cell.Style.BackColor = Color.Gray;
                    //    }
                    //} 
                    #endregion

                    dataGridViewX1_SelectionChanged(null, null);
                    dataGridViewX1_Sorted(null, null);
                    CountStudents();
                }
                else
                {
                    MsgBox.Show("取得資料發生錯誤...\n" + e.Error.Message);
                }
            }
        }

        /// <summary>
        /// 設定課程Button
        /// </summary>
        private ButtonX SetBotton(CLUBRecord cr)
        {
            //顏色每7個社團,顏色重置
            if (ButtonCount > 6)
                ButtonCount = 0;

            ButtonX item1 = new ButtonX();
            item1.FocusCuesEnabled = false;
            item1.Style = eDotNetBarStyle.Office2007;
            item1.ColorTable = eButtonColor.Flat;// eButtonColor.Office2007WithBackground;
            //item1.ButtonStyle = DevComponents.DotNetBar.eButtonStyle.ImageAndText;
            //item1.ImagePaddingHorizontal = 8;
            item1.AutoSize = true;
            item1.Shape = new DevComponents.DotNetBar.RoundRectangleShapeDescriptor(15);
            item1.TextAlignment = eButtonTextAlignment.Left;
            item1.Size = new Size(110, 23);
            item1.Text = cr.ClubName;
            item1.Image = GetColorBallImage(colors[ButtonCount]);
            _ClubNameList.Add("" + cr.UID, cr.ClubName);
            _ClubColorList.Add("" + cr.UID, colors[ButtonCount++]);

            item1.Tag = cr;
            item1.Click += new EventHandler(item1_Click);
            return item1;

        }

        /// <summary>
        /// 當按鈕被按下
        /// </summary>
        void item1_Click(object sender, EventArgs e)
        {
            ButtonX item1 = (ButtonX)sender;
            CLUBRecord cr = (CLUBRecord)item1.Tag;

            //int count = 0;
            //foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
            //{
            //    StudentJoinRow scjRow = (StudentJoinRow)row.DataBoundItem;
            //    if (scjRow.SCJoinList.Count > 1)
            //    {
            //        count++;
            //    }
            //}
            //if (count != 0)
            //    MsgBox.Show("有重覆參與社團之學生「" + count + "」名\n無法設定為「" + cr.ClubName + "」");

            foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
            {
                StudentJoinRow scjRow = (StudentJoinRow)row.DataBoundItem;
                if (scjRow.SCJoinList.Count > 1)
                {
                    continue;
                }

                scjRow.ChangeCLUB(cr);
                ((DataGridViewColorBallTextCell)row.Cells[7]).Color = _ClubColorList[((CLUBRecord)item1.Tag).UID];
                ((DataGridViewColorBallTextCell)row.Cells[7]).Value = _ClubNameList[((CLUBRecord)item1.Tag).UID];
            }

            CountStudents();
        }

        /// <summary>
        /// 選擇變更時
        /// </summary>
        private void dataGridViewX1_SelectionChanged(object sender, EventArgs e)
        {

            foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
            {
                StudentJoinRow scjRow = (StudentJoinRow)row.DataBoundItem;
                if (scjRow.SCJoinList.Count > 1)
                {
                    row.Selected = false;
                    return;
                }
            }
            labelX1.Text = "選取「" + dataGridViewX1.SelectedRows.Count + "」人";
        }

        /// <summary>
        /// 進行排序後
        /// 用途未明
        /// </summary>
        private void dataGridViewX1_Sorted(object sender, EventArgs e)
        {
            _RowIndex.Clear();
            foreach (DataGridViewRow row in this.dataGridViewX1.Rows)
            {
                _RowIndex.Add(row, row.Index);
            }
        }

        /// <summary>
        /// 計算學生度量
        /// </summary>
        private void CountStudents()
        {
            //建立:以社團ID為單位,學生人數統計
            Dictionary<string, List<string>> courseStudents = new Dictionary<string, List<string>>();
            foreach (SCJoin each in new_SCJoinDic.Values)
            {
                if (!courseStudents.ContainsKey(each.RefClubID))
                    courseStudents.Add(each.RefClubID, new List<string>());
                courseStudents[each.RefClubID].Add(each.RefStudentID);
            }

            //取得畫面上的Button?
            foreach (Control var in flowLayoutPanel1.Controls)
            {
                //Button的TAG是社團CLUBRecord
                CLUBRecord club = (CLUBRecord)var.Tag;
                //不包含在上述清單內,人數就是0人
                if (!courseStudents.ContainsKey(club.UID))
                {
                    var.Text = _ClubNameList[club.UID] + "(0人)";
                }
                else
                {
                    //取得學生人數
                    int totle = courseStudents[club.UID].Count;
                    //男女統計
                    int b = 0, g = 0;

                    //取得每個社團的學生清單
                    //(此區域寫法可能會比較慢...)
                    foreach (StudentRecord studentRec in Student.SelectByIDs(courseStudents[club.UID].ToArray()))
                    {
                        if (tool.CheckStatus(studentRec))
                        {
                            if (studentRec.Gender == "男") b++;
                            if (studentRec.Gender == "女") g++;
                        }
                    }
                    StringBuilder sb_bg = new StringBuilder();
                    //社團名稱
                    sb_bg.Append(_ClubNameList[club.UID]);
                    //男
                    sb_bg.Append("(" + (b > 0 ? " " + b + "男" : ""));
                    //女
                    sb_bg.Append((g > 0 ? " " + g + "女" : ""));
                    //totle-男-女>0　表示有n個未知性別
                    sb_bg.Append((totle - b - g > 0 ? " " + (totle - b - g) + "未知性別" : ""));
                    sb_bg.Append(" 共" + totle + "人" + " )");

                    var.Text = sb_bg.ToString();
                }
            }


        }

        /// <summary>
        /// 儲存調整後狀態
        /// </summary>
        private void btnSave_Click(object sender, EventArgs e)
        {
            if (!Save_BGW.IsBusy)
            {
                btnSave.Enabled = false;
                Save_BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("忙碌中請稍後再試!");
            }
        }

        void Save_BGW_DoWork(object sender, DoWorkEventArgs e)
        {

            List<SCJoin> UpdataList = new List<SCJoin>();

            foreach (StudentJoinRow row in Rowlist)
            {
                if (row.HasChanged)
                {
                    SCJoin scj = row.GetChange();
                    if (scj != null)
                    {
                        UpdataList.Add(scj);
                    }
                }
            }

            //_AccessHelper.InsertValues(InasertList);
            _AccessHelper.UpdateValues(UpdataList);


        }

        void Save_BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;


            if (e.Cancelled)
            {
                MsgBox.Show("儲存作業已取消...");
            }
            else
            {
                if (e.Error == null)
                {
                    MsgBox.Show("社團學生調整成功!!");

                    ClubEvents.RaiseAssnChanged();

                    BGW_FormLoad.RunWorkerAsync();
                }
                else
                {
                    MsgBox.Show("儲存資料發生錯誤...\n" + e.Error.Message);
                    return;
                }
            }
        }

        /// <summary>
        /// 取得每個Cell開頭的小圖
        /// </summary>
        public Image GetColorBallImage(Color color)
        {
            Bitmap bmp = new Bitmap(16, 16);
            Graphics graphics = Graphics.FromImage(bmp);
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            int w = 14,
                    x = 1,
                    y = 1;
            Color[] myColors = { color, Color.White, color, color };
            float[] myPositions = { 0.0f, 0.05f, 0.6f, 1.0f };
            ColorBlend myBlend = new ColorBlend();
            myBlend.Colors = myColors;
            myBlend.Positions = myPositions;
            using (LinearGradientBrush brush = new LinearGradientBrush(new Point(x, y), new Point(w, w), Color.White, color))
            {
                brush.InterpolationColors = myBlend;
                brush.GammaCorrection = true;
                graphics.FillRectangle(brush, x, y, w, w);
            }
            graphics.DrawRectangle(new Pen(Color.Black), x, y, w, w);
            return bmp;
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
