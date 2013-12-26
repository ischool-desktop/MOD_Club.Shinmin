using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.UDT;
using K12.Data;

namespace K12.Club.Shinmin
{
    public partial class ClubResultsInput : BaseForm
    {

        //1.是否修改資料檢查
        //2.資料儲存時,需要是數字型態
        //3.資料輸入時,必須小於100
        //4.社團修課學生狀態:
        成績取得器 GetPoint { get; set; }

        BackgroundWorker Save_BGW = new BackgroundWorker();
        BackgroundWorker BGW_FormLoad = new BackgroundWorker();

        private AccessHelper _AccessHelper = new AccessHelper();

        List<SCJoinRow> RowList = new List<SCJoinRow>();

        Campus.Windows.ChangeListener _ChangeListener = new Campus.Windows.ChangeListener();

        Dictionary<string, Log_Result> _logDic = new Dictionary<string, Log_Result>();

        Dictionary<string, ClassRecord> ClassDic { get; set; }

        bool IsChangeNow = false;

        public ClubResultsInput()
        {
            InitializeComponent();
        }

        private void ClubResultsInput_Load(object sender, EventArgs e)
        {
            //取得學生社團參與記錄
            BGW_FormLoad.DoWork += new DoWorkEventHandler(BGW_FormLoad_DoWork);
            BGW_FormLoad.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_FormLoad_RunWorkerCompleted);

            //儲存資料(更新社團參與記錄)
            Save_BGW.DoWork += new DoWorkEventHandler(Save_BGW_DoWork);
            Save_BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(Save_BGW_RunWorkerCompleted);

            dataGridViewX1.DataError += new DataGridViewDataErrorEventHandler(dataGridViewX1_DataError);

            _ChangeListener.StatusChanged += new EventHandler<Campus.Windows.ChangeEventArgs>(_ChangeListener_StatusChanged);
            _ChangeListener.Add(new Campus.Windows.DataGridViewSource(dataGridViewX1));

            if (ClubAdmin.Instance.SelectedSource.Count == 0)
                return;

            btnReport.Enabled = false;
            btnSave.Enabled = false;
            this.Text = "成績輸入(資料讀取中..)";

            //取得成績計算比例原則
            GetPoint = new 成績取得器();
            GetPoint.SetWeightProportion();

            if (GetPoint._wp != null)
            {
                dataGridViewX1.Columns[Col_pa_Score.Index].HeaderText += "(" + GetPoint._wp.PA_Weight + "%)";
                dataGridViewX1.Columns[Col_ar_Score.Index].HeaderText += "(" + GetPoint._wp.AR_Weight + "%)";
                dataGridViewX1.Columns[Col_aas_Score.Index].HeaderText += "(" + GetPoint._wp.AAS_Weight + "%)";
                dataGridViewX1.Columns[Col_far_Score.Index].HeaderText += "(" + GetPoint._wp.FAR_Weight + "%)";
            }

            BGW_FormLoad.RunWorkerAsync();

        }

        void _ChangeListener_StatusChanged(object sender, Campus.Windows.ChangeEventArgs e)
        {
            IsChangeNow = (e.Status == Campus.Windows.ValueStatus.Dirty);
        }

        void dataGridViewX1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {

            MsgBox.Show("輸入資料錯誤!!");
            e.Cancel = false;
        }

        void BGW_FormLoad_DoWork(object sender, DoWorkEventArgs e)
        {
            StringBuilder sb_3 = new StringBuilder();
            GetPoint = new 成績取得器();
            GetPoint.SetWeightProportion();
            GetPoint.GetSCJoinByClubIDList(ClubAdmin.Instance.SelectedSource);


            #region 社團老師資訊

            List<string> teacherIDList = new List<string>();
            foreach (CLUBRecord club in GetPoint._ClubDic.Values)
            {
                if (!string.IsNullOrEmpty(club.RefTeacherID))
                {
                    if (!teacherIDList.Contains(club.RefTeacherID))
                    {
                        teacherIDList.Add(club.RefTeacherID);
                    }
                }
            }

            List<TeacherRecord> TeacherList = Teacher.SelectByIDs(teacherIDList);
            Dictionary<string, TeacherRecord> ClubTeacherDic = new Dictionary<string, TeacherRecord>();
            foreach (TeacherRecord each in TeacherList)
            {
                if (!ClubTeacherDic.ContainsKey(each.ID))
                {
                    ClubTeacherDic.Add(each.ID, each);
                }
            }

            #endregion

            #region 取得班級資料

            //從學生Record內取得班級ID,再取得班級Record
            ClassDic = GetClassDic();

            #endregion

            RowList.Clear();
            //取得社團參與記錄
            foreach (List<SCJoin> each in GetPoint._SCJoinDic.Values)
            {
                #region 只有一筆資料
                if (each.Count == 1)
                {
                    SCJoin sch = each[0];

                    SCJoinRow scjRow = new SCJoinRow();
                    scjRow.SCJ = sch;
                    //學生
                    if (GetPoint._StudentDic.ContainsKey(sch.RefStudentID))
                    {
                        scjRow.student = GetPoint._StudentDic[sch.RefStudentID];

                        //社團
                        if (GetPoint._ClubDic.ContainsKey(sch.RefClubID))
                        {
                            scjRow.club = GetPoint._ClubDic[sch.RefClubID];

                            if (ClubTeacherDic.ContainsKey(GetPoint._ClubDic[sch.RefClubID].RefTeacherID))
                            {
                                scjRow.teacher = ClubTeacherDic[GetPoint._ClubDic[sch.RefClubID].RefTeacherID];
                            }

                        }

                        if (GetPoint._RSRDic.ContainsKey(sch.UID))
                        {
                            scjRow.RSR = GetPoint._RSRDic[sch.UID];
                        }

                        RowList.Add(scjRow);
                    }
                #endregion
                }
                else if (each.Count >= 1)
                {
                    #region 有兩筆資料
                    //錯誤訊息
                    StudentRecord sr = Student.SelectByID(each[0].RefStudentID);
                    sb_3.AppendLine("學生[" + sr.Name + "]有2筆以上社團記錄");

                    SCJoin sch = each[0];
                    SCJoinRow scjRow = new SCJoinRow();
                    scjRow.SCJ = sch;
                    //學生
                    if (GetPoint._StudentDic.ContainsKey(sch.RefStudentID))
                    {
                        scjRow.student = GetPoint._StudentDic[sch.RefStudentID];

                        //社團
                        if (GetPoint._ClubDic.ContainsKey(sch.RefClubID))
                        {
                            scjRow.club = GetPoint._ClubDic[sch.RefClubID];

                            if (ClubTeacherDic.ContainsKey(GetPoint._ClubDic[sch.RefClubID].RefTeacherID))
                            {
                                scjRow.teacher = ClubTeacherDic[GetPoint._ClubDic[sch.RefClubID].RefTeacherID];
                            }
                        }

                        if (GetPoint._RSRDic.ContainsKey(sch.UID))
                        {
                            scjRow.RSR = GetPoint._RSRDic[sch.UID];
                        }

                        RowList.Add(scjRow);
                    }
                    #endregion
                }
                else
                {
                    //沒有記錄繼續
                }
            }

            if (!string.IsNullOrEmpty(sb_3.ToString()))
            {
                MsgBox.Show(sb_3.ToString());
            }
        }

        /// <summary>
        /// 取得班級清單
        /// </summary>
        /// <returns></returns>
        private Dictionary<string, ClassRecord> GetClassDic()
        {
            Dictionary<string, ClassRecord> dic = new Dictionary<string, ClassRecord>();

            List<string> classIDList = new List<string>();
            foreach (StudentRecord sr in GetPoint._StudentDic.Values)
            {
                if (!string.IsNullOrEmpty(sr.RefClassID))
                {
                    classIDList.Add(sr.RefClassID);
                }
            }
            List<ClassRecord> classList = Class.SelectByIDs(classIDList);
            foreach (ClassRecord each in classList)
            {
                if (!dic.ContainsKey(each.ID))
                {
                    dic.Add(each.ID, each);
                }
            }
            return dic;
        }

        //畫面取得
        void BGW_FormLoad_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnReport.Enabled = true;
            btnSave.Enabled = true;
            this.Text = "成績輸入";

            if (e.Cancelled)
            {
                MsgBox.Show("資料取得已被中止");
            }
            else
            {
                if (e.Error == null)
                {
                    dataGridViewX1.AutoGenerateColumns = false;
                    RowList.Sort(SortSCJ);
                    dataGridViewX1.DataSource = RowList;

                    //Log
                    foreach (SCJoinRow each in RowList)
                    {
                        if (!_logDic.ContainsKey(each.SCJoinID))
                        {
                            _logDic.Add(each.SCJoinID, new Log_Result(each.SCJ.CopyExtension()));
                            _logDic[each.SCJoinID]._stud = each.student;
                            if (!string.IsNullOrEmpty(each.student.RefClassID))
                            {
                                if (ClassDic.ContainsKey(each.student.RefClassID))
                                {
                                    _logDic[each.SCJoinID]._class = ClassDic[each.student.RefClassID];
                                }
                            }
                        }
                    }

                    if (GetPoint._wp != null)
                    {
                        foreach (DataGridViewRow row in dataGridViewX1.Rows)
                        {
                            SetRowResults(row);
                        }
                    }
                    else
                    {
                        MsgBox.Show("尚未設定評量比例\n將無法試算出總成績資料!!");
                    }
                    _ChangeListener.Reset();
                    _ChangeListener.ResumeListen();
                    IsChangeNow = false;
                }
                else
                {
                    MsgBox.Show("發生錯誤:\n" + e.Error.Message);
                }
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            if (Save_BGW.IsBusy)
                return;

            if (!CheckDataGridValue())
            {
                MsgBox.Show("請修正資料後儲存!");
                return;
            }

            btnSave.Enabled = false;
            Save_BGW.RunWorkerAsync();
        }

        //儲存
        void Save_BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            List<SCJoin> list = new List<SCJoin>();
            foreach (SCJoinRow scj in RowList)
            {
                if (scj.HasChange)
                {
                    //Log
                    if (_logDic.ContainsKey(scj.SCJoinID))
                    {
                        Log_Result Set_Log = _logDic[scj.SCJoinID];
                        Set_Log._NewSch = scj.SCJ;
                    }

                    list.Add(scj.SCJ);
                }
            }

            _AccessHelper.UpdateValues(list);

            //Log
            FISCA.LogAgent.ApplicationLog.Log("社團", "成績輸入", GetLostConn());
        }

        private string GetLostConn()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("修改社團成績資料：");
            foreach (Log_Result each in _logDic.Values)
            {
                if (each.IsChange)
                {
                    sb.AppendLine(each.GetLogString(GetPoint));
                }
            }
            return sb.ToString();
        }

        private bool CheckDataGridValue()
        {
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.RowIndex < 0)
                        continue;

                    if (cell.ColumnIndex < 6)
                        continue;

                    if (cell.ColumnIndex > 9)
                        continue;

                    if (!CheckCellValue(cell))
                        return false;
                }
            }
            return true;

        }

        void Save_BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnSave.Enabled = true;

            if (e.Cancelled)
            {
                MsgBox.Show("儲存動作已中止!!");
            }
            else
            {
                if (e.Error == null)
                {
                    MsgBox.Show("資料儲存成功");

                    if (!BGW_FormLoad.IsBusy)
                    {
                        BGW_FormLoad.RunWorkerAsync();
                    }
                }
                else
                {
                    MsgBox.Show("發生錯誤:\n" + e.Error.Message);
                }
            }
        }

        private int SortSCJ(SCJoinRow a, SCJoinRow b)
        {
            string clubNameA = a.club.ClubNumber.PadLeft(3, '0');
            clubNameA += a.ClubName.PadLeft(10, '0');
            clubNameA += a.ClassIndex.PadLeft(3, '0');
            clubNameA += a.ClassName.PadLeft(5, '0');
            clubNameA += a.SeatNo.PadLeft(3, '0');
            clubNameA += a.StudentName.PadLeft(10, '0');

            string clubNameB = b.club.ClubNumber.PadLeft(3, '0');
            clubNameB += b.ClubName.PadLeft(10, '0');
            clubNameB += b.ClassIndex.PadLeft(3, '0');
            clubNameB += b.ClassName.PadLeft(5, '0');
            clubNameB += b.SeatNo.PadLeft(3, '0');
            clubNameB += b.StudentName.PadLeft(10, '0');

            return clubNameA.CompareTo(clubNameB);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void dataGridViewX1_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewX1.BeginEdit(true);
        }

        private bool ParseDec(string dec)
        {
            decimal decTry;
            bool a = decimal.TryParse(dec, out decTry);
            return a;
        }

        private void dataGridViewX1_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
            if (dataGridViewX1.CurrentCell.RowIndex < 0)
                return;

            if (dataGridViewX1.CurrentCell.ColumnIndex < 6)
                return;

            if (dataGridViewX1.CurrentCell.ColumnIndex > 9)
                return;

            dataGridViewX1.CurrentCell.ErrorText = "";
            dataGridViewX1.CurrentCell.Style.BackColor = Color.White;

            SCJoinRow scjRow = (SCJoinRow)dataGridViewX1.CurrentRow.DataBoundItem;
            scjRow.HasChange = true;

            if (_logDic.ContainsKey(scjRow.SCJoinID))
            {
                _logDic[scjRow.SCJoinID].IsChange = true;
            }

            if (!CheckCellValue(dataGridViewX1.CurrentCell))
            {
                return;
            }

            //進行成績及時計算
            if (GetPoint._wp != null)
            {
                SetRowResults(dataGridViewX1.CurrentRow);
            }
        }

        private void SetRowResults(DataGridViewRow row)
        {
            decimal? 平時活動比例 = ParseValue(row.Cells[Col_pa_Score.Index]);
            decimal? 出缺率比例 = ParseValue(row.Cells[Col_ar_Score.Index]);
            decimal? 活動力及服務比例 = ParseValue(row.Cells[Col_aas_Score.Index]);
            decimal? 成品成果考驗比例 = ParseValue(row.Cells[Col_far_Score.Index]);

            if (平時活動比例.HasValue)
                平時活動比例 = GetPoint._wp.PA_Weight * 平時活動比例 / 100;

            if (出缺率比例.HasValue)
                出缺率比例 = GetPoint._wp.AR_Weight * 出缺率比例 / 100;

            if (活動力及服務比例.HasValue)
                活動力及服務比例 = GetPoint._wp.AAS_Weight * 活動力及服務比例 / 100;

            if (成品成果考驗比例.HasValue)
                成品成果考驗比例 = GetPoint._wp.FAR_Weight * 成品成果考驗比例 / 100;

            decimal? results = 0;

            if (平時活動比例.HasValue)
                results += 平時活動比例.Value;

            if (出缺率比例.HasValue)
                results += 出缺率比例.Value;

            if (活動力及服務比例.HasValue)
                results += 活動力及服務比例.Value;

            if (成品成果考驗比例.HasValue)
                results += 成品成果考驗比例.Value;

            if (平時活動比例.HasValue || 出缺率比例.HasValue || 活動力及服務比例.HasValue || 成品成果考驗比例.HasValue)
            {
                row.Cells[colSResults.Index].Value = Math.Round(results.Value, MidpointRounding.AwayFromZero);
            }
            else
            {
                row.Cells[colSResults.Index].Value = "";
            }
        }

        private decimal? ParseValue(DataGridViewCell cell)
        {
            if (!string.IsNullOrEmpty("" + cell.Value))
            {
                decimal dc = 0;
                decimal.TryParse("" + cell.Value, out dc);
                return dc;
            }
            else
            {
                return null;
            }

        }

        private bool CheckCellValue(DataGridViewCell cell)
        {

            if (string.IsNullOrEmpty("" + cell.Value))
            {
                return true;
            }

            if (!ParseDec("" + cell.Value))
            {
                cell.ErrorText = "必須是數字";
                cell.Style.BackColor = Color.Red;

                //輸入的是數字如果未繼續編輯將會出現只能輸入單位數的狀況
                //dataGridViewX1.BeginEdit(false);
                return false;
            }

            //大於100黃色字
            decimal de = decimal.Parse("" + cell.Value);
            if (de > 100)
            {
                cell.Style.BackColor = Color.Yellow;
                //dataGridViewX1.BeginEdit(false);
                return true;
            }

            //有小數點黃色字
            if (de.ToString().Contains('.'))
            {
                cell.Style.BackColor = Color.Yellow;
                return true;
            }

            return true;
        }

        private void ClubResultsInput_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (IsChangeNow)
            {
                DialogResult dr = FISCA.Presentation.Controls.MsgBox.Show("確認放棄?", "尚未儲存資料", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dr != System.Windows.Forms.DialogResult.Yes)
                {
                    e.Cancel = true;
                }
            }
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            #region 匯出
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.FileName = "匯出成績輸入";
            saveFileDialog1.Filter = "Excel (*.xls)|*.xls";
            if (saveFileDialog1.ShowDialog() != DialogResult.OK) return;

            DataGridViewExport export = new DataGridViewExport(dataGridViewX1);
            export.Save(saveFileDialog1.FileName);

            if (new CompleteForm().ShowDialog() == DialogResult.Yes)
                System.Diagnostics.Process.Start(saveFileDialog1.FileName);
            #endregion
        }
    }
}
