using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using FISCA.UDT;

namespace K12.Club.Shinmin
{
    public partial class ClearingForm : BaseForm
    {

        /// <summary>
        /// UDT資料取得器
        /// </summary>
        private AccessHelper _AccessHelper { get; set; }

        BackgroundWorker BGW = new BackgroundWorker();

        List<ResultScoreRecord> InsertScoreList { get; set; }

        List<ResultScoreRecord> UPDateScoreList { get; set; }

        public ClearingForm()
        {
            InitializeComponent();
        }

        private void ClearingForm_Load(object sender, EventArgs e)
        {
            _AccessHelper = new AccessHelper();

            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            if (!BGW.IsBusy)
            {
                btnStart.Enabled = false;
                this.Text = "學期結算(系統結算中...)";
                BGW.RunWorkerAsync();
            }
            else
            {
                MsgBox.Show("系統忙碌中,稍後再試...");
            }
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //取得目前選擇課程
            成績取得器 tool = new 成績取得器();
            tool.GetSCJoinByClubIDList(ClubAdmin.Instance.SelectedSource);

            //取得運算比例
            tool.SetWeightProportion();

            //社團ID : 社團幹部obj
            Dictionary<string, 社團幹部obj> CadreDic = new Dictionary<string, 社團幹部obj>();

            #region 處理幹部記錄

            foreach (CLUBRecord each in tool._ClubDic.Values)
            {
                if (!CadreDic.ContainsKey(each.UID))
                {
                    CadreDic.Add(each.UID, new 社團幹部obj(each));
                }
            }
            string qr = string.Join("','", tool._ClubDic.Keys);
            List<CadresRecord> list = _AccessHelper.Select<CadresRecord>("ref_club_id in ('" + qr + "')");
            foreach (CadresRecord cr in list)
            {
                if (!CadreDic[cr.RefClubID]._Cadre1.ContainsKey(cr.RefStudentID))
                {
                    CadreDic[cr.RefClubID]._Cadre1.Add(cr.RefStudentID, cr.CadreName);
                }
                else
                {
                    CadreDic[cr.RefClubID]._Cadre1[cr.RefStudentID] += "," + cr.CadreName;
                }
            }

            #endregion


            Dictionary<string, ResultScoreRecord> ResultScoreDic = new Dictionary<string, ResultScoreRecord>();

            List<string> list_2 = new List<string>();
            foreach (List<SCJoin> each in tool._SCJoinDic.Values)
            {
                foreach (SCJoin scj in each)
                {
                    list_2.Add(scj.UID);
                }
            }
            string uq = string.Join("','", list_2);

            List<ResultScoreRecord> ResultList = _AccessHelper.Select<ResultScoreRecord>("ref_scjoin_id in ('" + uq + "')");
            foreach (ResultScoreRecord rsr in ResultList)
            {
                if (!ResultScoreDic.ContainsKey(rsr.RefSCJoinID))
                {
                    ResultScoreDic.Add(rsr.RefSCJoinID, rsr);
                }
            }

            UPDateScoreList = new List<ResultScoreRecord>();
            InsertScoreList = new List<ResultScoreRecord>();

            //_AccessHelper.DeletedValues(ResultList);

            foreach (List<SCJoin> scjList in tool._SCJoinDic.Values)
            {
                foreach (SCJoin scj in scjList)
                {
                    if (ResultScoreDic.ContainsKey(scj.UID))
                    {
                        #region 如果有原資料
                        if (tool._StudentDic.ContainsKey(scj.RefStudentID))
                        {
                            //社團
                            CLUBRecord cr = tool._ClubDic[scj.RefClubID];
                            //學生
                            StudentRecord sr = tool._StudentDic[scj.RefStudentID];
                            //原有社團成績記錄
                            ResultScoreRecord update_rsr = ResultScoreDic[scj.UID];

                            update_rsr.SchoolYear = cr.SchoolYear;
                            update_rsr.Semester = cr.Semester;

                            update_rsr.RefClubID = cr.UID; //社團ID
                            update_rsr.RefStudentID = sr.ID; //學生ID
                            update_rsr.RefSCJoinID = scj.UID; //參與記錄ID

                            update_rsr.ClubName = cr.ClubName;

                            //成績
                            update_rsr.ResultScore = tool.GetDecimalValue(scj); //成績

                            #region 幹部
                            if (CadreDic.ContainsKey(cr.UID))
                            {
                                if (CadreDic[cr.UID]._Cadre1.ContainsKey(sr.ID))
                                {
                                    update_rsr.CadreName = CadreDic[cr.UID]._Cadre1[sr.ID]; //幹部
                                }
                                else
                                {
                                    update_rsr.CadreName = "";
                                }
                            }
                            else
                            {
                                update_rsr.CadreName = "";
                            }
                            #endregion

                            UPDateScoreList.Add(update_rsr);
                        } 
                        #endregion
                    }
                    else
                    {
                        #region 完全沒有成績記錄
                        if (tool._StudentDic.ContainsKey(scj.RefStudentID))
                        {
                            //社團
                            CLUBRecord cr = tool._ClubDic[scj.RefClubID];
                            //學生
                            StudentRecord sr = tool._StudentDic[scj.RefStudentID];

                            ResultScoreRecord rsr = new ResultScoreRecord();
                            rsr.SchoolYear = cr.SchoolYear;
                            rsr.Semester = cr.Semester;

                            rsr.RefClubID = cr.UID; //社團ID
                            rsr.RefStudentID = sr.ID; //學生ID
                            rsr.RefSCJoinID = scj.UID; //參與記錄ID

                            rsr.ClubName = cr.ClubName;

                            bool check = false;

                            if (CheckNull(scj))
                            {
                                check = true;
                                rsr.ResultScore = tool.GetDecimalValue(scj); //成績
                            }

                            if (CadreDic.ContainsKey(cr.UID))
                            {
                                if (CadreDic[cr.UID]._Cadre1.ContainsKey(sr.ID))
                                {
                                    check = true;
                                    rsr.CadreName = CadreDic[cr.UID]._Cadre1[sr.ID]; //幹部
                                }
                                else
                                {
                                    rsr.CadreName = "";
                                }
                            }

                            if (check)
                            {
                                InsertScoreList.Add(rsr);
                            }
                        } 
                        #endregion
                    }
                }
            }

            try
            {
                _AccessHelper.InsertValues(InsertScoreList);
                _AccessHelper.UpdateValues(UPDateScoreList);
            }
            catch (Exception ex)
            {
                MsgBox.Show("新增社團成績發生錯誤!!\n" + ex.Message);
                e.Cancel = true;
                return;
            }

            #region 社團成績Log處理
            StringBuilder _sbLog = new StringBuilder();
            _sbLog.AppendLine("已進行社團結算：");
            if (InsertScoreList.Count > 0)
                _sbLog.AppendLine("共新增「" + InsertScoreList.Count + "」筆成績記錄");
            if (UPDateScoreList.Count > 0)
                _sbLog.AppendLine("共更新「" + UPDateScoreList.Count + "」筆成績記錄");
            _sbLog.AppendLine("");
            _sbLog.AppendLine("簡要明細如下：");
            if (InsertScoreList.Count > 0)
            {
                foreach (ResultScoreRecord each in InsertScoreList)
                {
                    if (tool._StudentDic.ContainsKey(each.RefStudentID))
                    {
                        if (string.IsNullOrEmpty(each.CadreName))
                        {
                            StudentRecord sr = tool._StudentDic[each.RefStudentID];
                            string de = each.ResultScore.HasValue ? each.ResultScore.Value.ToString() : "";
                            _sbLog.AppendLine(string.Format("學生新增「{0}」社團成績「{1}」", sr.Name, de));
                        }
                        else
                        {
                            StudentRecord sr = tool._StudentDic[each.RefStudentID];
                            string de = each.ResultScore.HasValue ? each.ResultScore.Value.ToString() : "";
                            _sbLog.AppendLine(string.Format("學生新增「{0}」社團成績「{1}」幹部記錄「{2}」", sr.Name, de, each.CadreName));
                        }
                    }
                }
            }

            if (UPDateScoreList.Count > 0)
            {
                _sbLog.AppendLine("");
                foreach (ResultScoreRecord each in UPDateScoreList)
                {
                    if (tool._StudentDic.ContainsKey(each.RefStudentID))
                    {
                        if (string.IsNullOrEmpty(each.CadreName))
                        {
                            StudentRecord sr = tool._StudentDic[each.RefStudentID];
                            string de = each.ResultScore.HasValue ? each.ResultScore.Value.ToString() : "";
                            _sbLog.AppendLine(string.Format("學生更新「{0}」社團成績「{1}」", sr.Name, de));
                        }
                        else
                        {
                            StudentRecord sr = tool._StudentDic[each.RefStudentID];
                            string de = each.ResultScore.HasValue ? each.ResultScore.Value.ToString() : "";
                            _sbLog.AppendLine(string.Format("學生更新「{0}」社團成績「{1}」幹部記錄「{2}」", sr.Name, de, each.CadreName));
                        }
                    }

                }
            }

            if (InsertScoreList.Count + UPDateScoreList.Count > 0)
            {
                try
                {
                    FISCA.LogAgent.ApplicationLog.Log("社團", "成績結算", _sbLog.ToString());
                }
                catch (Exception ex)
                {
                    MsgBox.Show("上傳Log記錄發生錯誤!!\n" + ex.Message);
                    e.Cancel = true;
                    return;
                }
            }

            #endregion
        }

        private bool CheckNull(SCJoin scj)
        {
            if (scj.AASScore.HasValue || scj.ARScore.HasValue || scj.FARScore.HasValue || scj.PAScore.HasValue)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.Text = "學期結算";
            btnStart.Enabled = true;

            if (e.Cancelled)
            {
                FISCA.Presentation.Controls.MsgBox.Show("結算操作已因錯誤中止!!");
            }
            else
            {
                if (e.Error == null)
                {
                    if (InsertScoreList.Count + UPDateScoreList.Count > 0)
                    {
                        StringBuilder sb_8 = new StringBuilder();
                        sb_8.AppendLine("結算完成!!");
                        if (InsertScoreList.Count != 0)
                            sb_8.AppendLine("新增[" + InsertScoreList.Count + "]筆");
                        if (UPDateScoreList.Count != 0)
                            sb_8.AppendLine("更新[" + UPDateScoreList.Count + "]筆");
                        FISCA.Presentation.Controls.MsgBox.Show(sb_8.ToString());

                        this.Close();
                    }
                    else
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("結算失敗!!無成績記錄可供結算!!");
                    }
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("結算發生錯誤!!\n" + e.Error.Message);
                }
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
