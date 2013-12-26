using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using FISCA.UDT;
using K12.Data;
using Aspose.Cells;
using System.IO;
using FISCA.Presentation.Controls;
using System.Windows.Forms;
using System.Diagnostics;
using FISCA.Presentation;
using System.Drawing;

namespace K12.Club.Shinmin
{
    class ClubTranscript
    {
        private BackgroundWorker BGW = new BackgroundWorker();

        ClubTraMag GetPoint { get; set; }

        WeightProportion _wp { get; set; }

        public ClubTranscript()
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //1.選擇社團
            //2.取得社團學生
            List<WeightProportion> wpList = tool._A.Select<WeightProportion>();
            if (wpList.Count >= 0)
            {
                _wp = wpList[0];
            }

            //取得相關的學生資料
            //1.基本資料
            //2.社團結算成績
            GetPoint = new ClubTraMag();

            //取得範本
            #region 建立範本

            Workbook template = new Workbook();
            template.Open(new MemoryStream(Properties.Resources.社團成績單_範本), FileFormatType.Excel2003);

            //每一張
            Workbook prototype = new Workbook();
            prototype.Copy(template);
            Worksheet ptws = prototype.Worksheets[0];

            //範圍
            Range ptHeader = ptws.Cells.CreateRange(0, 3, false);
            Range ptEachRow = ptws.Cells.CreateRange(3, 1, false);

            //儲存資料
            Workbook wb = new Workbook();
            wb.Copy(prototype);

            //取得Sheet
            Worksheet ws = wb.Worksheets[0];

            //int index = 0;
            int dataIndex = 0;
            int CountPage = 1;

            #endregion

            #region 填資料

            foreach (string clubID in GetPoint.TraDic.Keys)
            {
                ws.Cells.CreateRange(dataIndex, 3, false).Copy(ptHeader);
                //每一個社團
                CLUBRecord cr = GetPoint.CLUBDic[clubID];

                //標題
                string TitleName = string.Format("{0}學年度　第{1}學期　{2}　社團成績單", cr.SchoolYear.ToString(), cr.Semester.ToString(), cr.ClubName);
                ws.Cells[dataIndex, 0].PutValue(TitleName);
                dataIndex++;

                //代碼
                ws.Cells[dataIndex, 0].PutValue("代碼：" + cr.ClubNumber);

                //教師
                string TeacherString = "教師：" + GetTeacherName(cr);
                ws.Cells[dataIndex, 2].PutValue(TeacherString);

                //頁數
                ws.Cells[dataIndex, 7].PutValue("日期：" + DateTime.Now.ToString("yyyy/MM/dd HH:mm") + "　頁數:" + CountPage.ToString());
                CountPage++; //每班增加1頁
                dataIndex += 1;

                //處理比例訊息
                if (_wp != null)
                {
                    ws.Cells[dataIndex, 5].PutValue("平時活動(" + _wp.PA_Weight + "%)");
                    ws.Cells[dataIndex, 6].PutValue("出缺率(" + _wp.AR_Weight + "%)");
                    ws.Cells[dataIndex, 7].PutValue("活動力及服務(" + _wp.AAS_Weight + "%)");
                    ws.Cells[dataIndex, 8].PutValue("成品成果考驗(" + _wp.FAR_Weight + "%)");
                }

                dataIndex += 1;

                GetPoint.TraDic[clubID].Sort(SortTraDic);

                foreach (ClubTraObj each in GetPoint.TraDic[clubID])
                {
                    ws.Cells.CreateRange(dataIndex, 1, false).Copy(ptEachRow);

                    //基本資料
                    ws.Cells[dataIndex, 0].PutValue(each.student.Class != null ? each.student.Class.Name : "");
                    ws.Cells[dataIndex, 1].PutValue(each.student.SeatNo.HasValue ? each.student.SeatNo.Value.ToString() : "");
                    ws.Cells[dataIndex, 2].PutValue(each.student.Name);
                    ws.Cells[dataIndex, 3].PutValue(each.student.StudentNumber);
                    ws.Cells[dataIndex, 4].PutValue(each.student.Gender);

                    //成績(依據比例)
                    ws.Cells[dataIndex, 5].PutValue(each.SCJ.PAScore.HasValue ? tool.GetDecimalValue(each.SCJ.PAScore.Value, _wp.PA_Weight) : "");
                    ws.Cells[dataIndex, 6].PutValue(each.SCJ.ARScore.HasValue ? tool.GetDecimalValue(each.SCJ.ARScore.Value, _wp.AR_Weight) : "");
                    ws.Cells[dataIndex, 7].PutValue(each.SCJ.AASScore.HasValue ? tool.GetDecimalValue(each.SCJ.AASScore.Value, _wp.AAS_Weight) : "");
                    ws.Cells[dataIndex, 8].PutValue(each.SCJ.FARScore.HasValue ? tool.GetDecimalValue(each.SCJ.FARScore.Value, _wp.FAR_Weight) : "");

                    //學期成績
                    if (each.RSR != null)
                        ws.Cells[dataIndex, 9].PutValue(each.RSR.ResultScore.HasValue ? each.RSR.ResultScore.Value.ToString() : "");

                    dataIndex++;
                }

                //Range absenceStatisticsRange = ws.Cells.CreateRange(dataIndex, 0, 1, 10);
                //absenceStatisticsRange.SetOutlineBorder(BorderType.TopBorder, CellBorderType.Medium, Color.Black);
                ws.HPageBreaks.Add(dataIndex, 9);
            }

            #endregion

            e.Result = wb;

        }

        public int SortTraDic(ClubTraObj a1, ClubTraObj b1)
        {
            string StudentA = "";
            if (a1.student.Class != null)
            {
                StudentA += a1.student.Class.DisplayOrder.PadLeft(10, '0');
                StudentA += a1.student.Class.Name.PadLeft(10, '0');
            }
            else
            {
                StudentA += "00000000000000000000";
            }

            StudentA += a1.student.SeatNo.HasValue ? a1.student.SeatNo.Value.ToString().PadLeft(10, '0') : "0000000000";
            StudentA += a1.student.Name.PadLeft(10, '0');

            string StudentB = "";
            if (b1.student.Class != null)
            {
                StudentB += b1.student.Class.DisplayOrder.PadLeft(10, '0');
                StudentB += b1.student.Class.Name.PadLeft(10, '0');
            }
            else
            {
                StudentB += "00000000000000000000";
            }
            StudentB += b1.student.SeatNo.HasValue ? b1.student.SeatNo.Value.ToString().PadLeft(10, '0') : "0000000000";
            StudentB += b1.student.Name.PadLeft(10, '0');

            return StudentA.CompareTo(StudentB);
        }

        void BGW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Cancelled)
            {
                MsgBox.Show("作業已被中止!!");
            }
            else
            {
                if (e.Error == null)
                {
                    SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                    SaveFileDialog1.Filter = "Excel (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                    SaveFileDialog1.FileName = "社團成績單";

                    //資料
                    try
                    {
                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            Workbook inResult = (Workbook)e.Result;
                            inResult.Save(SaveFileDialog1.FileName);
                            Process.Start(SaveFileDialog1.FileName);
                            MotherForm.SetStatusBarMessage("社團成績單,列印完成!!");
                        }
                        else
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("檔案未儲存");
                            return;
                        }
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                        MotherForm.SetStatusBarMessage("檔案儲存錯誤,請檢查檔案是否開啟中!!");
                    }
                }
                else
                {
                    MsgBox.Show("列印資料發生錯誤\n" + e.Error.Message);
                    SmartSchool.ErrorReporting.ReportingService.ReportException(e.Error);
                }
            }
        }

        /// <summary>
        /// 取得老師名稱
        /// </summary>
        private string GetTeacherName(CLUBRecord cr)
        {
            string name = "";
            //老師1
            if (!string.IsNullOrEmpty(cr.RefTeacherID))
            {
                TeacherRecord tr = GetPoint.TeacherDic[cr.RefTeacherID];
                if (string.IsNullOrEmpty(tr.Nickname))
                {
                    name += tr.Name;
                }
                else
                {
                    name += tr.Name + "(" + tr.Nickname + ")";
                }
            }
            //老師2
            if (!string.IsNullOrEmpty(cr.RefTeacherID2))
            {
                TeacherRecord tr = GetPoint.TeacherDic[cr.RefTeacherID2];
                if (string.IsNullOrEmpty(tr.Nickname))
                {
                    name += "／" + tr.Name;
                }
                else
                {
                    name += "／" + tr.Name + "(" + tr.Nickname + ")";
                }
            }
            //老師3
            if (!string.IsNullOrEmpty(cr.RefTeacherID3))
            {
                TeacherRecord tr = GetPoint.TeacherDic[cr.RefTeacherID3];
                if (string.IsNullOrEmpty(tr.Nickname))
                {
                    name += "／" + tr.Name;
                }
                else
                {
                    name += "／" + tr.Name + "(" + tr.Nickname + ")";
                }
            }
            return name;
        }
    }
}
