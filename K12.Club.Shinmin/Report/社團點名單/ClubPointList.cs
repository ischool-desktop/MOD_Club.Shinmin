using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;
using System.ComponentModel;
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
    //社團點名單
    class AssociationsPointList
    {
        //背景模式
        private BackgroundWorker BGW = new BackgroundWorker();

        AccessHelper _AccessHelper = new AccessHelper();

        SCJoinDataLoad SDL { get; set; }

        public AssociationsPointList()
        {
            BGW.DoWork += new DoWorkEventHandler(BGW_DoWork);
            BGW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(BGW_RunWorkerCompleted);
            BGW.RunWorkerAsync();
        }

        void BGW_DoWork(object sender, DoWorkEventArgs e)
        {
            //建立社團比對學生ID的清單
            //以每個社團為一個單位
            //進行學生資料列印
            SDL = new SCJoinDataLoad();

            //建立Excel範本
            Workbook template = new Workbook();
            template.Open(new MemoryStream(Properties.Resources.社團點名單_範本), FileFormatType.Excel2003);

            //每一張
            Workbook prototype = new Workbook();
            prototype.Copy(template);
            Worksheet ptws = prototype.Worksheets[0];

            //範圍
            Range ptHeader = ptws.Cells.CreateRange(0, 4, false);
            Range ptEachRow = ptws.Cells.CreateRange(4, 1, false);

            //儲存資料
            Workbook wb = new Workbook();
            wb.Copy(prototype);
            //取得Sheet
            Worksheet ws = wb.Worksheets[0];

            int index = 0;
            int dataIndex = 0;

            //每一個社團
            foreach (string club in SDL.ClubByStudentList.Keys)
            {
                //社團資訊收集

                CLUBRecord cr = SDL.CLUBRecordDic[club];

                //社團標頭
                string TitleName1 = string.Format("{0}學年度／第{1}學期　社團點名單", cr.SchoolYear.ToString(), cr.Semester.ToString());
                string TitleName2 = cr.ClubName + "　(類型：" + cr.ClubCategory + ")";
                ws.Cells.CreateRange(dataIndex, 4, false).Copy(ptHeader);
                ws.Cells[dataIndex, 0].PutValue(TitleName1);
                dataIndex += 1;
                ws.Cells[dataIndex, 0].PutValue(TitleName2);
                dataIndex += 1;
                ws.Cells[dataIndex, 0].PutValue("代碼：" + cr.ClubNumber);
                ws.Cells[dataIndex, 1].PutValue("場地：" + cr.Location);
                ws.Cells[dataIndex, 3].PutValue("人數：" + SDL.ClubByStudentList[club].Count);
                //社團老師
                string TeacherNameA = GetTeacherName(cr.RefTeacherID);
                string TeacherNameB = GetTeacherName(cr.RefTeacherID2);
                string TeacherNameC = GetTeacherName(cr.RefTeacherID3);
                string TeacherName = "老師：" + TeacherNameA + "　" + TeacherNameB + "　" + TeacherNameC;

                ws.Cells[dataIndex, 4].PutValue(TeacherName);
                dataIndex += 2;

                foreach (StudentRecord stud in SDL.ClubByStudentList[club])
                {
                    ws.Cells.CreateRange(dataIndex, 1, false).Copy(ptEachRow);

                    string classname = string.IsNullOrEmpty(stud.RefClassID) ? "" : stud.Class.Name;
                    ws.Cells[dataIndex, 0].PutValue(classname);
                    ws.Cells[dataIndex, 1].PutValue(stud.SeatNo.HasValue ? stud.SeatNo.Value.ToString() : "");
                    ws.Cells[dataIndex, 2].PutValue(stud.Name);
                    ws.Cells[dataIndex, 3].PutValue(stud.StudentNumber);
                    ws.Cells[dataIndex, 4].PutValue(stud.Gender);
                    dataIndex += 1;
                }

                ws.Cells.CreateRange(dataIndex - 1, 0, 1, 6).SetOutlineBorder(BorderType.BottomBorder, CellBorderType.Medium, Color.Black);
                ws.HPageBreaks.Add(dataIndex, 6);
            }

            e.Result = wb;
        }

        private string GetTeacherName(string p)
        {
            string name = "";
            if (!string.IsNullOrEmpty(p))
            {
                if (SDL.TeacherDic.ContainsKey(p))
                {
                    TeacherRecord tr = SDL.TeacherDic[p];
                    if (!string.IsNullOrEmpty(tr.Nickname))
                    {
                        name = tr.Name + "(" + tr.Nickname + ")";
                    }
                    else
                    {
                        name = tr.Name;
                    }
                }
            }
            return name;
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
                    SaveFileDialog1.FileName = "社團點名單";

                    //資料
                    try
                    {
                        if (SaveFileDialog1.ShowDialog() == DialogResult.OK)
                        {
                            Workbook inResult = (Workbook)e.Result;
                            inResult.Save(SaveFileDialog1.FileName);
                            Process.Start(SaveFileDialog1.FileName);
                            MotherForm.SetStatusBarMessage("社團點名單,列印完成!!");
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
    }
}
